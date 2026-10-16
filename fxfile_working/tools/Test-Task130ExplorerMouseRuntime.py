"""Exercise consent in an isolated Task130 GUI fixture; never targets user data."""
import ctypes as C
from ctypes import wintypes as W
import hashlib, json, pathlib, subprocess, sys, time
from PIL import Image

arch = sys.argv[1]
suffix = arch + ('_external' if '--external' in sys.argv else '')
if '--tag' in sys.argv:
    tag=sys.argv[sys.argv.index('--tag')+1]
    assert tag and all(c.islower() or c.isdigit() or c=='_' for c in tag)
    suffix+='_'+tag
workspace = pathlib.Path(__file__).resolve().parents[2]
evidence = workspace / '__BUILD_TEMP_BACKUP__' / 'task130_results_20261015'
stage = evidence / ('gui_' + suffix)
fixture = evidence / ('fixture_' + suffix)
assert stage.resolve().is_relative_to(evidence.resolve()) and fixture.resolve().is_relative_to(evidence.resolve())
assert stage.is_dir() and fixture.is_dir()
u, k, g = C.windll.user32, C.windll.kernel32, C.windll.gdi32
P = C.c_void_p
for name in ['GetDlgItem','GetParent','GetAncestor','GetFocus','GetForegroundWindow','WindowFromPoint']:
    getattr(u,name).restype=P
u.SendMessageW.argtypes=[P,W.UINT,C.c_size_t,C.c_ssize_t];u.SendMessageW.restype=C.c_ssize_t
u.PostMessageW.argtypes=[P,W.UINT,C.c_size_t,C.c_ssize_t]
k.OpenProcess.argtypes=[W.DWORD,W.BOOL,W.DWORD];k.OpenProcess.restype=P
k.VirtualAllocEx.argtypes=[P,P,C.c_size_t,W.DWORD,W.DWORD];k.VirtualAllocEx.restype=P
k.WriteProcessMemory.argtypes=[P,P,P,C.c_size_t,P]
k.ReadProcessMemory.argtypes=[P,P,P,C.c_size_t,P]
k.VirtualFreeEx.argtypes=[P,P,C.c_size_t,W.DWORD]
k.CloseHandle.argtypes=[P]
k.GlobalAlloc.argtypes=[W.UINT,C.c_size_t];k.GlobalAlloc.restype=P
k.GlobalLock.argtypes=[P];k.GlobalLock.restype=P;k.GlobalUnlock.argtypes=[P]
u.OpenClipboard.argtypes=[P];u.SetClipboardData.argtypes=[W.UINT,P];u.SetClipboardData.restype=P
for name in ['CreateCompatibleDC','CreateCompatibleBitmap','SelectObject']:
    getattr(g,name).restype=P
g.CreateCompatibleDC.argtypes=[P];g.CreateCompatibleBitmap.argtypes=[P,C.c_int,C.c_int]
g.SelectObject.argtypes=[P,P];g.DeleteObject.argtypes=[P];g.DeleteDC.argtypes=[P]
g.GetObjectW.argtypes=[P,C.c_int,P]
u.GetWindowRect.argtypes=[P,P];u.PrintWindow.argtypes=[P,P,W.UINT]
u.GetDC.argtypes=[P];u.GetDC.restype=P;u.ReleaseDC.argtypes=[P,P]
u.IsWindow.argtypes=[P];u.IsClipboardFormatAvailable.argtypes=[W.UINT]
u.IsWindowEnabled.argtypes=[P]
u.GetWindowTextW.argtypes=[P,W.LPWSTR,C.c_int]
u.GetDlgCtrlID.argtypes=[P]
u.GetClassNameW.argtypes=[P,W.LPWSTR,C.c_int]
u.GetWindowThreadProcessId.argtypes=[P,P]
u.ShowWindow.argtypes=[P,C.c_int];u.SetForegroundWindow.argtypes=[P]
u.SetFocus.argtypes=[P];u.SetWindowPos.argtypes=[P,P,C.c_int,C.c_int,C.c_int,C.c_int,W.UINT]
u.IsWindowVisible.argtypes=[P];u.GetDlgItem.argtypes=[P,C.c_int]
CALLBACK=C.WINFUNCTYPE(W.BOOL,P,P)
u.EnumWindows.argtypes=[CALLBACK,P];u.EnumChildWindows.argtypes=[P,CALLBACK,P]

def wait_for(fn, seconds=30):
    deadline=time.monotonic()+seconds
    while time.monotonic()<deadline:
        result=fn()
        if result: return result
        time.sleep(.15)
    raise AssertionError('Timed out: '+str(fn))

def windows(parent=None):
    result=[]
    @CALLBACK
    def visit(hwnd,unused):
        result.append(hwnd);return True
    if parent is None: u.EnumWindows(visit,None)
    else: u.EnumChildWindows(parent,visit,None)
    return result

def pid_of(hwnd):
    pid=W.DWORD();u.GetWindowThreadProcessId(hwnd,C.byref(pid));return pid.value

def cls(hwnd):
    value=C.create_unicode_buffer(256);u.GetClassNameW(hwnd,value,256);return value.value

def text(hwnd):
    buf=C.create_unicode_buffer(200000)
    u.SendMessageW(hwnd,0xD,len(buf),C.addressof(buf));return buf.value

def caption(hwnd):
    value=C.create_unicode_buffer(2048);u.GetWindowTextW(hwnd,value,len(value));return value.value

def rect(hwnd):
    r=W.RECT();assert u.GetWindowRect(hwnd,C.byref(r));return r

def snap(folder):
    result={}
    for p in folder.rglob('*'):
        if p.is_file():
            try:result[str(p.relative_to(folder))]=hashlib.sha256(p.read_bytes()).hexdigest()
            except (FileNotFoundError, PermissionError):pass  # Retry the full expected hash set while shell copy holds an exclusive handle.
    return result

def screenshot(hwnd,path):
    r=rect(hwnd);width,height=r.right-r.left,r.bottom-r.top
    screen=u.GetDC(None);dc=g.CreateCompatibleDC(screen);bitmap=g.CreateCompatibleBitmap(screen,width,height)
    old=g.SelectObject(dc,bitmap)
    assert u.PrintWindow(hwnd,dc,2)
    g.SelectObject(dc,old)
    class HEADER(C.Structure):
        _fields_=[('size',W.DWORD),('width',W.LONG),('height',W.LONG),('planes',W.WORD),('bits',W.WORD),('compression',W.DWORD),('image',W.DWORD),('xp',W.LONG),('yp',W.LONG),('used',W.DWORD),('important',W.DWORD)]
    header=HEADER(40,width,-height,1,32,0,0,0,0,0,0)
    data=C.create_string_buffer(width*height*4)
    g.GetDIBits.argtypes=[P,P,W.UINT,W.UINT,P,P,W.UINT]
    assert g.GetDIBits(dc,bitmap,0,height,data,C.byref(header),0)
    Image.frombytes('RGB',(width,height),data.raw,'raw','BGRX').save(path)
    g.SelectObject(dc,old);g.DeleteObject(bitmap);g.DeleteDC(dc);u.ReleaseDC(None,screen)

startup=subprocess.STARTUPINFO();startup.dwFlags|=subprocess.STARTF_USESHOWWINDOW;startup.wShowWindow=0
proc=subprocess.Popen([str(stage/'fxfile.exe')],cwd=stage,startupinfo=startup)
handle=remote=None;records=[]
try:
    def main_window():
        for h in windows():
            if pid_of(h)==proc.pid and 'fxfile' in text(h).lower() and cls(h)!='#32770':return h
    main=wait_for(main_window,60);u.ShowWindow(main,1)
    lists=wait_for(lambda:[h for h in windows(main) if cls(h)=='SysListView32' and u.IsWindowVisible(h)] if
                   len([h for h in windows(main) if cls(h)=='SysListView32' and u.IsWindowVisible(h)])==6 else None,40)
    lists.sort(key=lambda h:(rect(h).top,rect(h).left))
    time.sleep(3)
    handle=k.OpenProcess(0x38,False,proc.pid);assert handle
    remote=k.VirtualAllocEx(handle,None,8192,0x3000,4);assert remote
    Pointer=C.c_uint64 if arch=='x64' else C.c_uint32
    class ITEM(C.Structure):
        _fields_=[('mask',W.UINT),('item',C.c_int),('sub',C.c_int),('state',W.UINT),('stateMask',W.UINT),('psz',Pointer),('cch',C.c_int),('image',C.c_int),('param',Pointer)]
    def put(value):
        assert k.WriteProcessMemory(handle,remote,C.byref(value),C.sizeof(value),None)
    def item_names(view):
        count=u.SendMessageW(view,0x1004,0,0);names=[]
        for index in range(count):
            item=ITEM(1,index,0,0,0,remote+1024,1024,0,0);put(item)
            u.SendMessageW(view,0x1073,index,remote)
            buf=C.create_unicode_buffer(1024);assert k.ReadProcessMemory(handle,remote+1024,buf,C.sizeof(buf),None)
            names.append(buf.value)
        return names
    def focus(hwnd):
        tid=u.GetWindowThreadProcessId(hwnd,None);current=k.GetCurrentThreadId()
        u.AttachThreadInput(current,tid,True)
        try:
            u.SetWindowPos(main,P(-1),0,0,0,0,0x13);u.SetForegroundWindow(main);u.SetFocus(hwnd)
        finally:u.AttachThreadInput(current,tid,False)
    def select(view,names):
        wait_for(lambda:all(name in item_names(view) for name in names))
        item=ITEM(8,0,0,0,3,0,0,0,0);put(item);u.SendMessageW(view,0x102B,C.c_size_t(-1).value,remote)
        actual=item_names(view)
        for name in names:
            item=ITEM(8,0,0,3,3,0,0,0,0);put(item);u.SendMessageW(view,0x102B,actual.index(name),remote)
        focus(view)
        assert u.SendMessageW(view,0x1032,0,0)==len(names)
    def command(view,command):
        focus(view);assert u.PostMessageW(main,0x111,command,0)
    def cut_states(view):
        return [u.SendMessageW(view,0x102C,index,4) for index,name in enumerate(item_names(view)) if name in ['alpha.txt','beta.txt','한글 폴더']]
    def dialogs():
        return [h for h in windows() if pid_of(h)==proc.pid and cls(h)=='#32770' and u.IsWindowVisible(h)]
    def consent(name,target=None,source_count=None):
        dialog=wait_for(lambda:next((h for h in dialogs() if u.GetDlgItem(h,1981)),None))
        summary=text(u.GetDlgItem(dialog,1980));details=text(u.GetDlgItem(dialog,1981))
        default=u.SendMessageW(dialog,0x400,0,0)&0xffff
        assert default==2,(name,default)
        class FONT(C.Structure):
            _fields_=[('height',W.LONG),('width',W.LONG),('escape',W.LONG),('orient',W.LONG),('weight',W.LONG),
                ('italic',W.BYTE),('underline',W.BYTE),('strike',W.BYTE),('charset',W.BYTE),('out',W.BYTE),
                ('clip',W.BYTE),('quality',W.BYTE),('pitch',W.BYTE),('face',W.WCHAR*32)]
        fonts=[];outer=rect(dialog)
        for identifier in [1980,1981,1,2]:
            control=u.GetDlgItem(dialog,identifier);font=FONT()
            assert g.GetObjectW(u.SendMessageW(control,0x31,0,0),C.sizeof(font),C.byref(font))
            assert abs(font.height)>=16,(name,identifier,font.height)
            bounds=rect(control)
            assert outer.left<=bounds.left<bounds.right<=outer.right and outer.top<=bounds.top<bounds.bottom<=outer.bottom
            fonts.append({'Control':identifier,'HeightPixels':font.height,'Face':font.face})
        if target:assert str(target) in details,(name,details)
        if source_count:assert str(source_count)+'개 선택 항목' in summary,(name,summary)
        screenshot(dialog,evidence/(suffix+'_'+name+'.png'))
        records.append({'Case':name,'Summary':summary,'Details':details,'DefaultButton':default,'Fonts':fonts})
        print('DIALOG '+name,flush=True)
        return dialog
    def respond_mouse(dialog,how):
        u.SetForegroundWindow(dialog)
        if how in ('confirm','cancel'):
            button=u.GetDlgItem(dialog,1 if how=='confirm' else 2);r=rect(button)
            u.SetCursorPos((r.left+r.right)//2,(r.top+r.bottom)//2)
            time.sleep(.12);u.mouse_event(2,0,0,0,0);time.sleep(.1);u.mouse_event(4,0,0,0,0)
        elif how in ('esc','enter'):
            vk=27 if how=='esc' else 13
            u.keybd_event(vk,0,0,0);time.sleep(.1);u.keybd_event(vk,0,2,0)
        else:u.PostMessageW(dialog,0x10,0,0)
        wait_for(lambda:not u.IsWindow(dialog),8);time.sleep(.5)
    import shutil,win32com.client,urllib.parse
    all_names=['alpha.txt','beta.txt','한글 폴더'];drag_source=fixture/'drag_source';baseline=snap(drag_source)
    master=fixture/'pristine_drag_source';shutil.copytree(drag_source,master)
    shell=win32com.client.Dispatch('Shell.Application');opened=[]
    for shift in (False,True):
        for names in (['alpha.txt'],all_names):
            for how in ('cancel','esc','confirm'):
                name=('move_' if shift else 'copy_')+str(len(names))+'_'+how
                dest=(pathlib.Path(r'C:\Users\ADMIN\AppData\Local\Temp\FxFileTask130')/suffix/('explorer_'+name) if '--c-drive' in sys.argv else fixture/('explorer_'+name));dest.mkdir(parents=True)
                shell.Open(str(dest))
                def find_explorer():
                    for win in shell.Windows():
                        try:
                            location=urllib.parse.unquote(win.LocationURL).replace('file:///','').replace('/','\\')
                            if location.casefold()==str(dest).casefold() and u.IsWindow(int(win.HWND)) and u.IsWindowVisible(int(win.HWND)):return win
                        except Exception:pass
                win=wait_for(find_explorer,20);opened.append(win);target=int(win.HWND)
                u.ShowWindow(target,1);u.SetWindowPos(target,P(-1),1200,70,680,730,0x40)
                def view_ready():
                    try:return win.Document.Folder.Self.Path.casefold()==str(dest).casefold()
                    except Exception:return False
                wait_for(view_ready,30);time.sleep(2)
                select(lists[3],names);index=item_names(lists[3]).index('alpha.txt')
                r=W.RECT();r.left=0;put(r);u.SendMessageW(lists[3],0x100E,index,remote)
                assert k.ReadProcessMemory(handle,remote,C.byref(r),C.sizeof(r),None)
                vr=rect(lists[3]);tr=rect(target)
                start=(vr.left+85,vr.top+(r.top+r.bottom)//2);end=(tr.right-150,tr.bottom-150)
                focus(lists[3]);u.SetWindowPos(target,P(-1),0,0,0,0,0x13);print('DRAG_TARGET '+str((target,end))+' RECT '+str((tr.left,tr.top,tr.right,tr.bottom)),flush=True);u.SetCursorPos(*start);u.mouse_event(2,0,0,0,0);time.sleep(.2)
                u.keybd_event(16 if shift else 17,0,0,0)
                for step in range(1,16):
                    u.SetCursorPos(start[0]+(end[0]-start[0])*step//15,start[1]+(end[1]-start[1])*step//15);time.sleep(.04)
                time.sleep(.8);u.mouse_event(4,0,0,0,0)
                d=consent('explorer_'+name,source_count=len(names))
                assert '외부 프로그램:' in text(u.GetDlgItem(d,1981))
                assert ('이동' if shift else '복사') in text(u.GetDlgItem(d,1980))
                assert not snap(dest) and snap(drag_source)==baseline
                # Keep Shift until OLE returns so the shell sees the approved effect.
                if how=='esc' or '--release-modifiers' in sys.argv:
                    u.keybd_event(17,0,2,0);u.keybd_event(16,0,2,0)
                respond_mouse(d,how);u.keybd_event(16,0,2,0);u.keybd_event(17,0,2,0)
                selected={key:value for key,value in baseline.items() if key.split('/')[0] in names or key.split('\\')[0] in names}
                if how=='confirm':
                    wait_for(lambda:snap(dest)==selected,30)
                    remaining={key:value for key,value in baseline.items() if key not in selected} if shift else baseline
                    wait_for(lambda:snap(drag_source)==remaining,20)
                    shutil.copytree(master,drag_source,dirs_exist_ok=True)
                    wait_for(lambda:snap(drag_source)==baseline)
                else:
                    time.sleep(.5);assert not snap(dest) and snap(drag_source)==baseline
                records[-1]['ActualMouseResponse']=how
                records[-1]['ModifierReleasedBeforeResponse']='--release-modifiers' in sys.argv
                records[-1]['ResultHashes']=snap(dest)
                print('EXPLORER_PASS '+name,flush=True)
                u.ShowWindow(target,0);time.sleep(.3)  # Keep owned windows alive until final cleanup; Quit is asynchronous.
    time.sleep(25);command(lists[4],30120);assert proc.wait(timeout=25)==0
    (evidence/(suffix+('_explorer_c_mouse.json' if '--c-drive' in sys.argv else '_explorer_mouse.json'))).write_text(json.dumps({'Architecture':arch,'Passed':True,'ActualMouseClicks':True,'Cases':records,'Scope':'Actual Explorer isolated destinations, single and multi file/folder copy/move; no cloud upload.'},ensure_ascii=False,indent=2),encoding='utf-8')
finally:
    u.keybd_event(16,0,2,0);u.keybd_event(17,0,2,0);u.mouse_event(4,0,0,0,0)
    if 'opened' in locals():
        for win in opened:
            try:win.Quit()
            except Exception:pass
    if remote:k.VirtualFreeEx(handle,remote,0,0x8000)
    if handle:k.CloseHandle(handle)
    if proc.poll() is None:
        for h in windows():
            if pid_of(h)==proc.pid and cls(h)=='#32770':u.PostMessageW(h,0x111,2,0)
        time.sleep(.5)
        for h in windows():
            if pid_of(h)==proc.pid and 'fxfile' in caption(h).lower() and cls(h)!='#32770':u.PostMessageW(h,0x111,30120,0)
