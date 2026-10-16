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
            except FileNotFoundError:pass  # Accepted deletion/move can race this observation.
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
    user_folder=pathlib.Path(r'D:\00_USER-AUTOCAD\01 Utility\습공기 선도\장한기술')
    def metadata():return {str(p.relative_to(user_folder)):(p.stat().st_size,p.stat().st_mtime_ns) for p in user_folder.rglob('*') if p.is_file()}
    before=metadata()
    def double_click_item(view,name):
        names=item_names(view);index=names.index(name) if name else 0
        r=W.RECT();r.left=0;put(r);u.SendMessageW(view,0x100E,index,remote)
        assert k.ReadProcessMemory(handle,remote,C.byref(r),C.sizeof(r),None)
        vr=rect(view);point=(vr.left+85,vr.top+(r.top+r.bottom)//2)
        focus(view);u.SetCursorPos(*point)
        for unused in range(2):
            u.mouse_event(2,0,0,0,0);time.sleep(.045);u.mouse_event(4,0,0,0,0);time.sleep(.045)
    for cycle in range(5):
        wait_for(lambda:'장한기술' in item_names(lists[0]),45)
        start=time.monotonic();double_click_item(lists[0],'장한기술')
        wait_for(lambda:('장한기술 -' in caption(main) and len(item_names(lists[0]))>1),45)
        entered=time.monotonic()-start;names=item_names(lists[0]);time.sleep(2)
        double_click_item(lists[0],None)
        wait_for(lambda:'장한기술' in item_names(lists[0]),45)
        records.append({'Cycle':cycle+1,'EnteredSeconds':round(entered,3),'RealRowCount':len(names)-1,'Names':names,'ParentReturned':True})
        print('NAVIGATION_PASS '+str(cycle+1)+' '+str(entered),flush=True)
    time.sleep(25);screenshot(main,evidence/(suffix+'_navigation.png'))
    command(lists[0],30120);assert proc.wait(timeout=30)==0
    assert before==metadata(),'Read-only navigation changed user files'
    (evidence/(suffix+'_navigation.json')).write_text(json.dumps({'Architecture':arch,'Passed':True,'Path':str(user_folder),'ActualMouseDoubleClick':True,'Cycles':records,'SourceFileMetadataUnchanged':True,'ExitCode':proc.returncode},ensure_ascii=False,indent=2),encoding='utf-8')
finally:
    u.mouse_event(4,0,0,0,0)
    if remote:k.VirtualFreeEx(handle,remote,0,0x8000)
    if handle:k.CloseHandle(handle)
    if proc.poll() is None:
        for h in windows():
            if pid_of(h)==proc.pid and cls(h)=='#32770':u.PostMessageW(h,0x111,2,0)
        for h in windows():
            if pid_of(h)==proc.pid and 'fxfile' in caption(h).lower() and cls(h)!='#32770':u.PostMessageW(h,0x111,30120,0)
