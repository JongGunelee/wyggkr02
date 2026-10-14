"""Exercise consent in an isolated Task129 GUI fixture; never targets user data."""
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
evidence = workspace / '__BUILD_TEMP_BACKUP__' / 'task129_results_20261014'
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
    def respond(dialog,accept=False,close=False,enter=False):
        if enter:
            # Send through this dialog's message queue, independent of global
            # foreground changes made by another desktop application.
            u.PostMessageW(dialog,0x100,13,1);u.PostMessageW(dialog,0x102,13,1);u.PostMessageW(dialog,0x101,13,1)
        elif close:u.PostMessageW(dialog,0x10,0,0)
        else:u.PostMessageW(u.GetDlgItem(dialog,1 if accept else 2),0xF5,0,0)  # BM_CLICK
        wait_for(lambda:not u.IsWindow(dialog));time.sleep(1)
    def expect_files(folder,expected):
        wait_for(lambda:snap(folder)==expected,45)
    all_names=['alpha.txt','beta.txt','한글 폴더']
    source=fixture/'source';copy=fixture/'copy_target';move=fixture/'move_target'
    baseline=snap(source);assert len(baseline)==3
    select(lists[0],all_names);command(lists[0],0xE122);time.sleep(.5)
    sequence=u.GetClipboardSequenceNumber();command(lists[1],0xE125)
    d=consent('copy_cancel',copy,3);assert snap(source)==baseline and not snap(copy)
    respond(d,enter=True);assert snap(source)==baseline and not snap(copy);assert u.GetClipboardSequenceNumber()==sequence
    for ordinal in range(2):
        command(lists[1],0xE125);d=consent('copy_enter_cancel_repeat_'+str(ordinal+1),copy,3)
        respond(d,enter=True);assert snap(source)==baseline and not snap(copy)
        assert u.GetClipboardSequenceNumber()==sequence
    command(lists[1],0xE125);d=consent('copy_accept',copy,3);assert not snap(copy);respond(d,True)
    expect_files(copy,baseline);assert snap(source)==baseline
    select(lists[0],all_names);sequence=u.GetClipboardSequenceNumber();states_before=cut_states(lists[0]);command(lists[0],0xE123)
    d=consent('cut_cancel',source_count=3);assert cut_states(lists[0])==states_before;respond(d)
    assert u.GetClipboardSequenceNumber()==sequence and snap(source)==baseline and cut_states(lists[0])==states_before
    command(lists[0],0xE123);d=consent('cut_accept',source_count=3);respond(d,True)
    sequence=wait_for(lambda:u.GetClipboardSequenceNumber() if u.GetClipboardSequenceNumber()!=sequence else None)
    assert cut_states(lists[0])==[4,4,4]
    assert snap(source)==baseline
    command(lists[2],0xE125);d=consent('move_cancel',move,3);respond(d,close=True)
    assert snap(source)==baseline and not snap(move) and u.GetClipboardSequenceNumber()==sequence
    command(lists[2],0xE125);d=consent('move_stale_clipboard',move,3)
    clipboard_text=C.create_unicode_buffer('Task129 newer clipboard must survive rejected move.')
    global_mem=k.GlobalAlloc(0x42,C.sizeof(clipboard_text));assert global_mem
    pointer=k.GlobalLock(global_mem);assert pointer;C.memmove(pointer,clipboard_text,C.sizeof(clipboard_text));k.GlobalUnlock(global_mem)
    assert u.OpenClipboard(main);u.EmptyClipboard();assert u.SetClipboardData(13,global_mem);u.CloseClipboard()
    newer_sequence=u.GetClipboardSequenceNumber();respond(d,True)
    for unused in range(20):
        for notice in dialogs():
            buttons=[child for child in windows(notice) if cls(child)=='Button' and u.IsWindowVisible(child)]
            if 'FxFile' in caption(notice) and len(buttons)==1:
                print('ACK_STALE_CLIPBOARD_NOTICE '+str(u.GetDlgCtrlID(buttons[0])),flush=True)
                u.PostMessageW(notice,0x111,u.GetDlgCtrlID(buttons[0]),buttons[0])
                wait_for(lambda:not u.IsWindow(notice))
        if u.IsWindowEnabled(main) and not dialogs():break
        time.sleep(.2)
    assert u.IsWindowEnabled(main) and not dialogs(),'Stale clipboard notice not dismissed'
    assert snap(source)==baseline and not snap(move) and u.GetClipboardSequenceNumber()==newer_sequence
    assert u.IsClipboardFormatAvailable(13) and not u.IsClipboardFormatAvailable(15)
    select(lists[0],all_names);command(lists[0],0xE123);d=consent('cut_after_stale_clipboard',source_count=3);respond(d,True)
    sequence=u.GetClipboardSequenceNumber()
    command(lists[2],0xE125);d=consent('move_accept',move,3);assert snap(source)==baseline and not snap(move);respond(d,True)
    expect_files(move,baseline);wait_for(lambda:not snap(source));time.sleep(2)
    if '--external' not in sys.argv:assert not u.IsClipboardFormatAvailable(15),'Completed move must clear its unchanged cut clipboard'
    delete=fixture/'delete_source';deleted_baseline=snap(delete)
    select(lists[5],['alpha.txt']);command(lists[5],30112);d=consent('delete_single_cancel',source_count=1)
    assert '휴지통으로 삭제' in text(u.GetDlgItem(d,1980));respond(d);assert snap(delete)==deleted_baseline
    select(lists[5],all_names)
    u.keybd_event(16,0,0,0);command(lists[5],30112)
    d=consent('delete_permanent_accept',source_count=3);u.keybd_event(16,0,2,0)
    assert '영구 삭제' in text(u.GetDlgItem(d,1980));assert snap(delete)==deleted_baseline;respond(d,True)
    deadline=time.monotonic()+45
    while snap(delete) and time.monotonic()<deadline:
        for native in dialogs():
            yes=u.GetDlgItem(native,6)
            if yes:u.PostMessageW(native,0x111,6,0)
        time.sleep(.2)
    assert not snap(delete),'Permanent delete did not complete'
    drag_source=fixture/'drag_source';drag_target=fixture/'drag_target';drag_baseline=snap(drag_source)
    def drag(shift=False,names=None):
        select(lists[3],all_names if names is None else names);view=lists[3];index=item_names(view).index('alpha.txt')
        r=W.RECT();r.left=0;put(r);u.SendMessageW(view,0x100E,index,remote);assert k.ReadProcessMemory(handle,remote,C.byref(r),C.sizeof(r),None)
        vr=rect(view);tr=rect(lists[4]);start=(vr.left+85,vr.top+(r.top+r.bottom)//2);end=(tr.left+80,tr.bottom-45)
        u.SetCursorPos(*start);u.mouse_event(2,0,0,0,0);time.sleep(.2)
        # Shift after the initial press changes drag effect without collapsing
        # the existing multi-selection into a Shift-click selection range.
        if shift:u.keybd_event(16,0,0,0)
        for step in range(1,16):
            u.SetCursorPos(start[0]+(end[0]-start[0])*step//15,start[1]+(end[1]-start[1])*step//15);time.sleep(.04)
        time.sleep(1.2)  # Allow the real OLE target to finish DragOver feedback.
        print('DRAG '+json.dumps({'start':start,'end':end,'selected':len(all_names if names is None else names),'shift':shift}),flush=True)
        u.mouse_event(4,0,0,0,0)
    drag(names=['alpha.txt']);d=consent('drag_single_text_cancel',drag_target,1)
    assert snap(drag_source)==drag_baseline and not snap(drag_target);respond(d)
    assert snap(drag_source)==drag_baseline and not snap(drag_target)
    for ordinal in range(3):
        # The user dump crashed in source IDataObject::Release after the first
        # cancelled shell drop. Exercise that lifetime repeatedly before reuse.
        assert proc.poll() is None
        drag(names=['alpha.txt']);d=consent('drag_single_cancel_repeat_'+str(ordinal+1),drag_target,1)
        respond(d);assert snap(drag_source)==drag_baseline and not snap(drag_target)
        assert proc.poll() is None and item_names(lists[3])
    drag();d=consent('drag_copy_cancel',drag_target,3);assert snap(drag_source)==drag_baseline and not snap(drag_target)
    respond(d);assert snap(drag_source)==drag_baseline and not snap(drag_target)
    drag(True);d=consent('drag_move_accept',drag_target,3);u.keybd_event(16,0,2,0)
    assert '이동' in text(u.GetDlgItem(d,1980));assert snap(drag_source)==drag_baseline and not snap(drag_target);respond(d,True)
    expect_files(drag_target,drag_baseline);wait_for(lambda:not snap(drag_source))
    screenshot(main,evidence/(suffix+'_final_screen.png'))
    command(lists[4],30120);assert proc.wait(timeout=25)==0
    report={'Architecture':arch,'External': '--external' in sys.argv,'Passed':True,'ExitCode':proc.returncode,
            'SourceOriginalHashes':baseline,'AcceptedCopyHashes':snap(copy),'AcceptedMoveHashes':snap(move),
            'DragOriginalHashes':drag_baseline,'AcceptedDragMoveHashes':snap(drag_target),
            'Dialogs':records,'Assertions':'Pre-confirmation and cancellation preserve hashes; accepted copy/move hashes match; permanent deletion occurs only after acceptance; cut clipboard preserved on cancellation; real mouse drag copy cancellation and Shift move verified'}
    (evidence/(suffix+'_runtime.json')).write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print('RUNTIME_PASS '+suffix,flush=True)
finally:
    if proc.poll() is None:
        for ordinal,h in enumerate(dialogs()):
            screenshot(h,evidence/(suffix+'_failure_dialog_'+str(ordinal)+'.png'))
        (evidence/(suffix+'_partial.json')).write_text(json.dumps({'Dialogs':records,'Files':{name:snap(fixture/name) for name in ['source','copy_target','move_target','drag_source','drag_target','delete_source']}},ensure_ascii=False,indent=2),encoding='utf-8')
    u.keybd_event(16,0,2,0);u.mouse_event(4,0,0,0,0)
    if remote:k.VirtualFreeEx(handle,remote,0,0x8000)
    if handle:k.CloseHandle(handle)
    if proc.poll() is None:
        # Preserve a failed fixture for investigation. Do not force-kill or
        # auto-accept any pending operation; normal cancellation is safe.
        for h in windows():
            if pid_of(h)==proc.pid and cls(h)=='#32770':u.PostMessageW(h,0x111,2,0)
        time.sleep(.5)
        for h in windows():
            if pid_of(h)==proc.pid and 'fxfile' in text(h).lower() and cls(h)!='#32770':u.PostMessageW(h,0x111,30120,0)
