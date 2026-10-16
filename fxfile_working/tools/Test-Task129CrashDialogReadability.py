"""Preview compiled BugTrap resources without raising an exception or executing DLL code."""
import ctypes as C
from ctypes import wintypes as W
import json, pathlib, sys
from PIL import Image

workspace=pathlib.Path(__file__).resolve().parents[2]
evidence=workspace/'__BUILD_TEMP_BACKUP__'/'task129_results_20261014'
arch=sys.argv[1]
assert arch in ['x64','x32']
u,k,g=C.windll.user32,C.windll.kernel32,C.windll.gdi32
P=C.c_void_p
k.LoadLibraryExW.argtypes=[W.LPCWSTR,P,W.DWORD];k.LoadLibraryExW.restype=P
k.FreeLibrary.argtypes=[P]
PROC=C.WINFUNCTYPE(C.c_ssize_t,P,W.UINT,C.c_size_t,C.c_ssize_t)
u.CreateDialogParamW.argtypes=[P,P,P,PROC,C.c_ssize_t];u.CreateDialogParamW.restype=P
u.DestroyWindow.argtypes=[P];u.GetDlgItem.argtypes=[P,C.c_int];u.GetDlgItem.restype=P
u.SendMessageW.argtypes=[P,W.UINT,C.c_size_t,C.c_ssize_t];u.SendMessageW.restype=C.c_ssize_t
u.ShowWindow.argtypes=[P,C.c_int];u.UpdateWindow.argtypes=[P]
u.GetWindowRect.argtypes=[P,P];u.GetClientRect.argtypes=[P,P]
u.GetWindowTextW.argtypes=[P,W.LPWSTR,C.c_int]
u.GetDC.argtypes=[P];u.GetDC.restype=P;u.ReleaseDC.argtypes=[P,P]
u.DrawTextW.argtypes=[P,W.LPCWSTR,C.c_int,P,W.UINT]
u.PrintWindow.argtypes=[P,P,W.UINT]
for name in ['CreateCompatibleDC','CreateCompatibleBitmap','SelectObject']:
    getattr(g,name).restype=P
g.CreateCompatibleDC.argtypes=[P];g.CreateCompatibleBitmap.argtypes=[P,C.c_int,C.c_int]
g.SelectObject.argtypes=[P,P];g.GetObjectW.argtypes=[P,C.c_int,P]
g.DeleteObject.argtypes=[P];g.DeleteDC.argtypes=[P]
class FONT(C.Structure):
    _fields_=[('height',W.LONG),('width',W.LONG),('escape',W.LONG),('orient',W.LONG),('weight',W.LONG),
        ('italic',W.BYTE),('underline',W.BYTE),('strike',W.BYTE),('charset',W.BYTE),('out',W.BYTE),
        ('clip',W.BYTE),('quality',W.BYTE),('pitch',W.BYTE),('face',W.WCHAR*32)]
class HEADER(C.Structure):
    _fields_=[('size',W.DWORD),('width',W.LONG),('height',W.LONG),('planes',W.WORD),('bits',W.WORD),
        ('compression',W.DWORD),('image',W.DWORD),('xp',W.LONG),('yp',W.LONG),('used',W.DWORD),('important',W.DWORD)]
@PROC
def preview(hwnd,msg,wp,lp):
    return 1 if msg==0x110 else 0

module=k.LoadLibraryExW(str(workspace/'fxfile_working'/'bin'/arch/'fxfile-crash.dll'),None,0x22)
assert module
dialog=u.CreateDialogParamW(module,P(111),None,preview,0)
assert dialog
try:
    u.ShowWindow(dialog,1);u.UpdateWindow(dialog)
    outer=W.RECT();assert u.GetWindowRect(dialog,C.byref(outer))
    controls=[]
    # Intro, warning, explanatory paragraphs and all action buttons.
    for identifier in [1031,1022,1025,1002,1013,2]:
        control=u.GetDlgItem(dialog,identifier)
        if not control:continue
        text=C.create_unicode_buffer(4096);u.GetWindowTextW(control,text,len(text))
        font=FONT();hf=u.SendMessageW(control,0x31,0,0)
        assert hf and g.GetObjectW(hf,C.sizeof(font),C.byref(font))
        assert abs(font.height)>=16
        bounds=W.RECT();assert u.GetWindowRect(control,C.byref(bounds))
        assert outer.left<=bounds.left<bounds.right<=outer.right and outer.top<=bounds.top<bounds.bottom<=outer.bottom
        dc=u.GetDC(control);old=g.SelectObject(dc,hf)
        measured=W.RECT(0,0,bounds.right-bounds.left,0)
        u.DrawTextW(dc,text.value,-1,C.byref(measured),0x400|0x10)
        g.SelectObject(dc,old);u.ReleaseDC(control,dc)
        assert measured.bottom<=bounds.bottom-bounds.top,(identifier,text.value,measured.bottom,bounds.bottom-bounds.top)
        controls.append({'Control':identifier,'Text':text.value,'FontPixels':font.height,'Face':font.face,
            'RequiredHeight':measured.bottom,'ControlHeight':bounds.bottom-bounds.top})
    assert len(controls)>=4
    width,height=outer.right-outer.left,outer.bottom-outer.top
    screen=u.GetDC(None);dc=g.CreateCompatibleDC(screen);bitmap=g.CreateCompatibleBitmap(screen,width,height)
    old=g.SelectObject(dc,bitmap);assert u.PrintWindow(dialog,dc,2);g.SelectObject(dc,old)
    header=HEADER(40,width,-height,1,32,0,0,0,0,0,0);pixels=C.create_string_buffer(width*height*4)
    g.GetDIBits.argtypes=[P,P,W.UINT,W.UINT,P,P,W.UINT]
    assert g.GetDIBits(dc,bitmap,0,height,pixels,C.byref(header),0)
    Image.frombytes('RGB',(width,height),pixels.raw,'raw','BGRX').save(evidence/(arch+'_crash_dialog_readable.png'))
    g.DeleteObject(bitmap);g.DeleteDC(dc);u.ReleaseDC(None,screen)
    (evidence/(arch+'_crash_dialog_readability.json')).write_text(json.dumps({'Passed':True,'Architecture':arch,
        'ResourceOnly':True,'Controls':controls,'WindowSize':[width,height]},ensure_ascii=False,indent=2),encoding='utf-8')
    print('CRASH_DIALOG_READABILITY_PASS '+arch)
finally:
    u.DestroyWindow(dialog);k.FreeLibrary(module)
