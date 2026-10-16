#include <windows.h>
#include <ole2.h>
#include <shlobj.h>
#include <fstream>
#include <string>
static std::wstring mode;
static void log(const std::string& s){std::ofstream f(mode+L"_native.log",std::ios::app);f<<GetTickCount()<<" "<<s<<"\n";}
struct Target:IDropTarget{
 ULONG refs=1;
 HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id,void** p){*p=nullptr;if(id==IID_IUnknown||id==IID_IDropTarget){*p=this;AddRef();return S_OK;}return E_NOINTERFACE;}
 ULONG STDMETHODCALLTYPE AddRef(){return ++refs;} ULONG STDMETHODCALLTYPE Release(){auto r=--refs;if(!r)delete this;return r;}
 HRESULT STDMETHODCALLTYPE DragEnter(IDataObject*,DWORD keys,POINTL,DWORD* e){log("ENTER");*e=(keys&MK_SHIFT)?DROPEFFECT_MOVE:DROPEFFECT_COPY;return S_OK;}
 HRESULT STDMETHODCALLTYPE DragOver(DWORD keys,POINTL,DWORD* e){*e=(keys&MK_SHIFT)?DROPEFFECT_MOVE:DROPEFFECT_COPY;return S_OK;}
 HRESULT STDMETHODCALLTYPE DragLeave(){log("LEAVE");return S_OK;}
 HRESULT STDMETHODCALLTYPE Drop(IDataObject* data,DWORD keys,POINTL,DWORD* e){log("DROP"); FORMATETC fmt={CF_HDROP,nullptr,DVASPECT_CONTENT,-1,TYMED_HGLOBAL};STGMEDIUM medium={};if(SUCCEEDED(data->GetData(&fmt,&medium))){auto drop=(HDROP)medium.hGlobal;log("FILES "+std::to_string(DragQueryFileW(drop,0xffffffff,nullptr,0)));ReleaseStgMedium(&medium);}*e=(keys&MK_SHIFT)?DROPEFFECT_MOVE:DROPEFFECT_COPY;return S_OK;}
};
struct Source:IDropSource{
 ULONG refs=1;DWORD effect=0,start=GetTickCount();HWND owner;bool prompted=false;
 HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id,void** p){*p=nullptr;if(id==IID_IUnknown||id==IID_IDropSource){*p=this;AddRef();return S_OK;}return E_NOINTERFACE;}
 ULONG STDMETHODCALLTYPE AddRef(){return ++refs;} ULONG STDMETHODCALLTYPE Release(){auto r=--refs;if(!r)delete this;return r;}
 HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL esc,DWORD keys){
  log("QUERY "+std::to_string(esc)+" "+std::to_string(keys)+" effect="+std::to_string(effect));
  if(esc||GetTickCount()-start>15000)return DRAGDROP_S_CANCEL;
  if(keys&(MK_LBUTTON|MK_RBUTTON))return S_OK;
  if(!effect)return S_OK;
  if(mode==L"plain")return DRAGDROP_S_DROP;
  if(prompted)return DRAGDROP_S_CANCEL;prompted=true;
  HWND capture=GetCapture();log("CAPTURE "+std::to_string((UINT_PTR)capture));
  ReleaseCapture();log("RELEASED "+std::to_string((UINT_PTR)GetCapture()));
  int result=MessageBoxW(owner,L"Synthetic capture test only. Target logs a drop; no files are copied.",L"FX130 Native Capture Confirmation",MB_OKCANCEL|MB_DEFBUTTON2|MB_TASKMODAL);
  log("DIALOG "+std::to_string(result));
  if(mode==L"recapture")SetCapture(capture);
  return result==IDOK?DRAGDROP_S_DROP:DRAGDROP_S_CANCEL;
 }
 HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD e){effect=e&7;log("FEEDBACK "+std::to_string(effect));return DRAGDROP_S_USEDEFAULTCURSORS;}
};
static LRESULT CALLBACK proc(HWND h,UINT m,WPARAM w,LPARAM l){if(m==WM_DESTROY){RevokeDragDrop(h);PostQuitMessage(0);return 0;}return DefWindowProcW(h,m,w,l);}
int wmain(int argc,wchar_t** argv){
 if(argc<2)return 1;mode=argv[1];std::ofstream(mode+L"_native.log",std::ios::trunc).close();OleInitialize(nullptr);
 WNDCLASSW wc={};wc.lpfnWndProc=proc;wc.hInstance=GetModuleHandleW(nullptr);wc.lpszClassName=L"FX130NativeProbe";wc.hbrBackground=(HBRUSH)(COLOR_WINDOW+1);RegisterClassW(&wc);
 bool target=mode==L"target";HWND h=CreateWindowW(wc.lpszClassName,target?L"FX130 Native Target":L"FX130 Native Source",WS_OVERLAPPEDWINDOW,target?1200:20,100,450,400,nullptr,nullptr,wc.hInstance,nullptr);
 ShowWindow(h,SW_SHOW);ShowWindow(h,SW_SHOW);SetWindowPos(h,HWND_TOPMOST,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE|SWP_SHOWWINDOW);
 if(target){Target* t=new Target;auto hr=RegisterDragDrop(h,t);log("REGISTER "+std::to_string(hr));std::ofstream("native_target_hwnd.txt")<<(UINT_PTR)h;MSG msg;while(GetMessageW(&msg,nullptr,0,0)>0){TranslateMessage(&msg);DispatchMessageW(&msg);}t->Release();OleUninitialize();return 0;}
 if(argc<3)return 2;HWND dest=(HWND)_wcstoui64(argv[2],nullptr,10);SetWindowPos(dest,HWND_TOPMOST,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE|SWP_SHOWWINDOW);SetForegroundWindow(dest);
 RECT r;GetWindowRect(dest,&r);POINT point={(r.left+r.right)/2,(r.top+r.bottom)/2};SetCursorPos(point.x,point.y);log("TARGET "+std::to_string((UINT_PTR)dest)+" POINTWINDOW "+std::to_string((UINT_PTR)WindowFromPoint(point)));
 wchar_t dir[MAX_PATH];GetCurrentDirectoryW(MAX_PATH,dir);std::wstring file=std::wstring(dir)+L"\\probe_synthetic.txt";std::ofstream(file)<<"Synthetic capture test";
 PIDLIST_ABSOLUTE parent=nullptr,full=nullptr;SHParseDisplayName(dir,nullptr,&parent,0,nullptr);SHParseDisplayName(file.c_str(),nullptr,&full,0,nullptr);LPCITEMIDLIST child=ILFindLastID(full);IDataObject* data=nullptr;
 HRESULT hr=SHCreateDataObject(parent,1,&child,nullptr,IID_PPV_ARGS(&data));if(FAILED(hr))return 3;
 Source* source=new Source;source->owner=h;
 SetForegroundWindow(h);SetCursorPos(100,200);mouse_event(MOUSEEVENTF_LEFTDOWN,0,0,0,0);log("DRAGSTART");
 DWORD effect=0;hr=DoDragDrop(data,source,DROPEFFECT_COPY,&effect);log("RESULT "+std::to_string(hr)+" EFFECT "+std::to_string(effect));mouse_event(MOUSEEVENTF_LEFTUP,0,0,0,0);
 data->Release();source->Release();CoTaskMemFree(parent);CoTaskMemFree(full);DestroyWindow(h);OleUninitialize();return 0;
}
