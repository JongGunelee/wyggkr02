// Read-only integration probe: production BBInfoFile, xpr runtime and real bin records.
#include "stdafx.h"
#include "bb_info_file.h"
#include "operation_path_pairs.h"
#include <cstdio>
#include <vector>
#include <cstring>

// Same native helper bodies as shell.cpp; no UI/Shell application object is needed.
namespace fxfile {
xpr_bool_t IsDriveSecure(xpr_tchar_t drive) {
    wchar_t root[4] = {drive,L':',L'\\',0}, fs[MAX_PATH+1]={0};
    ::GetVolumeInformationW(root,NULL,0,NULL,NULL,NULL,fs,MAX_PATH);
    return _wcsnicmp(fs,L"NTFS",4)==0;
}
const xpr_tchar_t *GetFileExt(const xpr_tchar_t *path) {
    if(!path)return NULL;const wchar_t *ext=::PathFindExtensionW(path);
    return ext && *ext ? ext : NULL;
}
}
CWinApp theProbeApp;
int wmain(int argc,wchar_t **argv) {
    if(argc!=2 && argc!=3)return 2;
    AfxWinInit(::GetModuleHandle(NULL),NULL,::GetCommandLineW(),0);
    if(!xpr::initialize())return 3;
    HANDLE file=::CreateFileW(argv[1],GENERIC_READ,FILE_SHARE_READ,NULL,OPEN_EXISTING,0,NULL);
    if(file==INVALID_HANDLE_VALUE)return 4;
    DWORD size=::GetFileSize(file,NULL), read=0;
    if(size<20 || size>1024*1024)return 5;
    std::vector<unsigned char> bytes(size);
    ::ReadFile(file,&bytes[0],size,&read,NULL);::CloseHandle(file);
    if(read!=size)return 6;
    int count,len;std::memcpy(&count,&bytes[4],4);std::memcpy(&len,&bytes[8],4);
    if(count<1 || len<1 || 12+len*2+count*8>static_cast<int>(size))return 7;
    std::vector<wchar_t> paths(len+1,0);std::memcpy(&paths[0],&bytes[12],len*2);
    fxfile::BBInfoFile bin;
    bool opened=bin.open(paths[0]);
    wchar_t directory[MAX_PATH+1]={0};bin.getDirPath(directory);
    std::wprintf(L"OS=%u NTFS=%d opened=%d records=%zu directory=%ls\n",xpr::getOsVer(),fxfile::IsDriveSecure(paths[0]),opened,bin.getCount(),directory);
    const wchar_t *path=&paths[0];int found=0;
    fxfile::OperationPathPairs<wchar_t> pairs;
    for(int i=0;i<count;++i){
        FILETIME time;std::memcpy(&time,&bytes[12+len*2+i*8],8);
        wchar_t recycled[MAX_PATH*2+1]={0};bool matched=bin.findOriginalPath(path,time,recycled);
        std::wprintf(L"item=%d matched=%d target=%ls\n",i,matched,recycled);
        if(matched){++found;pairs.add(recycled,path);}path+=wcslen(path)+1;
    }
    if(argc==3 && wcscmp(argv[2],L"--restore")==0 && found==count){
        path=&paths[0];
        for(int i=0;i<count;++i){
            if(!wcsstr(path,L"task127_results_20261012\\fixture_x64\\source\\") ||
                ::GetFileAttributesW(path)!=INVALID_FILE_ATTRIBUTES)return 9;
            path+=wcslen(path)+1;
        }
        SHFILEOPSTRUCTW operation={0};operation.wFunc=FO_MOVE;
        operation.pFrom=pairs.sources.c_str();operation.pTo=pairs.targets.c_str();
        operation.fFlags=FOF_MULTIDESTFILES|FOF_SILENT|FOF_NOERRORUI|FOF_NOCONFIRMATION;
        int result=::SHFileOperationW(&operation);
        std::printf("restore_result=%d aborted=%d source_chars=%zu target_chars=%zu\n",result,operation.fAnyOperationsAborted,pairs.sources.size(),pairs.targets.size());
        if(result || operation.fAnyOperationsAborted)return 10;
    }
    xpr::finalize();return found==count?0:8;
}
