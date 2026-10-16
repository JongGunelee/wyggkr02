// Mandatory, per-request consent for user file operations. No saved bypass.
#pragma once
#include <string>
#include <vector>

namespace fxfile {
class FileOperationConfirmation {
public:
    enum { Cut = 0x100, Link = 0x101, ExternalDrop = 0x102 };
    static bool confirm(HWND owner, UINT operation,
        const std::vector<std::wstring>& sources,
        const std::vector<std::wstring>& targets, WORD flags = 0);
    static bool confirm(const SHFILEOPSTRUCT* operation);
    static bool dataObjectPaths(IDataObject* data, std::vector<std::wstring>& paths);
    static bool confirmCut(HWND owner, IDataObject* data);
    static bool confirmShellVerb(HWND owner, const wchar_t* verb,
        IShellFolder* folder, LPCITEMIDLIST* pidls, UINT count);
    static bool confirmDrop(HWND owner, IDataObject* data, DWORD effect,
        const wchar_t* target);
    static DWORD chooseDropEffect(HWND owner, DWORD allowed);
    static std::vector<std::wstring> paths(const wchar_t* multi);
};
}
