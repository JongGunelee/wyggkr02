#include "stdafx.h"
#include "cmd_file_lock.h"

#include "file_lock_manager_dlg.h"
#include "folder_ctrl.h"
#include "explorer_ctrl.h"
#include "search_result_ctrl.h"
#include "functors.h"

namespace fxfile
{
namespace cmd
{
xpr_sint_t FileLockManagerCommand::canExecute(CommandContext &aContext)
{
    FXFILE_COMMAND_DECLARE_CTRL;
    // Keep the manager reachable from any real filesystem pane.  When no
    // child item is selected, execute() intentionally falls back to the
    // pane's current folder; this also avoids menu-focus transitions making a
    // valid selection look empty while CCmdUI is updating the popup.
    if (sExplorerCtrl != NULL && sExplorerCtrl->isFileSystemFolder())
        return StateEnable;
    if (sSearchResultCtrl != NULL && sSearchResultCtrl->GetSelectedCount() > 0)
        return StateEnable;
    if (sFolderCtrl != NULL && sFolderCtrl->isFileSystem())
        return StateEnable;
    return StateDisable;
}

void FileLockManagerCommand::execute(CommandContext &aContext)
{
    FXFILE_COMMAND_DECLARE_CTRL;
    std::vector<std::wstring> sPaths;
    FXFILE_COMMAND_IF_FOLDER_CTRL
    {
        HTREEITEM sItem = sFolderCtrl->GetSelectedItem();
        LPTVITEMDATA sData = sItem == NULL ? NULL :
            reinterpret_cast<LPTVITEMDATA>(sFolderCtrl->GetItemData(sItem));
        if (sData != NULL &&
            XPR_TEST_BITS(sData->mShellAttributes, SFGAO_FILESYSTEM))
        {
            wchar_t sPath[XPR_MAX_PATH + 1] = {0};
            GetName(sData->mShellFolder, sData->mPidl, SHGDN_FORPARSING, sPath);
            sPaths.push_back(sPath);
        }
    }
    FXFILE_COMMAND_ELSE_IF_SEARCH_RESULT_CTRL
    {
        POSITION sPosition = sSearchResultCtrl->GetFirstSelectedItemPosition();
        while (sPosition != NULL)
        {
            const int sIndex = sSearchResultCtrl->GetNextSelectedItem(sPosition);
            SrItemData *sData = reinterpret_cast<SrItemData *>(
                sSearchResultCtrl->GetItemData(sIndex));
            if (sData != NULL)
            {
                wchar_t sPath[XPR_MAX_PATH + 1] = {0};
                sData->getPath(sPath);
                sPaths.push_back(sPath);
            }
        }
    }
    FXFILE_COMMAND_ELSE_IF_EXPLORER_CTRL
    {
        FileSysItemDeque sItems;
        sExplorerCtrl->getSelFileSysItems(sItems, XPR_FALSE, SFGAO_FILESYSTEM);
        for (FileSysItemDeque::const_iterator sIt = sItems.begin();
             sIt != sItems.end(); ++sIt)
            if (*sIt != NULL)
                sPaths.push_back(std::wstring((*sIt)->mPath.c_str()));
        clear(sItems);

    }

    // Popup menus own keyboard focus while execute() is entered.  The common
    // command macros therefore may skip their focus-sensitive branch even
    // though CommandContext already resolved the active pane.  Fall back to
    // that resolved pane's current filesystem folder so the reachable menu
    // command always opens a meaningful, reversible manager.
    if (sPaths.empty() && sExplorerCtrl != NULL &&
        sExplorerCtrl->isFileSystemFolder())
    {
        const xpr_tchar_t *sCurrentPath = sExplorerCtrl->getCurPath();
        if (sCurrentPath != NULL && *sCurrentPath != XPR_STRING_LITERAL('\0'))
            sPaths.push_back(std::wstring(sCurrentPath));
    }
    if (sPaths.empty() && sFolderCtrl != NULL && sFolderCtrl->isFileSystem())
    {
        xpr::string sCurrentPath;
        sFolderCtrl->getCurPath(sCurrentPath);
        if (!sCurrentPath.empty())
            sPaths.push_back(std::wstring(sCurrentPath.c_str()));
    }
    if (sPaths.empty())
        return;
    FileLockManagerDlg sDialog(sPaths);
    sDialog.DoModal();
}
} // namespace cmd
} // namespace fxfile
