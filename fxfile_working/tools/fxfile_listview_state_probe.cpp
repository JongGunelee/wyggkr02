#define UNICODE
#define _UNICODE

#include <windows.h>
#include <commctrl.h>

#include <algorithm>
#include <cstdio>
#include <vector>

#pragma comment(lib, "user32.lib")

namespace
{
struct ListViewState
{
    HWND window;
    RECT bounds;
    int  controlId;
    int  itemCount;
    int  selectedCount;
    int  selectionMark;
    int  focusedItem;
    int  selectedItem;
    std::vector<int> selectedItems;
    bool visible;
};

BOOL CALLBACK FindFxFileWindow(HWND window, LPARAM parameter)
{
    if (IsWindowVisible(window) == FALSE)
        return TRUE;

    wchar_t title[512] = { 0 };
    GetWindowText(window, title, static_cast<int>(sizeof(title) / sizeof(title[0])));

    const wchar_t suffix[] = L" - fxfile";
    const size_t titleLength = wcslen(title);
    const size_t suffixLength = (sizeof(suffix) / sizeof(suffix[0])) - 1;
    if (titleLength < suffixLength ||
        _wcsicmp(title + titleLength - suffixLength, suffix) != 0)
    {
        return TRUE;
    }

    *reinterpret_cast<HWND *>(parameter) = window;
    return FALSE;
}

BOOL CALLBACK CollectListViews(HWND window, LPARAM parameter)
{
    wchar_t className[128] = { 0 };
    GetClassName(window, className, static_cast<int>(sizeof(className) / sizeof(className[0])));
    if (_wcsicmp(className, WC_LISTVIEW) != 0)
        return TRUE;

    ListViewState state = {};
    state.window = window;
    GetWindowRect(window, &state.bounds);
    state.controlId = GetDlgCtrlID(window);
    state.itemCount = static_cast<int>(SendMessage(window, LVM_GETITEMCOUNT, 0, 0));
    state.selectedCount = static_cast<int>(SendMessage(window, LVM_GETSELECTEDCOUNT, 0, 0));
    state.selectionMark = static_cast<int>(SendMessage(window, LVM_GETSELECTIONMARK, 0, 0));
    state.focusedItem = static_cast<int>(SendMessage(window, LVM_GETNEXTITEM,
        static_cast<WPARAM>(-1), LVNI_FOCUSED));
    state.selectedItem = static_cast<int>(SendMessage(window, LVM_GETNEXTITEM,
        static_cast<WPARAM>(-1), LVNI_SELECTED));
    int selectedIndex = -1;
    while ((selectedIndex = static_cast<int>(SendMessage(window, LVM_GETNEXTITEM,
        static_cast<WPARAM>(selectedIndex), LVNI_SELECTED))) >= 0)
    {
        state.selectedItems.push_back(selectedIndex);
    }
    state.visible = IsWindowVisible(window) != FALSE;

    reinterpret_cast<std::vector<ListViewState> *>(parameter)->push_back(state);
    return TRUE;
}
}

int wmain()
{
    HWND fxfileWindow = nullptr;
    EnumWindows(FindFxFileWindow, reinterpret_cast<LPARAM>(&fxfileWindow));
    if (fxfileWindow == nullptr)
    {
        fwprintf(stderr, L"FxFile window was not found.\n");
        return 2;
    }

    std::vector<ListViewState> states;
    EnumChildWindows(fxfileWindow, CollectListViews, reinterpret_cast<LPARAM>(&states));
    std::sort(states.begin(), states.end(), [](const ListViewState &left, const ListViewState &right) {
        if (left.bounds.top != right.bounds.top)
            return left.bounds.top < right.bounds.top;
        return left.bounds.left < right.bounds.left;
    });

    int visiblePane = 0;
    for (const ListViewState &state : states)
    {
        if (state.visible)
            ++visiblePane;

        wprintf(L"pane=%d visible=%d hwnd=0x%p id=%d rect=%ld,%ld,%ld,%ld "
                L"items=%d selected_count=%d selection_mark=%d focused_item=%d selected_item=%d\n",
                state.visible ? visiblePane : 0,
                state.visible ? 1 : 0,
                state.window,
                state.controlId,
                state.bounds.left,
                state.bounds.top,
                state.bounds.right,
                state.bounds.bottom,
                state.itemCount,
                state.selectedCount,
                state.selectionMark,
                state.focusedItem,
                state.selectedItem);

        wprintf(L" selected_items=");
        if (state.selectedItems.empty())
        {
            wprintf(L"none");
        }
        else
        {
            for (size_t index = 0; index < state.selectedItems.size(); ++index)
            {
                if (index != 0)
                    wprintf(L",");
                wprintf(L"%d", state.selectedItems[index]);
            }
        }
        wprintf(L"\n");
    }

    return states.empty() ? 3 : 0;
}
