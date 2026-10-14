#ifndef FXFILE_WINDOW_LAYOUT_BATCH_H
#define FXFILE_WINDOW_LAYOUT_BATCH_H

#include <vector>

namespace fxfile
{
// Child-window placement only; no activation, redraw suppression or cached rectangles.
class WindowLayoutBatch
{
public:
    void add(HWND window, const RECT &rect)
    {
        if (!window) return;
        RECT current;
        if (::GetWindowRect(window, &current))
        {
            ::MapWindowPoints(HWND_DESKTOP, ::GetParent(window),
                              reinterpret_cast<POINT *>(&current), 2);
            if (::EqualRect(&current, &rect)) return;
        }
        Placement placement = { window, rect };
        mPlacements.push_back(placement);
    }

    bool apply()
    {
        if (mPlacements.empty()) return true;
        HDWP batch = ::BeginDeferWindowPos(static_cast<int>(mPlacements.size()));
        for (size_t i = 0; batch && i < mPlacements.size(); ++i)
        {
            const Placement &p = mPlacements[i];
            batch = ::DeferWindowPos(batch, p.window, NULL,
                p.rect.left, p.rect.top, p.rect.right - p.rect.left,
                p.rect.bottom - p.rect.top, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        if (batch && ::EndDeferWindowPos(batch)) return true;
        // A failed DeferWindowPos invalidates the entire batch, including earlier entries.
        bool success = true;
        for (size_t i = 0; i < mPlacements.size(); ++i)
        {
            const Placement &p = mPlacements[i];
            if (!::SetWindowPos(p.window, NULL, p.rect.left, p.rect.top,
                p.rect.right - p.rect.left, p.rect.bottom - p.rect.top,
                SWP_NOZORDER | SWP_NOACTIVATE)) success = false;
        }
        return success;
    }

private:
    struct Placement { HWND window; RECT rect; };
    std::vector<Placement> mPlacements;
};
}
#endif
