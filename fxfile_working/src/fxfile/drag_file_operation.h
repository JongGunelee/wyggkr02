// Copyright (c) 2026 fxfile contributors. GPLv3; see LICENSE.
#pragma once

namespace fxfile
{
// One policy for lists, folder trees, drive and bookmark destinations.
// The default must not silently remove the source on the same volume.
inline DROPEFFECT getDragFileOperation(DWORD aKeyState, xpr_sint_t aConfiguredOperation)
{
    if (aKeyState & MK_ALT)     return DROPEFFECT_LINK;
    if (aKeyState & MK_CONTROL) return DROPEFFECT_COPY;
    if (aKeyState & MK_SHIFT)   return DROPEFFECT_MOVE;
    if (aConfiguredOperation == DRAG_FILE_OP_MOVE) return DROPEFFECT_MOVE;
    if (aConfiguredOperation == DRAG_FILE_OP_LINK) return DROPEFFECT_LINK;
    return DROPEFFECT_COPY;
}
} // namespace fxfile
