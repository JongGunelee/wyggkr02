// GPLv3. Tests the production policy, not a reimplementation.
#include <windows.h>
#include <oleidl.h>
#include <cstdio>
#include <initializer_list>
typedef int xpr_sint_t;
typedef DWORD DROPEFFECT; // MFC's afxole.h alias in the application.
enum { DRAG_FILE_OP_DEFAULT=0, DRAG_FILE_OP_MOVE, DRAG_FILE_OP_COPY, DRAG_FILE_OP_LINK };
#include "../src/fxfile/drag_file_operation.h"
int main()
{
    unsigned tested=0;
    const DWORD keys[] = {0, MK_SHIFT, MK_CONTROL, MK_ALT,
        MK_SHIFT|MK_CONTROL, MK_SHIFT|MK_ALT, MK_CONTROL|MK_ALT,
        MK_SHIFT|MK_CONTROL|MK_ALT};
    const DROPEFFECT expected[] = {DROPEFFECT_COPY, DROPEFFECT_MOVE,
        DROPEFFECT_COPY, DROPEFFECT_LINK, DROPEFFECT_COPY,
        DROPEFFECT_LINK, DROPEFFECT_LINK, DROPEFFECT_LINK};
    for(int option=-1; option<=4; ++option)
    for(unsigned k=0; k<8; ++k)
    for(DWORD button : {DWORD(0), DWORD(MK_LBUTTON), DWORD(MK_RBUTTON)})
    {
        DROPEFFECT want=expected[k];
        if(k==0 && option==DRAG_FILE_OP_MOVE) want=DROPEFFECT_MOVE;
        if(k==0 && option==DRAG_FILE_OP_LINK) want=DROPEFFECT_LINK;
        if(fxfile::getDragFileOperation(keys[k]|button,option)!=want)
        { std::printf("FAIL option=%d keys=%lu\n",option,keys[k]|button); return 1; }
        ++tested;
    }
    std::printf("PASS production drag policy %u cases\n",tested);
    return 0;
}
