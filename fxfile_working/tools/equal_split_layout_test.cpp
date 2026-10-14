#include <windows.h>
#include <cassert>
#include <cstdio>
#include <cstdlib>
#include <climits>
#include "../src/fxfile/equal_split_geometry.h"
#include "../src/fxfile/recycle_metadata.h"
#include "../src/fxfile/operation_path_pairs.h"
#include <vector>

static int calls, failAt, ended, fallback, beginCalls;
static bool failBegin, failEnd;
static HDWP last;
BOOL fakeGetWindowRect(HWND, LPRECT p) { *p = RECT{0,0,10,10}; return TRUE; }
HWND fakeGetParent(HWND) { return reinterpret_cast<HWND>(1); }
int fakeMapWindowPoints(HWND, HWND, LPPOINT, UINT) { return 0; }
HDWP fakeBeginDeferWindowPos(int) { ++beginCalls; last = reinterpret_cast<HDWP>(10); return failBegin ? NULL : last; }
HDWP fakeDeferWindowPos(HDWP batch, HWND, HWND, int, int, int, int, UINT flags)
{
    assert(batch == last); assert(flags == (SWP_NOZORDER | SWP_NOACTIVATE));
    if (++calls == failAt) return NULL;
    last = reinterpret_cast<HDWP>(static_cast<INT_PTR>(10 + calls)); return last;
}
BOOL fakeEndDeferWindowPos(HDWP batch) { assert(batch == last); ++ended; return !failEnd; }
BOOL fakeSetWindowPos(HWND, HWND, int, int, int, int, UINT flags)
{ assert(flags == (SWP_NOZORDER | SWP_NOACTIVATE)); ++fallback; return TRUE; }
#define GetWindowRect fakeGetWindowRect
#define GetParent fakeGetParent
#define MapWindowPoints fakeMapWindowPoints
#define BeginDeferWindowPos fakeBeginDeferWindowPos
#define DeferWindowPos fakeDeferWindowPos
#define EndDeferWindowPos fakeEndDeferWindowPos
#define SetWindowPos fakeSetWindowPos
#include "../src/fxfile/window_layout_batch.h"

int main()
{
    unsigned tests = 0;
    fxfile::OperationPathPairs<wchar_t> paths;
    paths.add(L"D:\\$Recycle.Bin\\folder", L"D:\\test\\한글 폴더");
    paths.add(L"D:\\$Recycle.Bin\\file", L"D:\\test\\alpha.txt");
    const wchar_t *source=paths.sources.c_str(), *target=paths.targets.c_str();
    assert(wcscmp(source,L"D:\\$Recycle.Bin\\folder")==0);
    assert(wcscmp(target,L"D:\\test\\한글 폴더")==0);
    source+=wcslen(source)+1;target+=wcslen(target)+1;
    assert(wcscmp(source,L"D:\\$Recycle.Bin\\file")==0);
    assert(wcscmp(target,L"D:\\test\\alpha.txt")==0);
    source+=wcslen(source)+1;target+=wcslen(target)+1;
    assert(*source==0 && *target==0);
    assert(paths.sources[paths.sources.size()-1]==0 && paths.sources.c_str()[paths.sources.size()]==0);
    ++tests;
    for (int rows=1; rows<=2; ++rows) for(int cols=1; cols<=3; ++cols)
    for(int width=1; width<1100; width+=7) for(int height=1; height<900; height+=11)
    {
        fxfile::EqualSplitGeometry g;
        const bool ok = fxfile::calculateEqualSplitGeometry(width,height,rows,cols,4,g);
        assert(ok == (width-4*(cols-1)>=cols && height-4*(rows-1)>=rows));
        if(ok)
        {
            int total=4*(cols-1), lo=INT_MAX, hi=0;
            for(int c=0;c<cols;++c){total+=g.columnSizes[c];lo=(g.columnSizes[c]<lo)?g.columnSizes[c]:lo;hi=(g.columnSizes[c]>hi)?g.columnSizes[c]:hi;}
            assert(total==width && hi-lo<=1);
            total=4*(rows-1); for(int r=0;r<rows;++r)total+=g.rowSizes[r]; assert(total==height);
            int i=0;for(int c=0;c<cols-1;++c,++i)assert(static_cast<int>(width*g.ratios[i]+0.5)==g.columnSizes[c]);
            for(int r=0;r<rows-1;++r,++i)assert(static_cast<int>(height*g.ratios[i]+0.5)==g.rowSizes[r]);
        }
        ++tests;
    }
    fxfile::EqualSplitGeometry edge;
    assert(fxfile::calculateEqualSplitGeometry(INT_MAX,INT_MAX,2,3,4,edge));
    assert(!fxfile::calculateEqualSplitGeometry(100,100,3,3,4,edge));
    for(uint64_t version=1; version<=2; ++version)
    {
        std::wstring path=L"D:\\test\\한글 파일.txt";
        size_t offset=version==2?28:24;
        std::vector<unsigned char> bytes(offset+(path.size()+1)*2,0);
        uint64_t fileSize=UINT64_C(9000000000),time=123456789;
        std::memcpy(&bytes[0],&version,8);std::memcpy(&bytes[8],&fileSize,8);std::memcpy(&bytes[16],&time,8);
        uint32_t chars=static_cast<uint32_t>(path.size()+1);
        if(version==2)std::memcpy(&bytes[24],&chars,4);
        std::memcpy(&bytes[offset],path.c_str(),chars*2);
        fxfile::RecycleMetadata parsed;
        assert(fxfile::parseRecycleMetadata(&bytes[0],bytes.size(),parsed));
        assert(parsed.originalPath==path && parsed.originalSize==fileSize && parsed.deletedTime==time);
        for(size_t n=0;n<bytes.size();++n)assert(!fxfile::parseRecycleMetadata(&bytes[0],n,parsed));
        uint64_t badVersion=3;std::memcpy(&bytes[0],&badVersion,8);
        assert(!fxfile::parseRecycleMetadata(&bytes[0],bytes.size(),parsed));
        ++tests;
    }
    for(int mode=0;mode<7;++mode)
    {
        calls=ended=fallback=beginCalls=0;failBegin=(mode==4);failEnd=(mode==5);failAt=(mode>=1&&mode<=3)?mode:0;
        fxfile::WindowLayoutBatch batch;
        RECT rect={0,0,mode==6?10:20,10};
        for(int i=1;i<=3;++i) batch.add(reinterpret_cast<HWND>(static_cast<INT_PTR>(i)),rect);
        assert(batch.apply());
        if(mode==0)assert(calls==3 && ended==1 && fallback==0);
        else if(mode<=4)assert(ended==0 && fallback==3);
        else if(mode==5)assert(ended==1 && fallback==3);
        else assert(calls==0 && ended==0 && fallback==0 && beginCalls==0);
        ++tests;
    }
    std::printf("PASS geometry and native batch scenarios: %u\n", tests);
}
