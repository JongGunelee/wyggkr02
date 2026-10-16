#pragma once
#include <string>

namespace fxfile
{
// Win32 multi-string arrays require embedded NUL separators and a final NUL.
// xpr::String intentionally ignores appending NUL; never use it for this payload.
template<class Char> struct OperationPathPairs
{
    std::basic_string<Char> sources;
    std::basic_string<Char> targets;
    void add(const Char *source, const Char *target)
    {
        sources.append(source); sources.push_back(Char());
        targets.append(target); targets.push_back(Char());
    }
};
}
