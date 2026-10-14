#pragma once
#include <cstdint>
#include <cstring>
#include <string>

namespace fxfile
{
struct RecycleMetadata
{
    uint64_t originalSize;
    uint64_t deletedTime;
    std::wstring originalPath;
};
inline bool parseRecycleMetadata(const unsigned char *data, size_t size, RecycleMetadata &out)
{
    if (!data || size < 26) return false;
    uint64_t version, originalSize, deletedTime;
    std::memcpy(&version, data, 8);
    std::memcpy(&originalSize, data + 8, 8);
    std::memcpy(&deletedTime, data + 16, 8);
    size_t offset = 24, characters = (size - offset) / 2;
    if (version == 2)
    {
        if (size < 30) return false;
        uint32_t count;
        std::memcpy(&count, data + 24, 4);
        offset = 28; characters = count;
        if (characters < 2 || characters > (size - offset) / 2) return false;
    }
    else if (version != 1) return false;
    std::wstring path;
    bool terminated = false;
    for (size_t i = 0; i < characters; ++i)
    {
        uint16_t ch;
        std::memcpy(&ch, data + offset + i * 2, 2);
        if (!ch) { terminated = true; break; }
        path.push_back(static_cast<wchar_t>(ch));
    }
    if (!terminated || path.size() < 3 || path[1] != L':' || path[2] != L'\\') return false;
    out.originalSize = originalSize;
    out.deletedTime = deletedTime;
    out.originalPath.swap(path);
    return true;
}
}
