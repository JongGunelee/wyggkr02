// FxFile-local, reversible file/folder operation locks.
#ifndef __FXFILE_FILE_OPERATION_LOCK_STORE_H__
#define __FXFILE_FILE_OPERATION_LOCK_STORE_H__ 1
#pragma once

#include <mutex>
#include <set>
#include <string>
#include <vector>

namespace fxfile
{
class FileOperationLockStore
{
public:
    static FileOperationLockStore &instance(void);

    bool isLocked(const std::wstring &aPath);
    bool affectsLockedPath(const std::wstring &aPath);
    bool hasLocks(void);
    bool setLocked(const std::vector<std::wstring> &aPaths, bool aLocked);
    bool isOperationBlocked(const SHFILEOPSTRUCT *aOperation,
                            std::wstring &aBlockedPath);

private:
    FileOperationLockStore(void);
    void ensureLoaded(void);
    bool save(void);
    std::wstring storagePath(void) const;
    static std::wstring normalize(const std::wstring &aPath);
    static bool isInside(const std::wstring &aPath,
                         const std::wstring &aParent);

private:
    std::mutex mMutex;
    bool mLoaded;
    std::set<std::wstring> mPaths;
};
} // namespace fxfile

#endif // __FXFILE_FILE_OPERATION_LOCK_STORE_H__
