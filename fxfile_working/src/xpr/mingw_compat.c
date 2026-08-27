// src/xpr/mingw_compat.c
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#include <stdio.h>
#include <stdarg.h>

/**
 * MinGW compatibility shim for MSVC.
 * This file provides definitions for symbols expected by MinGW-built static libraries (like libxml2.a).
 */

// 1. Stdio redirection
int __ms_vsnprintf(char* buffer, size_t count, const char* format, va_list argptr) {
    return _vsnprintf(buffer, count, format, argptr);
}

// 2. Stack probe shim
// MSVC uses __chkstk, MinGW uses ___chkstk_ms.
#if defined(_M_X64) || defined(__x86_64__)
void ___chkstk_ms(void) {
    // No-op for x64 usually suffices if the stack usage isn't extreme
}
#else
void ___chkstk_ms(void) {
    // x86 implementation (if needed)
}
#endif

// 3. __iob_func for MSVC 2015+ compatibility
// MinGW libraries often reference __iob_func for stdin/stdout/stderr access.
#if _MSC_VER >= 1900
FILE * __cdecl __iob_func(void) {
    static FILE _iobs[3];
    static int initialized = 0;
    if (!initialized) {
        _iobs[0] = *stdin;
        _iobs[1] = *stdout;
        _iobs[2] = *stderr;
        initialized = 1;
    }
    return _iobs;
}

// Map the import symbol as well
void* __imp___iob_func = __iob_func;
#endif

// 4. Wspiapi shims for libxml2.a
int WspiapiGetAddrInfo(const char* nodename, const char* servname, const struct addrinfo* hints, struct addrinfo** res) {
    return getaddrinfo(nodename, servname, hints, res);
}

void WspiapiFreeAddrInfo(struct addrinfo* ai) {
    freeaddrinfo(ai);
}

// 5. Missing standard functions referenced by name in some MinGW objects
int fprintf_shim(FILE* stream, const char* format, ...) {
    va_list args;
    va_start(args, format);
    int res = vfprintf(stream, format, args);
    va_end(args);
    return res;
}

int sscanf_shim(const char* buffer, const char* format, ...) {
    va_list args;
    va_start(args, format);
    int res = vsscanf(buffer, format, args);
    va_end(args);
    return res;
}
