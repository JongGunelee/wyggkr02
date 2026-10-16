# src/fxfile-crash/CMakeLists.txt

project(fxfile-crash)

# 1. 소스 파일 정의
file(GLOB FX_CRASH_SOURCES *.cpp *.h)
list(APPEND FX_CRASH_SOURCES BugTrap.def BugTrap.rc)

# 2. 타겟 생성 (DLL)
add_library(fxfile-crash SHARED ${FX_CRASH_SOURCES})

# 3. 컴파일 정의 및 인클루드 경로
target_compile_definitions(fxfile-crash PRIVATE BUGTRAP_EXPORTS FXFILE_CRASH_DL_EXPORTS)
target_include_directories(fxfile-crash PRIVATE 
    ./
    ${CMAKE_SOURCE_DIR}/lib/zlib/include
)

# 4. PCH 설정
target_precompile_headers(fxfile-crash PRIVATE stdafx.h)

# 5. 아키텍처별 zlib 경로 설정
if(FX_ARCH STREQUAL "x64")
    set(ZLIB_LIB_NAME "zlib-x64.lib")
else()
    set(ZLIB_LIB_NAME "zlib.lib")
endif()

find_library(ZLIB_LIB NAMES ${ZLIB_LIB_NAME} PATHS "${CMAKE_SOURCE_DIR}/lib/zlib/lib" NO_DEFAULT_PATH)
target_link_directories(fxfile-crash PRIVATE "${CMAKE_SOURCE_DIR}/lib/zlib/lib")

target_link_libraries(fxfile-crash PRIVATE 
    ${ZLIB_LIB}
    shlwapi.lib
    comctl32.lib
    ws2_32.lib
    version.lib
    wininet.lib
)

# 6. 출력 파일명 설정
set_target_properties(fxfile-crash PROPERTIES 
    OUTPUT_NAME "fxfile-crash"
    PREFIX ""
)
