# src/fxfile-upchecker/CMakeLists.txt

project(fxfile-upchecker)

# 1. 소스 파일 정의
file(GLOB UPCHECKER_SOURCES *.cpp *.h *.rc)

# 2. 타겟 생성 (MFC EXE)
add_executable(fxfile-upchecker WIN32 ${UPCHECKER_SOURCES})

# 3. 컴파일 정의 및 인클루드 경로
target_include_directories(fxfile-upchecker PRIVATE 
    ./
    ${CMAKE_SOURCE_DIR}/src/xpr/include
    ${CMAKE_SOURCE_DIR}/src/base
    ${CMAKE_SOURCE_DIR}/lib/libcurl/include
)

# 4. PCH 설정
target_precompile_headers(fxfile-upchecker PRIVATE stdafx.h)

# 5. 외부 라이브러리 경로 설정
# libcurl은 현재 x86용만 보임 (libcurl.lib)
# x64 환경에서도 동작하도록 libcurl 라이브러리 경로 확보 필요
# 임시로 lib/libcurl/lib 경로 지정
find_library(CURL_LIB NAMES libcurl PATHS "${CMAKE_SOURCE_DIR}/lib/libcurl/lib" NO_DEFAULT_PATH)

# 6. 타겟 링크
target_link_libraries(fxfile-upchecker PRIVATE 
    base
    xpr
    ${CURL_LIB}
    shlwapi.lib
    ws2_32.lib
    winmm.lib
)

# 7. 출력 파일명 설정
set_target_properties(fxfile-upchecker PROPERTIES 
    OUTPUT_NAME "fxfile-upchecker"
)
