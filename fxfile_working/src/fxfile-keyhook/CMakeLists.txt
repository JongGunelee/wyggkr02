# src/fxfile-keyhook/CMakeLists.txt

project(fxfile-keyhook)

# 1. 소스 파일 정의
set(KEYHOOK_SOURCES
    fxfile-keyhook.cpp
    fxfile-keyhook.h
    stdafx.cpp
)

# 2. 타겟 생성 (DLL)
add_library(fxfile-keyhook SHARED ${KEYHOOK_SOURCES})

# 3. 컴파일 정의 및 인클루드 경로
target_compile_definitions(fxfile-keyhook PRIVATE FXFILE_KEYHOOK_EXPORTS)
target_include_directories(fxfile-keyhook PRIVATE ./)

# 4. PCH 설정
target_precompile_headers(fxfile-keyhook PRIVATE stdafx.h)

# 5. 출력 파일명 설정
set_target_properties(fxfile-keyhook PROPERTIES 
    OUTPUT_NAME "fxfile-keyhook"
    PREFIX ""
)
