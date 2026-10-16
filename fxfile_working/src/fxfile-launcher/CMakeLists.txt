# src/fxfile-launcher/CMakeLists.txt

project(fxfile-launcher)

# 1. 소스 파일 정의
file(GLOB LAUNCHER_SOURCES *.cpp *.h *.rc)

# 2. 타겟 생성 (MFC EXE)
add_executable(fxfile-launcher WIN32 ${LAUNCHER_SOURCES})

# 3. 컴파일 정의 및 인클루드 경로
target_include_directories(fxfile-launcher PRIVATE 
    ./
    ${CMAKE_SOURCE_DIR}/src/fxfile-keyhook
    ${CMAKE_SOURCE_DIR}/src/base
)

# 4. PCH 설정
target_precompile_headers(fxfile-launcher PRIVATE stdafx.h)

# 5. 타겟 링크
target_link_libraries(fxfile-launcher PRIVATE 
    base
    xpr
    fxfile-keyhook
    shlwapi.lib
)

# 6. 출력 파일명 설정
set_target_properties(fxfile-launcher PROPERTIES 
    OUTPUT_NAME "fxfile-launcher"
)
