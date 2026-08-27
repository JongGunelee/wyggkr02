# src/fxfile/CMakeLists.txt

project(fxfile)

# 1. 소스 파일 정의
file(GLOB FXFILE_ROOT_SOURCES *.cpp *.h *.rc)
file(GLOB_RECURSE FXFILE_CFG_SOURCES cfg/*.cpp cfg/*.h)
file(GLOB_RECURSE FXFILE_CMD_SOURCES cmd/*.cpp cmd/*.h)
file(GLOB_RECURSE FXFILE_GUI_SOURCES gui/*.cpp gui/*.h)

set(FXFILE_SOURCES 
    ${FXFILE_ROOT_SOURCES} 
    ${FXFILE_CFG_SOURCES} 
    ${FXFILE_CMD_SOURCES} 
    ${FXFILE_GUI_SOURCES}
)

# 2. 타겟 생성 (MFC EXE)
# CMAKE_MFC_FLAG는 루트에서 2(공유 MFC)로 설정됨
add_executable(fxfile WIN32 ${FXFILE_SOURCES})

# 3. 컴파일 정의 및 인클루드 경로
target_include_directories(fxfile PRIVATE 
    ./
    cfg
    cmd
    cmd/router
    gui
    gui/rebar
    ${CMAKE_SOURCE_DIR}/src/xpr/include
    ${CMAKE_SOURCE_DIR}/src/base
    ${CMAKE_SOURCE_DIR}/src/fxfile-crash
    ${CMAKE_SOURCE_DIR}/lib/gfl/include
    ${CMAKE_SOURCE_DIR}/lib/vld/include
)

# 4. PCH 설정
target_precompile_headers(fxfile PRIVATE stdafx.h)

# 5. 아키텍처별 외부 라이브러리 경로 설정
if(FX_ARCH STREQUAL "x64")
    set(GFL_LIB_DIR "${CMAKE_SOURCE_DIR}/lib/gfl/lib64")
else()
    set(GFL_LIB_DIR "${CMAKE_SOURCE_DIR}/lib/gfl/lib")
endif()

find_library(GFL_LIB NAMES libgfl libgfl.lib PATHS ${GFL_LIB_DIR} NO_DEFAULT_PATH)

# 6. 타겟 링크
target_link_libraries(fxfile PRIVATE 
    base            # OBJECT 라이브러리
    xpr             # DLL 타겟
    fxfile-crash    # DLL 타겟
    ${GFL_LIB}
    shlwapi.lib
    msimg32.lib
    Mpr.lib
    htmlhelp.lib
    Version.lib
    Imm32.lib
    setupapi.lib
    bcrypt.lib
    Rstrtmgr.lib
)

# 7. 출력 파일명 및 디버그/맵 설정
set_target_properties(fxfile PROPERTIES 
    OUTPUT_NAME "fxfile"
    LINK_FLAGS "/MAP /DEBUG"
)
