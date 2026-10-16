# src/base/CMakeLists.txt

project(base)

# 1. 모든 소스 파일 정의 (OBJECT 라이브러리로 활용)
file(GLOB BASE_SOURCES *.cpp *.c *.h)

# 2. 타겟 생성 (OBJECT)
add_library(base OBJECT ${BASE_SOURCES})

# 3. 인클루드 경로 설정 (타 모듈 참조용)
target_include_directories(base PUBLIC ./)
target_compile_definitions(base PRIVATE XPR_CFG_OS_WINDOWS)
target_include_directories(base PRIVATE 
    ${CMAKE_SOURCE_DIR}/src/xpr/include
)

# 4. PCH 설정
target_precompile_headers(base PRIVATE 
    $<$<COMPILE_LANGUAGE:CXX>:"${CMAKE_CURRENT_SOURCE_DIR}/stdafx.h">
)
