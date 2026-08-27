cmake_minimum_required(VERSION 3.16)

# 1. 무결성 전환을 위한 심층 분석 반영: 빌드 경로 무결성 가드 (Path Guard)
set(DETECT_NON_ASCII_REGEX "[^ -~]+")
if(CMAKE_SOURCE_DIR MATCHES "${DETECT_NON_ASCII_REGEX}")
    message(WARNING "*******************************************************************************")
    message(WARNING "경보: 현재 소스 경로에 한글 또는 비-ASCII 문자가 포함되어 있습니다.")
    message(WARNING "경로: ${CMAKE_SOURCE_DIR}")
    message(WARNING "비주얼 스튜디오 빌드 도구와의 충돌(C1041)을 방지하기 위해 ASCII 경로 사용을 권장합니다.")
    message(WARNING "*******************************************************************************")
endif()

project(fxfile_root LANGUAGES C CXX)

# 2. 전역 컴파일러 및 링커 옵션 제어 (Modernization)
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 한글 코드 깨짐 방지를 위한 /utf-8 강제
add_compile_options("/utf-8" "/MP")

# MFC 사용 설정 (정적 링크: 1, 공유 링크: 2)
# fxfile.gyp의 UseOfMFC: 2 설정을 따름
set(CMAKE_MFC_FLAG 2)

# 유니코드 전역 설정
add_definitions(-DUNICODE -D_UNICODE)

# 레거시 정적 라이브러리(libxml 등)의 _iob 링크 관련 문제 해결 및 x86 Wspiapi 문제 해결
if(MSVC)
    link_libraries(legacy_stdio_definitions.lib ws2_32.lib)
endif()

# 3. 아키텍처 일관성 (Architecture Consistency)
if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    set(FX_ARCH "x64")
    set(FX_LIB_SUFFIX "64")
else()
    set(FX_ARCH "x32")
    set(FX_LIB_SUFFIX "")
endif()

# 최종 출력 경로 단일화
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY "${CMAKE_SOURCE_DIR}/bin/${FX_ARCH}")
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY "${CMAKE_SOURCE_DIR}/bin/${FX_ARCH}")
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY "${CMAKE_SOURCE_DIR}/obj/lib/${FX_ARCH}")

# 4. 모듈별 타겟 등록
add_subdirectory(src/base)
add_subdirectory(src/xpr)
add_subdirectory(src/fxfile-crash)
add_subdirectory(src/fxfile-keyhook)
add_subdirectory(src/fxfile)
add_subdirectory(src/fxfile-launcher)

# libcurl 64bit 라이브러리 부재로 x64 빌드에서 upchecker 제외
if(NOT FX_ARCH STREQUAL "x64")
    add_subdirectory(src/fxfile-upchecker)
endif()

# 5. 자동화된 빌드 후 작업 (Post-Build DLL & Resource Collection)
# 주요 실행 아키텍처(x64)에 맞춰 외부 DLL과 필수 리소스를 자동 수집합니다.
if(FX_ARCH STREQUAL "x64")
    set(EXTERNAL_DLLS
        "${CMAKE_SOURCE_DIR}/lib/libxml2/bin64/libxml2-2.dll"
        "${CMAKE_SOURCE_DIR}/lib/libxml2/bin64/zlib1.dll"
        "${CMAKE_SOURCE_DIR}/lib/zlib/bin/zlib-x64.dll"
        "${CMAKE_SOURCE_DIR}/lib/iconv/bin64/libiconv-2.dll"
        "${CMAKE_SOURCE_DIR}/lib/iconv/bin64/libcharset-1.dll"
        "${CMAKE_SOURCE_DIR}/lib/gfl/lib64/libgfl340.dll"
        "${CMAKE_SOURCE_DIR}/lib/gfl/lib64/libgfle340.dll"
        "${CMAKE_SOURCE_DIR}/lib/mingwrt/bin64/libgcc_s_sjlj-1.dll"
        "${CMAKE_SOURCE_DIR}/lib/mingwrt/bin64/libwinpthread-1.dll"
    )
elseif(FX_ARCH STREQUAL "x32")
    set(EXTERNAL_DLLS
        "${CMAKE_SOURCE_DIR}/lib/libxml2/bin/libxml2-2.dll"
        "${CMAKE_SOURCE_DIR}/lib/libxml2/bin/zlib1.dll"
        "${CMAKE_SOURCE_DIR}/lib/zlib/bin/zlib.dll"
        "${CMAKE_SOURCE_DIR}/lib/iconv/bin/libiconv-2.dll"
        "${CMAKE_SOURCE_DIR}/lib/iconv/bin/libcharset-1.dll"
        "${CMAKE_SOURCE_DIR}/lib/gfl/lib/libgfl340.dll"
        "${CMAKE_SOURCE_DIR}/lib/gfl/lib/libgfle340.dll"
        "${CMAKE_SOURCE_DIR}/lib/mingwrt/bin/libgcc_s_sjlj-1.dll"
        "${CMAKE_SOURCE_DIR}/lib/mingwrt/bin/libwinpthread-1.dll"
    )
endif()

# Collect the matching Visual C++ and MFC runtimes for app-local deployment.
# A copied portable package must start even when the destination PC does not
# already have the Microsoft Visual C++ Redistributable installed.
if(MSVC)
    set(CMAKE_INSTALL_SYSTEM_RUNTIME_LIBS_SKIP TRUE)
    set(CMAKE_INSTALL_MFC_LIBRARIES TRUE)
    include(InstallRequiredSystemLibraries)
    list(APPEND EXTERNAL_DLLS ${CMAKE_INSTALL_SYSTEM_RUNTIME_LIBS})
endif()

# 배포용 폴더 및 파일 복사 명령 정의
add_custom_target(collect_artifacts ALL
    COMMAND ${CMAKE_COMMAND} -E make_directory "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}/Languages"
    COMMAND ${CMAKE_COMMAND} -E copy_directory "${CMAKE_SOURCE_DIR}/src/fxfile/Languages" "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}/Languages"
    COMMAND ${CMAKE_COMMAND} -E copy ${EXTERNAL_DLLS} "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}/"
    # 빌드된 최신 실행 파일 및 내부 DLL을 루트 배포 폴더로 복사
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:fxfile> "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}/"
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:fxfile-launcher> "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}/"
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:xpr> "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}/"
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:fxfile-crash> "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}/"
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:fxfile-keyhook> "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}/"
    COMMENT "필수 DLL 및 리소스 폴더를 ${CMAKE_RUNTIME_OUTPUT_DIRECTORY}로 수집 중..."
)

# 모든 실행 파일 빌드 후 수집 타겟이 실행되도록 의존성 설정
add_dependencies(collect_artifacts fxfile fxfile-launcher)
