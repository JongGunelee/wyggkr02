# src/xpr/CMakeLists.txt

project(xpr)

# 1. 소스 파일 정의
set(XPR_SOURCES
    xpr/xpr_atomic_win.cpp
    xpr/xpr_char.cpp
    xpr/xpr_char_set.cpp
    xpr/xpr_cstr.cpp
    xpr/xpr_debug_win.cpp
    xpr/xpr_dllmain_win.cpp
    xpr/xpr_env.cpp
    xpr/xpr_file_io_win.cpp
    xpr/xpr_file_path_win.cpp
    xpr/xpr_file_sys_win.cpp
    xpr/xpr_guid.cpp
    xpr/xpr_hash.cpp
    xpr/xpr_hash_internal.h
    xpr/xpr_hash_knuth_multiplicative.cpp
    xpr/xpr_hash_knuth_multiplicative.h
    xpr/xpr_hash_pseudo_randomizing.cpp
    xpr/xpr_hash_pseudo_randomizing.h
    xpr/xpr_init.cpp
    xpr/xpr_lazy_library_load_win.cpp
    xpr/xpr_lazy_library_load.h
    xpr/xpr_memory.cpp
    xpr/xpr_process_sync_win.cpp
    xpr/xpr_process_win.cpp
    xpr/xpr_rcode.cpp
    xpr/xpr_shm_win.cpp
    xpr/xpr_string.cpp
    xpr/xpr_string_iterator.cpp
    xpr/xpr_string_tokenizer.cpp
    xpr/xpr_system_win.cpp
    xpr/xpr_text_file_io.cpp
    xpr/xpr_mutex_win.cpp
    xpr/xpr_thread_win.cpp
    xpr/xpr_time_win.cpp
    xpr/xpr_trace_logger.cpp
    xpr/xpr_ustring.cpp
    xpr/xpr_ver.cpp
    xpr/xpr_xml.cpp
    mingw_compat.c
)

set(XPR_HEADERS
    include/xpr.h
    include/xpr_atomic.h
    include/xpr_bit.h
    include/xpr_buffer.h
    include/xpr_char.h
    include/xpr_char_set.h
    include/xpr_config.h
    include/xpr_config_win.h
    include/xpr_cond_variable.h
    include/xpr_cstr.h
    include/xpr_debug.h
    include/xpr_define.h
    include/xpr_dlsym.h
    include/xpr_endian.h
    include/xpr_env.h
    include/xpr_file_io.h
    include/xpr_file_path.h
    include/xpr_file_sys.h
    include/xpr_guid.h
    include/xpr_hash.h
    include/xpr_init.h
    include/xpr_math.h
    include/xpr_memory.h
    include/xpr_memory_barrier.h
    include/xpr_net_io.h
    include/xpr_process.h
    include/xpr_process_sync.h
    include/xpr_rcode.h
    include/xpr_rw_lock.h
    include/xpr_scoped_ptr.h
    include/xpr_shm.h
    include/xpr_spin_lock.h
    include/xpr_std.h
    include/xpr_string.h
    include/xpr_string_iterator.h
    include/xpr_string_tokenizer.h
    include/xpr_system.h
    include/xpr_tchar.h
    include/xpr_text_file_io.h
    include/xpr_thread.h
    include/xpr_time.h
    include/xpr_trace_logger.h
    include/xpr_types.h
    include/xpr_uri.h
    include/xpr_ustring.h
    include/xpr_ver.h
    include/xpr_xml.h
)

# 2. 타겟 생성 (DLL)
add_library(xpr SHARED ${XPR_SOURCES} ${XPR_HEADERS})

# 3. 컴파일 정의 및 인클루드 경로
target_compile_definitions(xpr PRIVATE XPR_CFG_DL_EXPORT)
target_include_directories(xpr PUBLIC 
    include
    ${CMAKE_SOURCE_DIR}/src/xpr/include
    ${CMAKE_SOURCE_DIR}/lib/iconv/include
    ${CMAKE_SOURCE_DIR}/lib/libxml2/include/libxml2
)

# 4. 아키텍처별 라이브러리 경로 설정 및 링크
if(FX_ARCH STREQUAL "x64")
    set(LIBXML2_DIR "${CMAKE_SOURCE_DIR}/lib/libxml2/lib64")
    set(ZLIB_DIR "${CMAKE_SOURCE_DIR}/lib/zlib/lib")
    set(ICONV_DIR "${CMAKE_SOURCE_DIR}/lib/iconv/lib64")
    set(ZLIB_NAME "zlib-x64")
else()
    set(LIBXML2_DIR "${CMAKE_SOURCE_DIR}/lib/libxml2/lib")
    set(ZLIB_DIR "${CMAKE_SOURCE_DIR}/lib/zlib/lib")
    set(ICONV_DIR "${CMAKE_SOURCE_DIR}/lib/iconv/lib")
    set(ZLIB_NAME "zlib")
endif()

find_library(LIBXML2_LIB NAMES libxml2.dll.a libxml2 libxml2_a PATHS ${LIBXML2_DIR} NO_DEFAULT_PATH)
find_library(ZLIB_LIB NAMES ${ZLIB_NAME}.lib ${ZLIB_NAME} libz.dll.a libz PATHS ${ZLIB_DIR} NO_DEFAULT_PATH)
find_library(ICONV_LIB NAMES libiconv.dll.a libiconv PATHS ${ICONV_DIR} NO_DEFAULT_PATH)

target_link_libraries(xpr PRIVATE 
    ${LIBXML2_LIB} 
    ${ZLIB_LIB} 
    ${ICONV_LIB}
    ws2_32 
    winmm
    legacy_stdio_definitions.lib
)

# 5. 출력 파일명 설정 (GYP: libxprw.dll / libxpr.dll)
# 유니코드 버전에 따라 명칭 변경 로직이 있으나 우선 기본 xpr로 진행 후 필요시 조정
set_target_properties(xpr PROPERTIES 
    OUTPUT_NAME "libxprw"
    PREFIX ""
)
