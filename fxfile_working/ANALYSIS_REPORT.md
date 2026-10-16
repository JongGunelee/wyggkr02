# fxfile 오픈소스 프로젝트 심층분석 보고서
## Windows 11 호환성 및 메모리 최적화

**분석일**: 2026-02-11  
**대상 리포지토리**: https://github.com/fxfile/fxfile  
**분석 범위**: 전체 소스코드 (237 커밋, 1800+ 파일)  
**작업 디렉토리**: `C:\Users\PC\Downloads\0000 FxFile\`

---

## 1. 프로젝트 개요

| 항목 | 내용 |
|------|------|
| 이름 | fxfile (구 flyExplorer) |
| 유형 | Windows 전용 파일 관리자 |
| 언어 | C++ (72.5%), C (14.3%), NSIS (11.8%) |
| 프레임워크 | MFC (Microsoft Foundation Classes) |
| 빌드 시스템 | GYP (Chromium 기반) |
| 라이선스 | GPLv3 |
| 최초 배포 | 2013년 |
| 마지막 커밋 | 2016-05-02 |

---

## 2. 아키텍처 분석

### 2.1 프로젝트 구조
```
fxfile/
├── build/          - 빌드 설정 파일
├── dist/           - 배포(인스톨러) 관련
├── docs/htmlhelp/  - HTML 도움말 문서
├── lib/            - 외부 라이브러리 (gfl, gtest, iconv, libcurl, libxml2, libxslt, mingwrt, vld, zlib)
├── src/
│   ├── base/       - 기본 유틸리티 (40 파일)
│   ├── fxfile/     - 핵심 애플리케이션 (215+ 파일, 5 서브디렉토리)
│   ├── fxfile-crash/    - 크래시 핸들러
│   ├── fxfile-keyhook/  - 키보드 후킹
│   ├── fxfile-launcher/ - 런처 프로그램
│   ├── fxfile-upchecker/- 업데이트 체커
│   └── xpr/        - XPR 크로스 플랫폼 런타임
├── build.bat       - MSVS 솔루션 생성 배치
└── fxfile.gyp      - GYP 빌드 구성
```

---

## 3. 발견 및 수정된 문제들

### 🔴 심각 (Critical) - 모두 수정 완료

#### 3.1 ✅ Windows 버전 타겟팅 (targetver.h)
- **문제**: Windows XP(0x0501)/98(0x0410) 타겟으로 설정
- **영향**: Windows 10/11의 최신 API 사용 불가
- **수정**: Windows 10(0x0A00)으로 업데이트, NTDDI_VERSION 추가

#### 3.2 ✅ GetVersionEx Deprecated (SystemInfo.cpp, xpr_system_win.cpp)
- **문제**: Windows 8.1부터 deprecated된 GetVersionEx 사용
- **영향**: Windows 10/11에서 잘못된 버전 정보 반환 (6.2로 고정)
- **수정**: RtlGetVersion (ntdll.dll)으로 교체, 올바른 Win11 버전 감지

#### 3.3 ✅ VirtualAlloc 고정 주소 할당 (SystemInfo.cpp)
- **문제**: 고정 주소(0x100000)에 메모리 할당 시도
- **영향**: 64비트 Windows 11의 ASLR로 인해 할당 실패 → 크래시
- **수정**: NULL로 변경, MEM_RESERVE|MEM_COMMIT으로 적절한 메모리 관리

#### 3.4 ✅ 64비트 구조체 미스매치 (SystemInfo.h)
- **문제**: VM_COUNTERS, SYSTEM_HANDLE이 32비트 DWORD만 사용
- **영향**: 64비트 Windows에서 포인터/크기 불일치 → 메모리 접근 오류
- **수정**: SIZE_T, ULONG_PTR 사용으로 64비트 호환성 확보

#### 3.5 ✅ XPR 메모리 함수 로직 버그 (xpr_memory.cpp)
- **문제**: `GetProcessHeap()` 반환값 검사 조건이 **반전**됨!
  - `if (sHeap != XPR_NULL) return error;` ← 성공 시 에러를 반환!
- **영향**: 모든 xpr_malloc/xpr_calloc/xpr_realloc이 실패, 또는 NULL 힙으로 할당 → **메모리 손상 및 크래시의 주요 원인**
- **수정**: `if (sHeap == XPR_NULL)` 으로 조건문 정정

#### 3.6 ✅ OS 버전 상수 미정의 (xpr_system.h)
- **문제**: Windows 8.1/10/11 버전 상수가 없어 `kOsVerWinHigher`로 분류
- **영향**: Windows 버전 기반 기능 분기가 올바르게 작동하지 않음
- **수정**: Win8.1~Win11(24H2) 및 Win Server 2022 상수 추가

#### 3.7 ✅ UNICODE_STRING 재정의 충돌 (SystemInfo.h)
- **문제**: Windows SDK에 이미 정의된 UNICODE_STRING을 재정의
- **영향**: 헤더 포함 순서에 따라 컴파일 오류 또는 크기 불일치
- **수정**: `#ifndef _UNICODE_STRING_DEFINED` 가드 추가

### 🟡 경고 (Warning) - 수정 완료

#### 3.8 ✅ InitCommonControlsEx 불완전 (win_app.cpp)
- **문제**: ICC_WIN95_CLASSES만 초기화
- **수정**: 전체 현대 컨트롤 클래스 플래그 추가

#### 3.9 ✅ DPI 인식 미설정 (win_app.cpp)
- **문제**: 고해상도 디스플레이에서 UI가 흐릿하게 렌더링
- **수정**: Per-Monitor DPI Awareness V2 동적 설정 추가

#### 3.10 ✅ 스레드 핸들 누수 (xpr_thread_win.cpp)
- **문제**: Thread::join()에서 NULL 핸들에 대해 CloseHandle 호출
- **수정**: NULL 체크 내부로 CloseHandle 이동

#### 3.11 ✅ 프로세스 정보 버퍼 크기 부족 (SystemInfo.h)
- **문제**: BufferSize가 64KB로 고정 - 현대 Windows 시스템에 부족
- **수정**: 512KB로 증가

#### 3.12 ✅ 스레드 타임아웃 너무 짧음 (SystemInfo.cpp)
- **문제**: GetFileNameThread의 WaitForSingleObject 타임아웃이 100ms
- **수정**: 500ms로 증가, 주석 개선

#### 3.13 ✅ 안전하지 않은 CRT 함수 경고 (stdafx.h)
- **문제**: _tcscpy, _tcscat 등 deprecated 함수 사용 시 컴파일 경고
- **수정**: _CRT_SECURE_NO_WARNINGS 추가 (점진적 마이그레이션 지원)

#### 3.14 ✅ 현대 헤더 미포함 (stdafx.h)
- **문제**: VersionHelpers.h, Shellapi.h 미포함
- **수정**: 필요 헤더 추가

---

## 4. 수정 이력 요약

| # | 파일 | 수정 유형 | 상태 |
|---|------|----------|------|
| 1 | `src/fxfile/targetver.h` | Win10/11 타겟 업데이트 | ✅ 완료 |
| 2 | `src/fxfile/SystemInfo.cpp` | GetVersionEx→RtlGetVersion, VirtualAlloc 고정주소 제거, 타임아웃 증가 | ✅ 완료 |
| 3 | `src/fxfile/SystemInfo.h` | UNICODE_STRING 가드, 64비트 구조체, 버퍼 크기 | ✅ 완료 |
| 4 | `src/fxfile/win_app.cpp` | Common Controls 확장, DPI 인식 | ✅ 완료 |
| 5 | `src/fxfile/stdafx.h` | _CRT_SECURE_NO_WARNINGS, VersionHelpers.h | ✅ 완료 |
| 6 | `src/xpr/include/xpr_system.h` | Win8.1~Win11 버전 상수 추가 | ✅ 완료 |
| 7 | `src/xpr/xpr/xpr_system_win.cpp` | RtlGetVersion 도입, Win10/11 감지 로직 | ✅ 완료 |
| 8 | `src/xpr/xpr/xpr_memory.cpp` | **GetProcessHeap 조건 반전 버그 수정** | ✅ 완료 |
| 9 | `src/xpr/xpr/xpr_thread_win.cpp` | NULL 핸들 CloseHandle 방지 | ✅ 완료 |

---

## 5. 백업 및 작업 환경

```
C:\Users\PC\Downloads\0000 FxFile\
├── fxfile_original_backup/   ← 원본 백업 (변경 불가)
└── fxfile_working/           ← 작업 사본 (수정 완료)
```

---

## 6. 추가 고려사항 (향후 작업)

### 안전하지 않은 문자열 함수
- `_tcscpy`, `_tcscat`, `_stprintf` 등이 **50개 이상 파일**에서 사용됨
- 점진적으로 `_s` 접미사 안전 버전으로 마이그레이션 권장
- 현재 `_CRT_SECURE_NO_WARNINGS`로 컴파일 경고만 억제

### 외부 라이브러리 업데이트
- `lib/` 디렉토리의 라이브러리들이 2013~2016년 버전
- 보안 취약점이 있을 수 있으나, 빌드 환경 종속성으로 인해 신중한 업데이트 필요

### 빌드 시스템 현대화
- GYP 빌드 시스템은 더 이상 유지보수되지 않음
- CMake로 마이그레이션하면 최신 Visual Studio와의 호환성 향상

---

## 7. 메모리 문제 근본 원인 분석

사용자가 보고한 **메모리 문제와 트러블의 주요 원인**은 다음과 같이 분석됩니다:

### 🔴 최우선 원인: XPR 메모리 할당 반전 버그 (#3.5)
`xpr_memory.cpp`에서 `GetProcessHeap()` 성공/실패 조건이 **반대**로 되어 있어,
Windows 힙 API를 통한 **모든 메모리 할당이 실패**하거나 **손상된 힙**으로 할당되었음.
이는 가장 기본적인 메모리 관리 계층에서 발생하는 버그로, 
전체 애플리케이션의 불안정성을 야기하는 **근본 원인**.

### 🟡 보조 원인: ASLR 비호환 (#3.3)
`VirtualAlloc`의 고정 주소(0x100000) 할당이 64비트 Windows 11에서 실패하여
프로세스 정보 조회가 불가능해지고, 이후 NULL 포인터 접근으로 크래시 발생.

### 🟡 보조 원인: 64비트 구조체 불일치 (#3.4)
32비트 크기의 구조체로 64비트 OS 정보를 읽으면 데이터 오프셋이 어긋나서
**잘못된 메모리 영역**을 읽고 쓰게 됨 → 메모리 손상.
