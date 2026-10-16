# CMake 빌드 시스템 가이드 (Modernization Guide)

본 문서는 fxfile 프로젝트의 레거시 GYP 빌드 시스템에서 현대적인 CMake 빌드 시스템으로의 전환 결과와 사용 방법을 상세히 기술합니다.

---

## 1. 전환 배경 및 결과
2013년에 수립된 GYP 시스템은 Python 2.7 의존성 및 한글 경로 빌드 오류(C1041) 등 현대적인 Windows 11 환경에서 여러 한계를 드러냈습니다. 이를 해결하기 위해 **CMake 기반의 현대적 빌드 엔진**으로 전면 교체되었습니다.

### 주요 개선 사항
- **탈(脫) Python 의존성**: 더 이상 Python 설치나 복잡한 환경 변수 설정이 필요 없습니다.
- **무결점 빌드 (Path Guard)**: 빌드 경로에 한글이나 특수문자가 포함될 경우 발생하던 고질적인 PDB 잠금 문제를 사전에 감지하고 차단합니다.
- **아키텍처 인지형 빌드**: x64 및 x32 환경을 자동으로 구분하여 그에 맞는 외부 라이브러리(libxml2, GFL 등)를 링크하며, 빌드 타겟에 따라 `bin/x64` 또는 `bin/x32` 폴더로 자동 배포됩니다.
- **자동화된 배포 패키징**: 빌드 완료 즉시 필요한 모든 DLL과 리소스 파일이 `bin/$(Arch)` 폴더로 자동 수집됩니다.

---

## 2. 사용 방법 (One-Click Build)

### 2.1 사전 준비 사항 (Prerequisites)
빌드 엔진 구동을 위해 다음 도구 중 하나를 통해 CMake가 반드시 설치되어 있어야 합니다.

1.  **Visual Studio Installer 활용 (권장)**:
    - [Visual Studio Installer] 실행 → 사용 중인 VS 버전 [수정] 클릭.
    - [C++를 사용한 데스크톱 개발] 항목 선택.
    - 우측 [세부 정보]에서 **"Windows용 C++ CMake 도구"** 체크 후 설치.
2.  **독립형 설치 (Standalone)**:
    - [cmake.org](https://cmake.org/download/)에서 Windows x64 MSI 다운로드 및 설치.
    - 설치 과정에서 `Add CMake to the system PATH` 옵션을 선택하십시오.

### 2.2 빌드 단계 (x64 또는 x32 실행)
> **후속 정정(Task 057): 아래 `AutoBuild-And-Cleanup.ps1` 방식은 폐기되어 실행 시 실패합니다.** D: 루트 샌드박스, 단일 아키텍처, 수동 `bin` 복사는 현재 세 패키지 롤백/manifest와 드라이브·TEMP 하드게이트를 보장하지 못합니다. 현재 명령은 `preflight_build_environment.bat` 통과 후 `build_deploy_all.bat`입니다. 아래 내용은 역사적 참고로만 보존합니다.

1. `[역사적 절차—실행 금지]` `fxfile_working` 폴더 내에서 `tools\AutoBuild-And-Cleanup.ps1` 스크립트를 실행합니다.
   ```powershell
   # 64비트 빌드
   .\tools\AutoBuild-And-Cleanup.ps1 -Arch x64
   # 32비트 빌드
   .\tools\AutoBuild-And-Cleanup.ps1 -Arch x32
   ```
2. 빌드가 완료되면 `bin\x64` (또는 `bin\x32`) 폴더에 모든 실행 파일과 라이브러리가 모입니다.
3. 배포 폴더(`fxfile_run_x64` 또는 `fxfile_run_x32`)에서 실행 무결성을 확인합니다.

> **[후속 정정]** `build_master.bat` 직접 구동과 `AutoBuild-And-Cleanup.ps1` 사용은 모두 금지/차단된다. `build_deploy_all.bat`만 현재 통합 진입점이다.

### 2.3 Visual Studio에서 작업하기
- `build_cmake` 폴더 내의 **`fxfile_root.sln`** 파일을 열어서 이전과 동일하게 개발 및 디버깅을 진행할 수 있습니다.

---

## 3. 프로젝트 구조 및 관리

### 3.1 주요 디렉토리 구성
- **`/CMakeLists.txt`**: 루트 설정 (전역 컴파일 옵션, MFC 활성화).
- **`/src/base`**: 공통 소스 코드 (Object Library 방식).
- **`/src/xpr`, `/src/fxfile-crash`, `/src/fxfile-keyhook`**: 핵심 DLL 모듈.
- **`/src/fxfile`, `/src/fxfile-launcher`, `/src/fxfile-upchecker`**: 메인 실행 파일, 런처 및 업데이트 체크 모듈.

### 3.2 신규 파일 추가 시
새로운 `.cpp` 또는 `.h` 파일을 추가할 경우, 해당 모듈 폴더 내의 `CMakeLists.txt` 파일의 `SOURCES` 리스트에 파일명을 추가하면 즉시 반영됩니다.

---

## 4. 문제 해결 (Troubleshooting)

| 현상 | 원인 | 해결 방법 |
| :--- | :--- | :--- |
| **Path Guard 경고** | 현재 경로에 한글이 포함됨 | 프로젝트를 `D:\fxfile_build`와 같은 순수 영문 경로로 옮겨서 빌드하십시오. |
| **DLL 로드 실패** | 외부 라이브러리 누락 또는 아키텍처 불일치 | `collect_artifacts` 타겟이 정상 작동했는지 확인하고 `bin/x64` 또는 `bin/x32`에 DLL이 있는지 확인하십시오. x32 빌드 시관 MinGW 런타임 DLL(`libgcc_s_sjlj-1.dll` 등) 누락 여부도 확인하십시오. |
| **C1041 오류** | PDB 서버 충돌 | `/MP` 옵션과 ASCII 경로 가드가 적용되어 있으나, 기존 `obj` 폴더를 삭제 후 재시도하십시오. |

---

> [!IMPORTANT]
> 이제 fxfile은 2026년 표준 개발 환경에 최적화되었습니다. 본 가이드를 준수하여 무결점 빌드 환경을 유지하시기 바랍니다.

**— 가이드 최종 수립 (2026-02-21) —**
