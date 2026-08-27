# FxFile 통합 빌드·배포·검증

`build_deploy_all.bat`는 설치본 x64, 포터블 run_x64, 포터블 run_x32를 하나의 통제된 배포 세트로 유지한다.

전체 초보자용 환경 설치·오류 분석·감사·롤백 절차는 작업 루트 `CHANGELOG_HISTORY-1차.md`의 초입 `0.1~0.8`, Task 035와 Task 057을 따른다.

## 매 빌드 전 필수: 드라이브·TEMP/TMP 프리플라이트

순서는 반드시 **드라이브 읽기 전용 검사 → 하드게이트 판정 → D: Task TEMP 생성·쓰기/flush/delete probe → 프로세스 범위 TEMP/TMP 전환 → x64 → 재검사 → x32 → 재검사 → 배포/smoke**다. TEMP/TMP 폴더부터 먼저 만들지 않는다.

하드게이트:

- 기본 경로에서 SystemDrive(C:)는 **5GiB 이상 그리고 5% 이상**이어야 한다. 어느 하나라도 미달하면 configure·빌드·배포·smoke를 시작하지 않는다.
- 권장 C: 상태는 **10GiB 이상 그리고 10% 이상**이다.
- 단일 C: 볼륨 환경(Task 062)에서는 C: 여유가 **10GiB 이상 그리고 10% 이상**일 때 안전한 단일 드라이브 워크플로를 공식 지원한다.
- 다중 볼륨 환경에서 C: 용량이 부족할 때는 Task 059의 감사형 D: 예외를 사용한다.
- 저용량 상태에서 명시적 예외 승인이 없으면 `VerifyOnly`, 소스 정적 감사와 문서 작업만 허용한다.

명시적 저용량 D: 예외 조건:

- 두 인수를 반드시 함께 사용한다: `-AllowLowSystemDriveWithDTemp -LowSystemDriveApproval I_ACCEPT_LOW_SYSTEM_DRIVE_RISK`.
- C:는 전 단계에서 최소 1GiB를 절대 유지한다. 승인된 저용량 workflow의 최초 측정값을 기준으로 전체 누적 감소가 **1GiB를 초과**하면 다음 단계로 진행하지 않는다. 체크포인트 간 변화량과 최초 기준 대비 누적 변화량을 모두 기록한다.
- 1GiB는 Windows·백신·Codex 로그 등 외부 프로세스의 배경 변동을 흡수하는 상한일 뿐이다. 빌드 산출물·캐시·TEMP를 C:에 의도적으로 쓰거나 D: TEMP 계약을 우회하도록 허가하는 용량이 아니다.
- 프로젝트와 Task TEMP/TMP는 실제 고정 로컬 `D:\`에 있어야 하고 20GiB 이상 여유, 비-reparse·비-offline·비클라우드 경계를 통과해야 한다.
- TEMP/TMP는 현재 PowerShell과 그 자식 프로세스에만 적용한다. 사용자·시스템 환경변수는 바꾸지 않는다.
- 승인 값, C:/D: 수치, 단계별 변화량, TEMP 경로, cleanup·rollback 결과를 preflight 보고서와 deployment manifest에 기록한다.

```bat
preflight_build_environment.bat
```

저용량 D: 예외를 사용하도록 사용자가 명시 승인한 경우:

```bat
preflight_build_environment.bat -AllowLowSystemDriveWithDTemp -LowSystemDriveApproval I_ACCEPT_LOW_SYSTEM_DRIVE_RISK
```

이 명령은 최초 설치/도구 업데이트 후뿐 아니라 **매 통합 빌드 직전** 실행한다. 필수 항목 실패가 0건이고 x64/x32 CMake configure 시뮬레이션이 모두 성공해야 한다. 기본 경로에서 C: 하드게이트가 실패하거나, 승인 예외 경로에서 1GiB 절대 하한·D: 경계·20GiB 예약·1GiB 누적 배경변동 상한 중 하나라도 실패하면 D: Task TEMP를 만들거나 configure를 실행하지 않는다. 통과하면 `__BUILD_TEMP_BACKUP__\preflight_날짜_시간\build_temp`를 해당 프로세스와 자식 CMake에만 TEMP/TMP로 사용한다. 환경변수는 항상 복원하고, 새 빌드 프로세스가 0개이며 경계 검사가 통과할 때만 그 Task TEMP를 제거한다. 프로세스·reparse·정리 오류가 남으면 경로와 PID를 보존하고 필수 FAIL로 보고한다. 결과와 로그는 같은 preflight 증거 폴더에 남는다. `.vsconfig`는 VS2022 v143 x64/x86, MFC, CMake 도구, Windows 11 SDK 26100 설치 기준이다.

`-SkipConfigureSimulation`은 원인 조사용 진단 모드일 뿐이며 의도적으로 필수 실패를 남긴다. 이 옵션의 실행 결과로 빌드·배포를 승인할 수 없다.

하드게이트를 통과해 전체 프리플라이트가 실행된 PASS 후보의 `preflight_report.json`에서는 `BuildTempCreated`, `BuildTempProbePassed`, `BuildTempCleanupStatus`, `RemainingBuildProcesses`, `EnvironmentRestored`, `HostPowerShellVersion`, `WorkflowInputs`를 함께 확인한다. 승인 예외에서는 `AllowedSystemDriveDecreaseBytes=1073741824`, 최초/최종 C: 바이트와 누적 변화량도 확인한다. PASS 보고서는 필수 빌드 입력 7개가 정확히 한 번씩 기록되고 현재 SHA-256과 모두 같아야 한다. C: 게이트에서 조기 FAIL한 보고서는 후속 도구 파일을 더 읽지 않고 저장소 수치와 TEMP/configure 생략 사유만 기록한다.

## 기본 실행

모든 FxFile 창과 launcher/upchecker/updater를 먼저 종료하고 위 프리플라이트가 PASS한 직후 실행한다.

```bat
build_deploy_all.bat
```

승인된 저용량 D: 예외 빌드·배포:

```bat
build_deploy_all.bat -AllowLowSystemDriveWithDTemp -LowSystemDriveApproval I_ACCEPT_LOW_SYSTEM_DRIVE_RISK
```

통합 도구는 프리플라이트 결과를 맹신하지 않고 실행 직전 C:/프로젝트 볼륨을 다시 검사한다. 승인 예외에서는 workflow 최초 C: 측정값을 고정 기준으로 삼고, 모든 체크포인트와 `build_master.bat`의 독립 검사에서 `현재 C: >= 최초 C: - 1GiB` 및 절대 1GiB 하한을 함께 강제한다. 통과 후 `__BUILD_TEMP_BACKUP__\build_temp_<시각>_<PID>`를 새로 만들고 프로세스 범위 TEMP/TMP로 설정하며, x64·x32·배포·smoke 경계마다 공간을 재검사한다. 실패/정상 종료 시 원래 TEMP/TMP를 복원하고, 해당 워크플로가 만든 빌드 프로세스가 0개일 때만 Task TEMP를 정리한다.

또한 가장 최근 `preflight_*` 시도 하나가 `PASS`, 필수 실패 0, 실제 x64/x32 configure 실행, 2시간 이내, TEMP probe/정리/환경 복원 성공이어야 한다. 최신 시도가 FAIL·손상·보고서 미생성이면 이전 PASS로 건너뛰지 않는다. preflight가 기록한 통합 스크립트·배치·CMake·`.vsconfig` 해시가 현재와 다르면 다시 프리플라이트한다.

`build_master.bat`를 직접 실행하는 방식은 저장공간 프리플라이트와 통합 롤백을 우회하므로 차단되어 있다. 반드시 `build_deploy_all.bat`를 사용한다.

기본 동작:

1. Release x64 빌드
2. Release x32 빌드
3. PE 아키텍처와 필수 산출물 검증
4. 기존 세 패키지의 설정과 덮어쓸 런타임 파일 백업
5. 동일한 x64 산출물을 설치본과 run_x64에 배포
6. 같은 소스의 x32 산출물을 run_x32에 배포
7. 설치본의 로컬 설정 10개와 launcher 설정을 run_x64/x32에 동기화
8. 루트 `fxfile.ini`·`.fxfile` 부재 검증
9. 산출물에 없는 루트 EXE/DLL을 삭제하지 않고 배포 백업으로 이동
10. EXE/DLL/언어/설정 SHA-256 검증
11. 별도 복제 패키지에서 INI 없는 x64/x32 정상 시작·종료 시험
12. AppData와 설치본 설정이 시험 중 변경되지 않았는지 검증
13. 결과 manifest 저장

성공 manifest에는 `BuildTempRoot`와 단계별 `StorageCheckpoints`가 포함되어 C:/프로젝트 볼륨과 실제 TEMP/TMP를 감사할 수 있다. 저용량 예외의 각 체크포인트에는 직전 대비 변화량과 workflow 최초값 대비 누적 변화량, 1GiB 허용 상한을 함께 기록한다.

백업과 manifest는 작업 루트의 `__BUILD_TEMP_BACKUP__\unified_deploy_날짜_시간`에 생성된다. 배포 또는 시험이 실패하면 덮어쓴 파일은 자동 복원되며, 새로 추가된 파일은 삭제하지 않고 rollback 백업으로 이동한다.

## 기존 빌드 산출물만 배포

```bat
build_deploy_all.bat -Mode DeployVerify
```

## 현재 세 패키지 읽기 전용 검증

```bat
build_deploy_all.bat -Mode VerifyOnly
```

## 동적 실행시험 생략

```bat
build_deploy_all.bat -SkipSmokeTest
```

동적 시험을 생략해도 아키텍처, 필수 파일, 설정·언어·실행 파일 해시는 검증된다.

## 2×2 시작 화면 원자 표시 검증

시작 시 메인 프레임은 먼저 표시되지만, 내부 ExplorerView 네 개는 숨긴 상태에서 프레임 소유의 단일 지연 작업으로 초기화한다. 각 뷰를 Shell/COM 안전성을 지키며 UI 스레드에서 순서대로 준비한 뒤, 메시지 루프로 중간 반환하지 않고 네 패널을 한 번에 표시하고 프레임을 한 번만 다시 그린다. Shell/COM 객체와 HWND를 작업 스레드에서 무리하게 병렬 생성하지 않는다.

통합 smoke 시험은 최종 준비 완료 전 보이는 `SysListView32` 개수를 감시한다. 예상 패널이 4개일 때 1·2·3개가 보인 순간이 관측되면 실패한다. 성공 manifest의 각 `SmokeTests` 항목은 다음 값을 가져야 한다.

```text
ReadyViewCount             = 4
ExpectedViewCount          = 4
AtomicLayoutPublication    = true
PartialVisibleViewCounts   = 빈 배열
ExitCode                   = 0
ForcedTermination          = false
```

`FxFile.StartupLayoutReadyViewCount`는 중간 진행률이 아니라 원자 공개 완료 신호다. 2×2 시작에서는 0에서 4로 직접 바뀐다. 총 시간은 실제 폴더와 Shell 확장 응답에 영향을 받으므로 ‘0초’를 보장하지 않으며, 프레임 골격이 먼저 보이고 네 패널 내용이 준비된 순간 동시에 나타나는 것을 합격 기준으로 삼는다.

흰 화면 방지를 위해 첫 프레임은 실제 패널의 최종 사각형과 저장 경로를 사용한 2×2 골격을 그린다. 골격 paint가 완료되면 `FxFile.StartupLayoutSkeletonPainted=1`을 게시하고, 실제 ExplorerView 네 개는 숨은 상태에서 준비한 뒤 한 번에 교체한다. 네 패널 준비는 첫 `UpdateWindow` 직후 직접 시작해 비핵심 post message보다 우선한다. backward/forward/history PIDL 변환은 최종 공개 뒤 수행하며, 즉시 종료 시에는 저장 전에 강제 완료해 이력을 보존한다.

manifest의 추가 합격값:

```text
SkeletonSeconds          = 첫 2×2 골격 paint 시간
ReadySeconds             = 실제 저장 패널 전체 공개 시간
SkeletonToReadySeconds   = 골격 표시 뒤 실제 내용 교체까지의 시간
```

`SkeletonSeconds`가 없으면 큰 흰 배경 회귀로 간주해 smoke를 실패시킨다. `tools\Measure-ActualShortcutStartup.ps1`도 실제 바로가기 시험에서 `ClickToSkeletonPaintedMs`, `SkeletonPainted`, 최종 4패널 시각을 분리 기록한다.

## 배포 경계

- `bin\x64`와 `bin\x32`의 루트에서 EXE와 DLL만 배포한다.
- `bin\x64`에 남아 있을 수 있는 오래된 `fxfile.ini`, 설정 폴더, PDB, MAP, LIB, EXP는 배포하지 않는다.
- 산출물 목록에 없는 배포 루트의 EXE/DLL은 혼재를 막기 위해 복구 백업으로 이동하고, 세 패키지의 루트 바이너리 목록이 각 아키텍처 산출물과 정확히 같은지 검증한다.
- 설치본 `fxfile` 폴더가 사용자 환경의 정본이다.
- x64와 x32의 설정 파일은 동일하게 유지하되 EXE/DLL은 반드시 아키텍처별로 분리한다.
- 설치본의 updater 하위 폴더와 컴퓨터별 AppCompat 레지스트리는 통합 핵심 런타임 세트 밖이다.
- 다른 컴퓨터에서 절대경로 자산까지 같아지려면 해당 컴퓨터에도 같은 C:/D: 경로와 파일이 있어야 한다.

## 환경 설정의 설정 파일 위치 3개 옵션

`도구 > 환경 설정 > 고급 > 설정 파일`에는 다음 세 가지가 있다.

| 옵션 | 저장 위치 | 용도 |
|---|---|---|
| `%AppData% 폴더(기본값)` | `%AppData%\fxfile\conf` | Windows 사용자 프로필 기반 공유 설정 |
| `프로그램 설치 폴더` | `<fxfile.exe 폴더>\fxfile` | 실행 폴더별 독립·포터블 설정. 세 배포본의 권장값 |
| `사용자 정의 폴더` | 사용자가 선택한 정확한 폴더 | 별도 디스크·동기화 폴더 등 명시적 저장 위치 |

프로그램 설치 폴더 모드는 실행 루트의 `fxfile.ini`·`.fxfile`을 요구하지 않는다. `fxfile\fxfile.conf`와 `fxfile\fxfile-main.conf` 핵심 쌍을 자동 감지하며, 설정 본체는 같은 `fxfile` 폴더에 계속 저장한다.

AppData 또는 사용자 정의 모드는 `%AppData%\fxfile\.fxfile` 포인터를 사용한다. AppData 모드의 포인터 값은 `%AppData%\fxfile\conf`, 사용자 정의 모드는 선택한 절대경로다. 여러 FxFile 복사본이 AppData 포인터를 공유할 수 있으므로 독립 배포 세트는 프로그램 폴더 모드를 사용한다.

설정 위치 변경은 동기화용 복사가 아니라 canonical 설정 파일의 이동이다. 전환 전에 양쪽 폴더를 백업하고, 다른 FxFile 프로세스를 모두 종료하며, 대상 폴더의 쓰기 권한을 확인한다. 전환 후 정상 종료·재시작하여 옵션 선택, 북마크, 도구 모음, 2×2 패널과 각 경로가 유지되는지 확인한다.

### 격리 시뮬레이션 도구

```powershell
$tool = '.\tools\Test-ConfigDirectoryOptions.ps1'
$state = '..\__BUILD_TEMP_BACKUP__\task_config_dir_options_manual'
& $tool -Mode Prepare -StateRoot $state
# sandbox_x64\fxfile.exe에서 한 옵션을 실제 적용하고 정상 종료
& $tool -Mode Audit -StateRoot $state -Expected AppData
& $tool -Mode Audit -StateRoot $state -Expected Custom
& $tool -Mode Audit -StateRoot $state -Expected Program
& $tool -Mode Restore -StateRoot $state
```

`Prepare`는 실제 AppData와 세 운영 패키지의 inventory를 보존하고 시험용 x64 sandbox를 만든다. `Audit`은 Program/AppData/Custom 중 활성 위치, canonical 10개, 포인터 파일과 SHA-256을 기록한다. `Restore`는 실제 AppData를 원복하고 `AppDataInventoryExact=True`, `ProductionPackagesUnchanged=True`를 확인한다.

canonical 설정 10개에는 `fxfile-view_set.conf`와 `fxfile-updater.conf`도 포함된다. 2026-08-12 수정 전 코드는 두 파일을 위치 전환에서 누락했고, 시작 직후 전환하면 지연 로드 전 최근 파일 목록이 손실될 수도 있었다. 현재 코드는 두 파일을 공식 이동 대상에 추가하고 이동 전에 최근 파일 목록을 강제 로드한다.

완전한 실제 시험 결과와 증거 경로는 작업 루트 `CHANGELOG_HISTORY-1차.md`의 Task 042를 참조한다.

## 사용자 화면 기본값과 종료 상태 저장

- 파일 목록 `크기` 열, `도구 > 환경 설정 > 표시 > 용량 표시`의 단일 선택 및 다중 선택 합계 기본 단위는 모두 `바이트`다. 세 설정 키 값은 `SIZE_UNIT_BYTE=10`이다.
- 메인 창 위치·크기 복원은 별도 체크박스가 없는 상시 자동 기능이다.
- 정상 종료 시 `fxfile-main.conf`의 `main.window.position`과 `main.window.status`를 저장하고 다음 시작 시 복원한다.
- 창 배치 API가 실패하거나 0 크기를 반환하면 마지막 정상값을 보존한다.
- Windows 11 좌/우 Snap 상태에서 `WINDOWPLACEMENT.rcNormalPosition`은 Snap 직전의 일반 창 좌표이므로 저장에 사용하지 않는다. `showCmd=SW_SHOWNORMAL`인 일반/Snap 창은 현재 실제 창 사각형을 저장하고, 진짜 최대화/최소화 상태에서만 복원 사각형을 유지한다.
- Snap 창의 보이지 않는 크기 조절 테두리가 작업영역을 몇 픽셀 넘더라도 화면 밖 창으로 오판하지 않도록, 저장 사각형과 작업영역의 실제 교차 여부로 가시성을 판정한다.
- 모니터가 제거되었거나 저장 위치가 모든 현재 모니터에서 완전히 벗어난 경우에만 사용 가능한 작업 영역으로 보정될 수 있다.

현재 창 위치·크기와 2×2 패널, 탭, 북마크 바, 도구 모음 등의 상태를 종료 전에 즉시 저장하려면 `도구(T) > 모든 설정 저장하기(T)`를 사용한다. 전용 기본 단축키는 없지만 `도구 > 단축키 설정`에서 이 명령에 원하는 단축키를 지정할 수 있으며, 지정값은 `fxfile-accel.dat`에 저장된다. `환경 설정` 창의 `적용/확인`은 일반 환경설정만 저장하므로 메인 창 위치를 즉시 저장하는 명령이 아니다.

### 창 위치·크기 영구 잠금과 해제

현재 창 외곽 배치를 이후 실행에서도 고정하려면 다음 순서로 사용한다.

1. `도구(T) > 창 위치·크기 잠금(L)`에 체크가 있으면 먼저 한 번 눌러 해제한다.
2. 창을 원하는 모니터·위치·크기 또는 최대화 상태로 배치한다.
3. 같은 메뉴를 다시 눌러 체크한다.
4. 체크하는 순간 현재 창 위치·크기·상태와 `main.window.position_locked=1`이 `fxfile-main.conf`에 즉시 저장된다.

체크가 있는 동안에는 정상 종료이나 `모든 설정 저장하기`를 실행해도 창 위치·크기만 덮어쓰지 않는다. 2×2 분할, 각 패널 경로, 탭, 북마크, 도구 모음 등 나머지 사용 상태는 계속 정상 저장된다. 다시 자유롭게 마지막 위치를 자동 저장하려면 같은 메뉴를 눌러 체크를 해제한다. 해제 즉시 `main.window.position_locked=0`이 저장되고, 이후 정상 종료 위치가 다음 실행 위치가 된다.

재실행 시 화면상의 위치·크기는 동일하게 복원되지만 Windows 내부의 Snap Group 소속까지 재생하는 것은 보장 범위가 아니다.

### 2×2 창 경로 및 분할 크기 잠금

`도구(T)` 메뉴에는 창 외곽 배치 잠금과 별개로 다음 두 체크 명령이 있다. 둘 다 신규 설치 기본값은 해제이며, 해제 상태에서는 기존처럼 마지막 사용 상태가 계속 저장된다.

- `창 경로·위치 잠금(P)`: 체크하는 순간 현재 생성된 각 패널(최대 6개)의 활성 경로를 별도 스냅숏으로 저장한다. 다음 실행 때 저장된 각 패널 경로를 우선 복원한다. 실행 중에는 다른 폴더로 자유롭게 이동할 수 있지만 다음 재시작에서는 잠긴 경로가 다시 적용된다. 다시 누르면 체크가 해제되고 이후에는 마지막 사용 경로를 따르는 기본 동작으로 돌아간다.
- `창 분할·크기 잠금(S)`: 체크하는 순간 현재 행·열 개수, 가로/세로 분할 비율 및 분할 픽셀 크기를 별도 스냅숏으로 저장한다. 다음 실행 때 해당 분할을 우선 복원한다. 체크 해제 후에는 사용자가 조정한 마지막 분할 상태가 다시 정상 저장된다.

2×2 환경을 고정하려면 먼저 원하는 네 패널 경로와 분할선을 맞춘 뒤 두 명령에 각각 체크한다. 경로만 고정하거나 분할만 고정하는 것도 가능하다. 창 전체의 화면상 위치·외곽 크기는 기존 `창 위치·크기 잠금(L)`이 담당하므로 세 잠금은 독립적으로 조합할 수 있다.

관련 설정 키는 `fxfile-main.conf`의 `main.view.path_locked`, `main.view1.locked_path`~`main.view6.locked_path`, `main.view.split_locked`, `main.view.locked_row_count`, `main.view.locked_column_count`, `main.view.locked_ratio1`~`3`, `main.view.locked_size1`~`3`이다. 사용자가 체크할 때만 스냅숏이 갱신되며, 단순 종료는 잠긴 스냅숏을 덮어쓰지 않는다.

### 내장 간단 계산기

`도구(T)` 메뉴의 세 잠금 명령 바로 아래 `계산기(C)`를 누르거나, 메인 도구 모음의 돋보기 `검색` 바로 오른쪽에 있는 계산기 아이콘을 누르면 별도 프로그램 없이 경량 계산기가 열린다. `보기(V) > 도구 모음(T) > 사용자 지정(C)...`의 `현재 도구 모음 단추(T)` 목록에도 `검색` 다음 항목으로 `계산기`가 표시되며, 이 창에서 추가·제거·순서 변경이 가능하다.

새 설치 또는 초기 도구 모음에는 계산기 버튼이 기본 표시된다. 기존 세 배포본은 `fxfile-toolbar.dat`에도 `검색(ID 34014) -> 계산기(ID 34017)` 순서를 반영했으므로 즉시 같은 배치로 시작한다. 계산기 아이콘은 작은/큰 아이콘 및 활성/비활성 도구 모음 이미지에 모두 등록되어 있다.

계산식 입력란은 다음을 지원한다.

- 사칙연산: `+`, `-`, `*`, `/`
- 괄호와 우선순위: `(12.5 + 3) * 2 / 4`
- 소수와 단항 부호: `-2.5 + +4`
- Enter 또는 `계산(C)` 버튼으로 계산, `지우기(L)`로 초기화, Esc 또는 `닫기`로 종료

잘못된 식과 0 나눗셈은 결과란에 오류를 표시하고 FxFile은 계속 실행된다. 계산기는 설정 파일이나 외부 자산을 만들지 않는다.

### File integrity monitoring(FIM)

하나 이상의 일반 파일을 선택한 뒤 `파일(F) > CRC Chunsum 검사(V)...` 바로 아래의 `File integrity 모니터링(I)...`을 누른다. 폴더는 모니터링 대상에서 제외된다.

대화상자를 열 때 선택 파일마다 파일 크기, 최종 수정 시각과 SHA-256을 계산하고 그 값을 세션 기준선으로 삼는다. 목록에는 전체 경로, 바이트 크기, 수정 시각, SHA-256, 기준선 대비 무결성 상태, 첫 번째 선택 파일 대비 차이 판정이 표시된다.

- `지금 재검사`: 현재 파일을 다시 읽어 기준선과 비교한다. 크기 변경, 내용/해시 변경, 읽기 실패를 구분한다.
- `현재값을 기준선으로`: 현재 상태를 새 정상 기준선으로 승인한다.
- `3초마다 자동 재검사`: 체크한 동안 3초 간격으로 재검사한다. 기본값은 해제다.
- `보고서 복사`: 경로, SHA-256, 기준선 상태와 비교 판정을 탭 구분 텍스트로 클립보드에 복사한다.

첫 번째 선택 파일이 비교 기준이다. 다른 파일이 같은 바이트 크기와 같은 SHA-256이면 `완벽하게 동일`, 크기가 다르면 `크기 다름`, 크기는 같지만 SHA-256이 다르면 `내용/해시 다름`으로 표시한다. 이름·경로·NTFS 권한·소유자·대체 데이터 스트림 같은 메타데이터 동일성까지 증명하는 기능은 아니며, 선택한 일반 파일의 기본 데이터 스트림 내용 동일성을 검사한다.

기준선은 FIM 대화상자 세션 안에만 유지되고 별도 데이터베이스나 설정 파일을 만들지 않는다. 따라서 장기 감시가 필요하면 대화상자를 계속 열어 자동 재검사를 사용하거나, 보고서를 외부 기록으로 보관한 뒤 다음 세션에서 비교해야 한다. 대용량·다수 파일은 SHA-256 전체 읽기 동안 UI가 잠시 바쁠 수 있으므로 자동 재검사는 필요한 경우에만 켠다. FIM은 대상 파일을 읽기만 하며 수정·삭제하지 않는다.

변경 및 실제 종료·재실행 시험 결과는 `CHANGELOG_HISTORY-1차.md`의 Task 043, Task 045, Task 046, Task 047, Task 048을 참조한다.

## 적응형 고성능 파일 복사·이동 엔진

### 엔진 선택 원칙

일반 탐색창의 복사·이동·삭제는 `FileOpThread`의 백그라운드 작업에서 실행한다. 현재 구현은 모든 작업을 한 엔진에 강제로 넣지 않고 사전 검사 결과에 따라 다음 세 경로 중 하나를 고른다.

1. **안전 조건을 만족하는 로컬 일반 파일·폴더**는 `CopyFile2` 기반 적응형 엔진을 사용한다.
2. **그 밖의 일반 셸 작업과 휴지통 삭제**는 Vista 이후 권장 API인 `IFileOperation`으로 자동 복귀한다.
3. **이름 충돌 자동 이름 변경과 구형 다중 목적지 매핑처럼 `IFileOperation`으로 원래 의미를 그대로 표현하지 못하는 경우에만** `SHFileOperation`을 최종 호환 경로로 사용한다.

Microsoft는 Vista부터 `SHFileOperation`의 대체 API로 `IFileOperation`을 권장한다. 다만 `IFileOperation`은 현대적인 UI·취소·셸 호환 API이지 자동 고속화 API는 아니다. 이 PC의 1,500개 소파일 시험에서는 단순 교체가 빨라지지 않았으므로 **고속 경로**는 `CopyFile2`, **최신 호환·휴지통 경로**는 `IFileOperation`, **의미 보존이 필요한 최종 호환 경로**만 `SHFileOperation`으로 역할을 분리했다.

공식 기술 자료:

- SHFileOperation 대체 권고: <https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shfileoperationw>
- IFileOperation: <https://learn.microsoft.com/windows/win32/api/shobjidl_core/nn-shobjidl_core-ifileoperation>
- CopyFile2 확장 매개변수·취소·플래그: <https://learn.microsoft.com/windows/win32/api/winbase/ns-winbase-copyfile2_extended_parameters>
- Windows ODX 자동 오프로딩: <https://learn.microsoft.com/windows-hardware/drivers/ifs/offloaded-data-transfers>
- Robocopy `/MT`, `/J`: <https://learn.microsoft.com/windows-server/administration/windows-commands/robocopy>

### 작업량별 제한 병렬화

병렬 수를 무제한으로 늘리지 않는다. 특히 설치 컴퓨터의 D:는 기계식 HDD이므로 과도한 병렬 읽기/쓰기는 헤드 탐색, 백신 검사, 캐시 경쟁을 증가시킬 수 있다.

- 원본과 목적지 볼륨의 `StorageDeviceSeekPenaltyProperty`를 조회한다.
- HDD(탐색 페널티 있음) 또는 알 수 없는 장치가 한쪽이라도 포함되면 최대 2개 작업자로 제한한다.
- 양쪽 모두 SSD(탐색 페널티 없음)이고 파일 32개 이상, 최대 파일 64 MiB 이하, 평균 8 MiB 이하일 때만 최대 4개 작업자를 허용한다.
- 파일 8개 이상, 평균 16 MiB 이하인 일반 작업은 최대 2개 작업자다.
- 그 밖의 대용량·소수 파일: 1개 작업자
- 실제 작업자 수는 논리 CPU 수보다 많아지지 않는다.

대용량에 `COPY_FILE_NO_BUFFERING` 또는 Robocopy `/J`를 무조건 적용하지 않는다. 이 컴퓨터의 동일 D: 시험에서는 무버퍼 방식이 더 느렸고, buffered 측정은 Windows 쓰기 캐시 영향을 받았다. 따라서 제품 기본값은 Windows의 buffered `CopyFile2`이며, 저장장치·전송 방향을 별도로 입증하지 않은 상태에서 `/J`를 고정 기본값으로 만들지 않는다.

### 안정성 보장과 자동 복귀

- 고속 경로는 전체 원본 트리를 먼저 열거하고 대상 충돌 여부를 확인한 뒤 시작한다.
- 각 파일은 `COPY_FILE_FAIL_IF_EXISTS`로 복사해 사전 검사 뒤 다른 프로그램이 만든 파일도 덮어쓰지 않는다.
- 빈 폴더를 포함한 폴더 구조와 폴더 생성/접근/수정 시각 및 일반 속성을 복원한다.
- Windows 진행 창에서 진행률·현재 파일을 표시하고 `취소`를 받을 수 있다.
- 취소·오류·작업 중 원본 변경이 감지되면 원본은 삭제하지 않고, 이번 고속 작업이 새로 만든 대상 파일과 디렉터리를 역순 롤백한다. 고속 경로는 시작 전 대상 부재를 보장하므로 기존 사용자 파일을 롤백 대상으로 오인하지 않는다.
- 다른 볼륨 이동은 모든 복사가 성공한 뒤 원본 크기·마지막 수정시각, 대상 크기·마지막 수정시각, 디렉터리 신규 항목 유무를 다시 검사한다. 그 뒤에만 원본을 삭제한다. 삭제가 실패하면 완성된 대상은 보존해 데이터 손실보다 중복을 선택한다.
- 같은 볼륨 이동은 원래부터 파일 데이터 복사가 아닌 파일 시스템 메타데이터 이동이므로 Windows Shell 경로를 유지한다.
- `FILE_ATTRIBUTE_RECALL_ON_OPEN`/`RECALL_ON_DATA_ACCESS` 같은 클라우드 자리표시자와 재분석·오프라인 항목은 고속 경로에서 제외한다. 고속 경로의 사전 조건이 맞지 않으면 먼저 최신 `IFileOperation`, 꼭 필요한 의미 호환만 구형 셸 경로가 처리한다.
- 작업 스레드는 Shell/진행 COM 객체를 사용하기 전에 STA로 초기화한다.
- 실패·취소 작업은 붙여넣기 선택 통지와 사용자 실행 취소 기록을 성공 작업으로 잘못 등록하지 않는다.

### 성능·무결성 회귀시험

2026-08-12의 같은 D: 표본은 1,500개 × 16 KiB, 합계 24,576,000바이트다.

| 경로 | 관측 시간 |
|---|---:|
| 기존 `SHFileOperation` | 112.660초 |
| `CopyFile2` 직렬 탐색 | 12.319초 |
| 원시 `CopyFile2` 4개 제한 병렬 | 5.123초 |
| Robocopy `/MT:8` | 6.321초 |
| 최종 제품 적응형 x64 | 2.926초 |
| 최종 제품 적응형 x86 | 4.263초 |

캐시·백신·백그라운드 부하가 달라질 수 있으므로 38배를 고정 성능 보증으로 해석하지 않는다. 합격의 핵심은 동일 표본에서 기존 경로보다 명확히 빠르고, 다음 무결성 조건을 모두 만족하는 것이다.

- 1,500개 상대 경로·길이·마지막 수정시각·SHA-256 차이 0건
- 중첩/빈 폴더의 파일·디렉터리 시간·속성 차이 0건
- 기존 대상 충돌 시 고속 경로 거부, 기존 파일 변경 0건
- 40개 다중 선택 복사 결과 해시 차이 0건
- 실제 취소 후 원본 1,500개 보존, 이번 실행이 만든 대상 세대 롤백, 기존 대상 변경 0건
- C:→D: 101개 다른 볼륨 이동 후 대상 해시 차이 0건, 빈 폴더 보존, 성공 확인 뒤 원본 제거
- x64/x32 Release 빌드 및 양 아키텍처 격리 smoke 통과

시험 소스는 `tools\file_copy_engine_probe.cpp`, `tools\adaptive_file_operation_test.cpp`다. 운영 배포에는 시험 EXE/OBJ를 포함하지 않는다. 실제 수치와 증거 경로는 `CHANGELOG_HISTORY-1차.md` Task 051을 참조한다.

## 최신 셸 fallback·삭제 최적화·작업 잠금 (Task 052)

### 삭제 정책

- 일반 `Delete`는 복구 가능성을 우선하여 `IFileOperation`과 `FOFX_RECYCLEONDELETE`로 휴지통에 보낸다. 적응형 직접 삭제는 `FOF_ALLOWUNDO`가 설정된 요청을 의도적으로 거부한다.
- `Shift+Delete` 영구 삭제만 사전 검증을 모두 통과한 로컬 일반 항목에 한해 적응형 직접 삭제를 사용한다. 루트, Windows 디렉터리, Windows Resource Protection(WRP), UNC/네트워크, 재분석·클라우드·희소·암호화·오프라인·읽기 전용·시스템 항목 및 삭제 권한이 없는 대상은 직접 삭제하지 않는다.
- 파일은 저장장치 종류와 작업량에 맞춘 제한 병렬로 삭제하고, 디렉터리는 자식부터 부모 순서로 제거한다. 폴더 자체를 포함한 중첩 트리도 같은 정책을 사용한다.
- 취소·오류가 발생하면 이미 삭제된 수와 남은 수를 별도로 계산하여 표시한다. 영구 삭제는 본질적으로 롤백할 수 없으므로 일부 성공을 전체 성공으로 보고하지 않으며, 성공 통지와 실행 취소 기록도 남기지 않는다.
- 보호 대상은 최신 셸 경로에서도 차단한다. “속도를 위해 휴지통 복원과 Windows 보호 경계를 희생하지 않는다”가 고정 원칙이다.

### FxFile 작업 잠금과 Windows 권한

`편집(E) > 파일·폴더 잠금 관리(L)...`에서 선택 항목 또는 현재 폴더를 관리한다.

- **FxFile 잠금/해제**는 활성 설정 폴더의 `fxfile-operation-locks.conf`에 원자 저장된다. 잠긴 파일·폴더 및 하위 항목에 대한 복사 목적지 덮어쓰기, 이동, 삭제, 이름 변경, 다중 이름 변경, 파일 스크랩 작업을 FxFile 내부에서 차단한다.
- 잠금이 하나라도 있으면 외부 복사/이동 옵션이 켜져 있어도 해당 붙여넣기는 통합 내부 엔진으로 처리하여 잠금 검사를 우회하지 못하게 한다.
- **읽기 전용/쓰기 가능**은 일반 파일 속성만 변경한다. 폴더의 Read-only 비트는 Windows에서 보호 잠금 의미가 아니므로 폴더 보안 기능으로 오인하지 않는다.
- **Windows 보안**은 `SHObjectProperties`로 운영체제의 보안/ACL UI를 연다. 관리자 승인·자격 증명은 Windows가 직접 처리하며 FxFile은 Windows 암호를 입력받거나 저장하지 않고 소유권을 몰래 변경하지 않는다.
- Restart Manager는 `RmGetList`를 사용해 해당 경로를 사용 중인 프로세스를 **진단 표시만** 한다. 다른 프로세스의 핸들을 강제로 닫거나 `RmShutdown`으로 종료하지 않는다.
- Windows/WRP 보호 파일은 `SfcIsFileProtected`와 경로 경계로 판정하여 속성·삭제·소유권 자동 변경 대상에서 제외한다.

### 회귀시험과 배포 게이트

- x64/x86 각각 140개 복사 후 상대 경로·SHA-256 차이 0건.
- x64/x86 각각 중첩 100개 영구 삭제: x64 5.509초, x86 7.782초, 원본 트리 완전 제거.
- 일반 삭제: 적응형 직접 경로 거부, `IFileOperation` 휴지통 성공, 원본 경로 제거 및 복구 가능 경로 사용.
- `notepad.exe` 보호 표본: 직접 영구 삭제 거부, 최신 셸도 접근 거부, 전후 SHA-256 불변.
- 작업 중 원본 파일을 추가한 2 GiB 폴더 복사: 오류 1006, 원본과 추가 파일 보존, 이번 실행이 만든 대상 트리 완전 롤백.
- 최신 셸 x64 읽기 전용 파일 복사 속성·해시 보존, x86 이동 원본 제거/대상 생성 확인.
- 실제 최종 설치본 GUI에서 `편집 > 파일·폴더 잠금 관리`가 활성화되고 대화상자가 현재 경로, FxFile 잠금 상태, Restart Manager 진단과 WRP 경고를 표시하는 것을 확인했다. 시험 중 잠금/ACL/보안 변경 버튼은 누르지 않았다.

빌드·배포 도구는 중간 산출물 루트 EXE와 `Release` EXE가 다르거나 소스 `Korean.xml`과 아키텍처 산출물의 언어 파일이 다르면 배포를 중단한다. 이는 최신 컴파일 결과 대신 오래된 실행 파일·메뉴 문자열을 배포했던 재발 가능성을 차단한다. 최종 성공 manifest와 해시는 `CHANGELOG_HISTORY-1차.md` Task 052를 정본으로 한다.

### 잠금 관리 팝업의 호버 도움말

`편집(E) > 파일·폴더 잠금 관리(L)...`를 연 뒤 버튼이나 목록 위에 마우스 포인터를 약 0.35초 동안 두면 해당 기능의 목적, 적용 범위, 실행 방법, 효과와 주의사항이 툴팁으로 표시된다. 호버만으로는 파일·속성·권한을 변경하지 않는다. 설명은 최대 30초 동안 유지되며 긴 문장은 여러 줄로 표시된다.

| 호버 위치 | 핵심 설명과 실제 적용 범위 |
|---|---|
| 선택한 파일·폴더 목록 | 팝업이 관리할 대상과 상태를 보여 준다. `Windows 보안`은 선택한 한 항목에만, FxFile 잠금/해제와 읽기 전용/쓰기 가능은 목록 전체에 적용된다. |
| 새로 고침 | 파일 존재 여부, 파일/폴더 종류, FxFile 잠금, 읽기 전용, WRP 보호 여부와 선택 항목의 사용 프로세스 진단을 다시 읽는다. 변경 작업은 하지 않는다. |
| FxFile 잠금 | 목록 전체를 `fxfile-operation-locks.conf`에 기록하여 FxFile 안의 복사 덮어쓰기·이동·삭제·이름 변경을 막는다. 폴더 잠금은 하위 항목까지 포함하지만 Windows ACL이나 다른 프로그램까지 잠그지는 않는다. |
| FxFile 잠금 해제 | 목록 전체의 FxFile 내부 작업 차단만 해제한다. 읽기 전용, Windows ACL, 다른 프로그램의 열린 핸들은 그대로다. |
| 읽기 전용 | 목록의 일반 파일에 Read-only 속성을 켠다. 폴더와 Windows 보호 항목은 건너뛰며, 삭제 방지용 보안 잠금은 아니다. |
| 쓰기 가능 | 목록의 일반 파일에서 Read-only 속성만 해제한다. ACL·소유권·사용 중인 프로세스 잠금은 변경하지 않는다. |
| Windows 보안 | 목록에서 선택한 한 항목의 Windows 보안/ACL 속성 창을 연다. 관리자 승인과 암호 입력은 Windows가 직접 처리하며 FxFile은 자격 증명을 받거나 저장하지 않는다. |
| 잠금 사용 프로그램·서비스 목록 | 선택한 경로를 사용하는 프로세스를 Restart Manager로 읽어 보여 주는 진단 전용 영역이다. 프로세스 종료나 핸들 강제 폐쇄는 수행하지 않는다. |
| 닫기 | 팝업만 닫는다. 이미 적용한 FxFile 잠금, 파일 속성 또는 Windows 보안 변경은 되돌리지 않는다. |

안전한 기본 사용 순서는 `대상 확인 → 새로 고침 → 상태 확인 → 필요한 기능 한 번 실행 → 다시 새로 고침으로 결과 확인 → 닫기`다. Windows 시스템 파일 또는 권한 문제는 `Windows 보안`에서 사용자가 직접 판단하며, FxFile은 보호 파일 소유권 탈취나 타 프로세스 강제 종료를 자동화하지 않는다.

공식 기술 자료:

- `IFileOperation`: <https://learn.microsoft.com/windows/win32/api/shobjidl_core/nn-shobjidl_core-ifileoperation>
- `SHFileOperation` 대체 권고: <https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shfileoperationw>
- 저장장치 seek penalty 조회: <https://learn.microsoft.com/windows/win32/api/winioctl/ne-winioctl-storage_property_id>
- Restart Manager 함수: <https://learn.microsoft.com/windows/win32/rstmgr/functions>
- Windows Resource Protection: <https://learn.microsoft.com/windows/win32/wfp/about-windows-file-protection>
- 보안 속성 UI `SHObjectProperties`: <https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shobjectproperties>

### 운영상 속도 저하 요인과 권장 절차

- 같은 볼륨 이동은 대체로 즉시 끝나며, 다른 볼륨 이동은 물리적으로 복사 후 삭제하므로 느린 쪽 디스크 속도가 상한이다.
- 수천 개 소파일은 데이터 용량보다 파일 생성·메타데이터·백신 검사 횟수가 병목이다.
- 이 컴퓨터에서는 ALYac과 AhnLab V3 실시간 서비스가 함께 실행되고, TeraBox가 실제 빌드 PCH 파일을 잠근 사례가 확인됐다. 대규모 전송·빌드 전에는 사용자가 진행 중 동기화를 확인한 뒤 TeraBox 같은 클라우드 동기화를 일시 정지할 수 있다.
- 보안 프로그램을 자동 중단하거나 실시간 보호를 끄지 않는다. 예외 설정이 꼭 필요하면 신뢰 가능한 전용 빌드/출력 폴더만 해당 보안 제품의 공식 절차로 제한하고, 일반 다운로드·문서 폴더를 광범위하게 제외하지 않는다.
- 독립 대량 배치 작업에는 Windows 기본 포함 Robocopy도 유효하다. 예: `robocopy 원본 대상 /E /COPY:DAT /DCOPY:DAT /R:2 /W:2 /MT:4`. 실제 이동은 복사 검증 뒤 원본 삭제가 포함되므로 먼저 `/L` 또는 별도 복사로 대상·제외 규칙을 확인한다.
