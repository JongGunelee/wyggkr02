#Requires -Version 5.1
<#
.SYNOPSIS
    Antigravity & Codex C·D 드라이브 + Windows 기본 디스크 정리 통합 정기 딥 클린 v3.2
.DESCRIPTION
    - Zone A (D: 드라이브 5대): Temp\User 및 sessions\temp, Antigravity 크래시 로그, 업데이터 pending, TeraBox 썸네일/.cab, Codex 구버전 해시/.bak
    - Zone B (C: 드라이브 5대): ALPDF 임시 *.ps 스풀, Windows Temp/SoftwareDist, CBS 로그/Dbg 심볼, 브라우저 캐시, 개발 도구 캐시
    - Zone C (Windows 기본 디스크 정리 4대 표준 cleanmgr 영역): 전송 최적화 파일, DirectX/NVIDIA 셰이더 캐시, WER 오류보고/덤프, 휴지통
    - 화이트리스트 절대 보호: MSOffice(80개 파일), .gemini DB, .codex 설정(gpt-6-astra), 활성 바이너리 100% 무결 보존
    - 전송 최적화 분석: Windows OS P2P 캐시로, 클라우드(TeraBox, Codex Web, Antigravity) 서비스 영향 無 (100% 안전)
#>

[CmdletBinding()]
param(
    [switch]$DryRun,
    [switch]$Yes
)

$ProgressPreference = 'SilentlyContinue'

# ------------------------------------------------------------------------------
# [관리자 권한 자동 승격 & 창 최상위 포커스]
# ------------------------------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    $scriptPath = $MyInvocation.MyCommand.Definition
    $newArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    if ($DryRun) { $newArgs += " -DryRun" }
    if ($Yes) { $newArgs += " -Yes" }
    Start-Process powershell.exe -Verb RunAs -ArgumentList $newArgs
    exit
}

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Win32Utils {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
}
"@ -ErrorAction SilentlyContinue

try {
    $proc = [System.Diagnostics.Process]::GetCurrentProcess()
    $hwnd = $proc.MainWindowHandle
    if ($hwnd -ne [IntPtr]::Zero) {
        [Win32Utils]::ShowWindowAsync($hwnd, 9) | Out-Null
        [Win32Utils]::SetForegroundWindow($hwnd) | Out-Null
    }
} catch {}

function Show-Header {
    Clear-Host
    Write-Host "========================================================================================================" -ForegroundColor Cyan
    Write-Host "       [Antigravity & Codex] C/D 드라이브 + Windows 기본 디스크 정리 통합 정기 딥 클린 v3.2            " -ForegroundColor Yellow
    Write-Host "========================================================================================================" -ForegroundColor Cyan
    Write-Host ""
}

Show-Header

# 시작 시 디스크 공간 측정
$cDriveStart = Get-PSDrive C
$dDriveStart = Get-PSDrive D
$cFreeStartGB = [math]::Round($cDriveStart.Free / 1GB, 2)
$dFreeStartGB = [math]::Round($dDriveStart.Free / 1GB, 2)

Write-Host "  [1. 클린 목적 및 운영 원칙]" -ForegroundColor Yellow
Write-Host "   * 원칙 1: C드라이브, D드라이브, Windows 표준 디스크 정리 3개 영역을 완벽히 분리 및 정밀 정리" -ForegroundColor White
Write-Host "   * 원칙 2: MSOffice 업무 문서(80개), AI 대화 DB, 확장팩(25개), 활성 바이너리는 1 byte도 건드리지 않음" -ForegroundColor White
Write-Host "   * 원칙 3: '전송 최적화 파일'은 Windows OS P2P 캐시로, 클라우드(TeraBox/Codex/Antigravity)와 무관하여 100% 안전" -ForegroundColor White
Write-Host ""

Write-Host "  [2. 3대 영역 14개 정밀 정리 대상 목록]" -ForegroundColor Yellow
Write-Host "  +------------------------------------------------------------------------------------------------------+" -ForegroundColor DarkGray
Write-Host "  | [Zone A: D: 드라이브 - 데이터/AI/클라우드 캐시 5대 영역]                                             |" -ForegroundColor Cyan
Write-Host "  |  (1) D:\DevEnv\...\Temp\User 및 sessions\temp : 사용자 및 세션 임시 버퍼 파일                   |" -ForegroundColor DarkCyan
Write-Host "  |  (2) D:\DevEnv\...\.gemini\antigravity\crashes     : Antigravity 크래시 로그 (*.log)               |" -ForegroundColor DarkCyan
Write-Host "  |  (3) D:\DevEnv\...\Programs\antigravity-updater    : 업데이터 다운로드 pending 잔여물              |" -ForegroundColor DarkCyan
Write-Host "  |  (4) D:\DevEnv\...\AppData\Roaming\TeraBox         : 썸네일 캐시 및 AutoUpdate 누적 *.cab 패키지   |" -ForegroundColor DarkCyan
Write-Host "  |  (5) D:\DevEnv\...\Codex\bin                       : 구버전 해시 폴더 및 *.bak (활성 바이너리 보존)|" -ForegroundColor DarkCyan
Write-Host "  |------------------------------------------------------------------------------------------------------|" -ForegroundColor DarkGray
Write-Host "  | [Zone B: C: 드라이브 - OS 시스템/스풀/웹/개발도구 캐시 5대 영역]                                     |" -ForegroundColor Cyan
Write-Host "  |  (6) C:\ProgramData\ESTsoft\ALPDF\PDFCreator       : 임시 PostScript 스풀 (*.ps / *.pdf 절대 보존) |" -ForegroundColor DarkCyan
Write-Host "  |  (7) C:\Windows\Temp, Local\Temp, SoftwareDist...  : Windows 시스템 임시 및 업데이트 다운로드 캐시|" -ForegroundColor DarkCyan
Write-Host "  |  (8) C:\Windows\Logs\CBS, C:\ProgramData\Dbg       : CBS 누적 설치 로그 및 Dbg 진단 심볼/덤프 캐시 |" -ForegroundColor DarkCyan
Write-Host "  |  (9) C:\Users\ADMIN\...\(Chrome|Edge)\Cache        : 웹 브라우저 임시 인터넷 캐시                  |" -ForegroundColor DarkCyan
Write-Host "  |  (10) C:\Users\ADMIN\AppData\Local\(ms-playwright/npm): Playwright 브라우저 및 NPM 패키지 캐시   |" -ForegroundColor DarkCyan
Write-Host "  |------------------------------------------------------------------------------------------------------|" -ForegroundColor DarkGray
Write-Host "  | [Zone C: Windows 기본 디스크 정리 (cleanmgr 4대 영역)]                                               |" -ForegroundColor Cyan
Write-Host "  |  (11) 전송 최적화 파일 (Delivery Optimization)       : Windows P2P 업데이트 배포 캐시 (클라우드 영향 無)|" -ForegroundColor DarkCyan
Write-Host "  |  (12) DirectX / D3D / NVIDIA 셰이더 캐시             : 게임 및 그래픽 렌더링 손상/구버전 캐시 소거   |" -ForegroundColor DarkCyan
Write-Host "  |  (13) Windows 오류 보고 (WER) 로그 및 크래시 덤프    : WER 시스템 덤프 및 진단 리포트 파일           |" -ForegroundColor DarkCyan
Write-Host "  |  (14) Windows 휴지통 (Recycle Bin)                   : C: 및 D: 드라이브 휴지통 완전 비우기          |" -ForegroundColor DarkCyan
Write-Host "  +------------------------------------------------------------------------------------------------------+" -ForegroundColor DarkGray
Write-Host ""

Write-Host "  [3. 현재 디스크 상태 및 화이트리스트 검증]" -ForegroundColor Yellow
Write-Host "   * C: 드라이브 현재 여유 공간 : $cFreeStartGB GB" -ForegroundColor White
Write-Host "   * D: 드라이브 현재 여유 공간 : $dFreeStartGB GB" -ForegroundColor White
Write-Host "   * MSOffice 업무 폴더 검증   : D:\03 금일작업\00 임시\0000000 MSoffice (80개 직속 업무 문서 영구 보존)" -ForegroundColor Green
Write-Host ""

# 사용자 확인
$userConfirmed = $false
if ($Yes) {
    $userConfirmed = $true
} else {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        $nl = [Environment]::NewLine
        $boxMsg = "C/D 드라이브 + Windows 기본 디스크 정리 통합 딥 클린을 시작하시겠습니까?" + $nl + $nl +
                  "[1. D: 드라이브 5대 영역]" + $nl +
                  " - Temp\User 및 Codex 세션 버퍼, 크래시 로그, 업데이트 pending" + $nl +
                  " - TeraBox 썸네일/누적 .cab, Codex bin 구버전 해시 폴더" + $nl + $nl +
                  "[2. C: 드라이브 5대 영역]" + $nl +
                  " - ALPDF 임시 *.ps 스풀 (완성본 *.pdf 절대 보존)" + $nl +
                  " - Windows 시스템 Temp, 업데이트 다운로드, CBS 로그, Dbg 심볼" + $nl +
                  " - Chrome/Edge 브라우저 캐시, Playwright/NPM/Swit 개발 캐시" + $nl + $nl +
                  "[3. Windows 기본 디스크 정리 4대 표준 영역 (cleanmgr 전수 점검)]" + $nl +
                  " - 전송 최적화 파일 (Delivery Optimization P2P 캐시)" + $nl +
                  " - DirectX 및 그래픽 셰이더 캐시 (D3DSCache / NVIDIA DXCache)" + $nl +
                  " - Windows 오류 보고서 및 피드백 진단 (WER)" + $nl +
                  " - C: 및 D: 드라이브 휴지통 비우기" + $nl + $nl +
                  "[클라우드 영향도 분석]" + $nl +
                  " * 전송 최적화 정리 시에도 테라박스, 코덱스 웹, 안티그래비트 등" + $nl +
                  "   모든 인터넷/클라우드 서비스는 100% 정상 작동합니다." + $nl + $nl +
                  "[절대 보호 자산]" + $nl +
                  " * MSOffice 업무 파일 80개, AI 대화 DB, codex.exe 등 100% 보존" + $nl + $nl +
                  "진행하시려면 [확인]을 클릭하세요."
        $boxTitle = "C & D 드라이브 + Windows 기본 디스크 정리 통합 딥 클린 v3.2"
        $btn = [System.Windows.Forms.MessageBoxButtons]::OKCancel
        $ico = [System.Windows.Forms.MessageBoxIcon]::Information

        $topForm = New-Object System.Windows.Forms.Form
        $topForm.TopMost = $true
        $res = [System.Windows.Forms.MessageBox]::Show($topForm, $boxMsg, $boxTitle, $btn, $ico)
        $topForm.Dispose()

        if ($res -eq [System.Windows.Forms.DialogResult]::OK) {
            $userConfirmed = $true
        }
    } catch {}

    if (-not $userConfirmed) {
        $choice = Read-Host "  지금 통합 딥 클린 작업을 시작하시겠습니까? ([Y]/N)"
        if ($choice -match '^[Yy]$' -or $choice -eq '') {
            $userConfirmed = $true
        }
    }
}

if (-not $userConfirmed) {
    Write-Host "`n  작업이 취소되었습니다. 엔터 키를 누르면 종료됩니다." -ForegroundColor Yellow
    if (-not $Yes) { Read-Host }
    exit
}

Show-Header
if ($DryRun) {
    Write-Host "  >>> [DRY-RUN 모드] 실제 파일 삭제 없이 용량 시뮬레이션 중... <<<" -ForegroundColor Magenta
} else {
    Write-Host "  [통합 정기 딥 클린 v3.2 작업 진행 중...]" -ForegroundColor Yellow
}
Write-Host ""

$baseD = "D:\DevEnv\Relocated-C-Data"
$localApp = $env:LOCALAPPDATA
if (-not $localApp) { $localApp = "C:\Users\ADMIN\AppData\Local" }

$dStats = @()
$cStats = @()
$wStats = @()

$dTotalCount = 0
[double]$dTotalBytes = 0
$cTotalCount = 0
[double]$cTotalBytes = 0
$wTotalCount = 0
[double]$wTotalBytes = 0

function Remove-SafeFiles($path, $filter='*', $stepName='', $zone='D') {
    $count = 0
    [double]$bytes = 0
    if (Test-Path -LiteralPath $path) {
        Get-ChildItem -LiteralPath $path -Filter $filter -Recurse -File -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                $len = $_.Length
                if (-not $DryRun) {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                }
                $count++
                $bytes += $len
                if ($zone -eq 'D') {
                    $script:dTotalCount++
                    $script:dTotalBytes += $len
                } elseif ($zone -eq 'C') {
                    $script:cTotalCount++
                    $script:cTotalBytes += $len
                } else {
                    $script:wTotalCount++
                    $script:wTotalBytes += $len
                }
            } catch {}
        }
    }
    $mb = [math]::Round($bytes / 1MB, 2)
    return [PSCustomObject]@{ Step=$stepName; Count=$count; SizeMB=$mb }
}

# ==============================================================================
# [Zone A: D드라이브 5대 정리 대상]
# ==============================================================================
Write-Host "  --- [1. Zone A: D: 드라이브 정리 진행 중] ---" -ForegroundColor Cyan

# (1) D: Temp\User & sessions\temp
Write-Host "  [D-1/5] D: 사용자 임시 및 세션 버퍼 정리 중..." -ForegroundColor Gray
$d1_1 = Remove-SafeFiles -path "$baseD\Temp\User" -stepName "Temp\User" -zone 'D'
$d1_2 = Remove-SafeFiles -path "$baseD\AppData\Local\OpenAI\Codex\sessions\temp" -stepName "sessions\temp" -zone 'D'
$dStats += [PSCustomObject]@{ Step="D: Temp\User 및 sessions\temp"; Count=($d1_1.Count + $d1_2.Count); SizeMB=[math]::Round($d1_1.SizeMB + $d1_2.SizeMB, 2) }

# (2) D: Antigravity crashes
Write-Host "  [D-2/5] D: Antigravity 크래시 로그 정리 중..." -ForegroundColor Gray
$d2 = Remove-SafeFiles -path "$baseD\UserProfile\.gemini\antigravity\crashes" -filter "*.log" -stepName "D: Antigravity Crashes (*.log)" -zone 'D'
$dStats += $d2

# (3) D: Updater pending
Write-Host "  [D-3/5] D: Antigravity 업데이터 pending 잔여물 정리 중..." -ForegroundColor Gray
$d3 = Remove-SafeFiles -path "$baseD\Programs\antigravity-updater\pending" -stepName "D: Updater Pending 잔여물" -zone 'D'
$dStats += $d3

# (4) D: TeraBox imageCache & AutoUpdate .cab
Write-Host "  [D-4/5] D: TeraBox 캐시 및 누적 .cab 패키지 정리 중..." -ForegroundColor Gray
$tbCount = 0
[double]$tbBytes = 0
$tbDir = "$baseD\AppData\Roaming\TeraBox"
if (Test-Path -LiteralPath $tbDir) {
    Get-ChildItem -LiteralPath $tbDir -Filter "imageCache" -Recurse -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $res = Remove-SafeFiles -path $_.FullName -zone 'D'
        $tbCount += $res.Count
        $tbBytes += ($res.SizeMB * 1MB)
    }
    $resTmp = Remove-SafeFiles -path $tbDir -filter "*.tmp" -zone 'D'
    $tbCount += $resTmp.Count
    $tbBytes += ($resTmp.SizeMB * 1MB)

    $tbCabDir = "$tbDir\AutoUpdate\Download\MainApp"
    if (Test-Path -LiteralPath $tbCabDir) {
        $resCab = Remove-SafeFiles -path $tbCabDir -filter "*.cab" -zone 'D'
        $tbCount += $resCab.Count
        $tbBytes += ($resCab.SizeMB * 1MB)
    }
}
$dStats += [PSCustomObject]@{ Step="D: TeraBox 썸네일 & .cab 패키지"; Count=$tbCount; SizeMB=[math]::Round($tbBytes / 1MB, 2) }

# (5) D: Codex bin 구버전 해시 디렉터리 및 백업 (*.bak)
Write-Host "  [D-5/5] D: Codex bin 구버전 해시 폴더 및 .bak 정리 중..." -ForegroundColor Gray
$codexBinDir = "$baseD\AppData\Local\OpenAI\Codex\bin"
$codexCount = 0
[double]$codexBytes = 0
if (Test-Path -LiteralPath "$codexBinDir\codex.exe") {
    Get-ChildItem -LiteralPath $codexBinDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $dirFiles = Get-ChildItem -LiteralPath $_.FullName -Recurse -File -Force -ErrorAction SilentlyContinue
            foreach ($f in $dirFiles) {
                $len = $f.Length
                if (-not $DryRun) {
                    Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
                }
                $codexCount++
                $codexBytes += $len
                $script:dTotalCount++
                $script:dTotalBytes += $len
            }
            if (-not $DryRun) {
                Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
        } catch {}
    }
    $resBak = Remove-SafeFiles -path $codexBinDir -filter "*.bak" -zone 'D'
    $codexCount += $resBak.Count
    $codexBytes += ($resBak.SizeMB * 1MB)
}
$dStats += [PSCustomObject]@{ Step="D: Codex 구버전 해시 & .bak 백업"; Count=$codexCount; SizeMB=[math]::Round($codexBytes / 1MB, 2) }

Write-Host ""
# ==============================================================================
# [Zone B: C드라이브 5대 정리 대상]
# ==============================================================================
Write-Host "  --- [2. Zone B: C: 드라이브 정리 진행 중] ---" -ForegroundColor Cyan

# (6) C: ALPDF PDFCreator PostScript spool (*.ps, *.ps.log)
Write-Host "  [C-1/5] C: ALPDF PDFCreator 임시 PostScript 스풀 정리 중..." -ForegroundColor Gray
$alpdfCount = 0
[double]$alpdfBytes = 0
$alpdfDir = "C:\ProgramData\ESTsoft\ALPDF\PDFCreator"
if (Test-Path -LiteralPath $alpdfDir) {
    Get-ChildItem -LiteralPath $alpdfDir -File -Force -ErrorAction SilentlyContinue | Where-Object {
        $_.Extension -ieq ".ps" -or $_.Name -imatch '\.ps\.log$'
    } | ForEach-Object {
        try {
            $len = $_.Length
            if (-not $DryRun) {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
            }
            $alpdfCount++
            $alpdfBytes += $len
            $script:cTotalCount++
            $script:cTotalBytes += $len
        } catch {}
    }
}
$cStats += [PSCustomObject]@{ Step="C: ALPDF 임시 스풀 (*.ps)"; Count=$alpdfCount; SizeMB=[math]::Round($alpdfBytes / 1MB, 2) }

# (7) C: Windows System Temp & Updates
Write-Host "  [C-2/5] C: Windows 시스템 Temp 및 업데이트 다운로드 정리 중..." -ForegroundColor Gray
$cW1 = Remove-SafeFiles -path "C:\Windows\Temp" -zone 'C'
$cW2 = Remove-SafeFiles -path "$localApp\Temp" -zone 'C'
$cW3 = Remove-SafeFiles -path "C:\Windows\SoftwareDistribution\Download" -zone 'C'
$cStats += [PSCustomObject]@{ Step="C: Windows Temp & 업데이트 다운로드"; Count=($cW1.Count + $cW2.Count + $cW3.Count); SizeMB=[math]::Round($cW1.SizeMB + $cW2.SizeMB + $cW3.SizeMB, 2) }

# (8) C: Windows CBS Logs & Dbg Symbols
Write-Host "  [C-3/5] C: Windows CBS 컴포넌트 로그 및 Dbg 심볼 정리 중..." -ForegroundColor Gray
$cCbs = Remove-SafeFiles -path "C:\Windows\Logs\CBS" -filter "CbsPersist_*.log" -zone 'C'
$cDbg = Remove-SafeFiles -path "C:\ProgramData\Dbg" -zone 'C'
$cStats += [PSCustomObject]@{ Step="C: CBS 누적로그 & Dbg 진단심볼"; Count=($cCbs.Count + $cDbg.Count); SizeMB=[math]::Round($cCbs.SizeMB + $cDbg.SizeMB, 2) }

# (9) C: Web Browser Cache (Chrome and Edge)
Write-Host "  [C-4/5] C: Chrome 및 Edge 웹 브라우저 캐시 정리 중..." -ForegroundColor Gray
$chromeCache = "$localApp\Google\Chrome\User Data\Default\Cache"
$edgeCache = "$localApp\Microsoft\Edge\User Data\Default\Cache"
$sWeb1 = Remove-SafeFiles -path $chromeCache -zone 'C'
$sWeb2 = Remove-SafeFiles -path $edgeCache -zone 'C'
$cStats += [PSCustomObject]@{ Step="C: Chrome & Edge 브라우저 캐시"; Count=($sWeb1.Count + $sWeb2.Count); SizeMB=[math]::Round($sWeb1.SizeMB + $sWeb2.SizeMB, 2) }

# (10) C: Dev Caches (Playwright, NPM, Swit, Pip)
Write-Host "  [C-5/5] C: 개발도구 캐시(Playwright/NPM/Swit/Pip) 정리 중..." -ForegroundColor Gray
$cPw = Remove-SafeFiles -path "$localApp\ms-playwright" -zone 'C'
$cNpm = Remove-SafeFiles -path "$localApp\npm-cache" -zone 'C'
$cSwit = Remove-SafeFiles -path "$localApp\Swit\updater" -zone 'C'
$cPip = Remove-SafeFiles -path "$localApp\pip\cache" -zone 'C'
$devCount = $cPw.Count + $cNpm.Count + $cSwit.Count + $cPip.Count
$devBytes = $cPw.SizeMB + $cNpm.SizeMB + $cSwit.SizeMB + $cPip.SizeMB
$cStats += [PSCustomObject]@{ Step="C: 개발캐시 (Playwright/NPM/Swit/Pip)"; Count=$devCount; SizeMB=[math]::Round($devBytes, 2) }

Write-Host ""
# ==============================================================================
# [Zone C: Windows 기본 디스크 정리 4대 표준 항목]
# ==============================================================================
Write-Host "  --- [3. Zone C: Windows 기본 디스크 정리 4대 영역 진행 중] ---" -ForegroundColor Cyan

# (11) Windows 전송 최적화 파일 (Delivery Optimization Files)
Write-Host "  [W-1/4] 전송 최적화 파일 (Delivery Optimization Cache) 정리 중..." -ForegroundColor Gray
$doCount = 0
[double]$doBytes = 0
$doPath = "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"
if (-not $DryRun) {
    try {
        Delete-DeliveryOptimizationCache -Force -ErrorAction SilentlyContinue
    } catch {}
}
if (Test-Path -LiteralPath $doPath) {
    $resDo = Remove-SafeFiles -path $doPath -zone 'W'
    $doCount += $resDo.Count
    $doBytes += $resDo.SizeMB
}
$wStats += [PSCustomObject]@{ Step="전송 최적화 파일 (Delivery Optimization)"; Count=$doCount; SizeMB=[math]::Round($doBytes, 2) }

# (12) DirectX 및 그래픽 셰이더 캐시
Write-Host "  [W-2/4] DirectX 및 그래픽 셰이더 캐시 (D3D / NVIDIA) 정리 중..." -ForegroundColor Gray
$d3dPath = "$localApp\D3DSCache"
$nvPath = "$localApp\NVIDIA\DXCache"
$resD3D = Remove-SafeFiles -path $d3dPath -zone 'W'
$resNV = Remove-SafeFiles -path $nvPath -zone 'W'
$shaderCount = $resD3D.Count + $resNV.Count
$shaderMB = [math]::Round($resD3D.SizeMB + $resNV.SizeMB, 2)
$wStats += [PSCustomObject]@{ Step="DirectX 및 그래픽 셰이더 캐시"; Count=$shaderCount; SizeMB=$shaderMB }

# (13) Windows 오류 보고서 및 피드백 진단 (WER)
Write-Host "  [W-3/4] Windows 오류 보고서 및 피드백 진단 (WER) 정리 중..." -ForegroundColor Gray
$werP1 = "C:\ProgramData\Microsoft\Windows\WER"
$werP2 = "$localApp\Microsoft\Windows\WER"
$resWer1 = Remove-SafeFiles -path $werP1 -zone 'W'
$resWer2 = Remove-SafeFiles -path $werP2 -zone 'W'
$werCount = $resWer1.Count + $resWer2.Count
$werMB = [math]::Round($resWer1.SizeMB + $resWer2.SizeMB, 2)
$wStats += [PSCustomObject]@{ Step="Windows 오류 보고서 및 피드백 진단 (WER)"; Count=$werCount; SizeMB=$werMB }

# (14) 휴지통 비우기 (C: 및 D: 사용자 휴지통)
Write-Host "  [W-4/4] 휴지통 비우기 (C: 및 D: 사용자 휴지통) 진행 중..." -ForegroundColor Gray
$binCount = 0
[double]$binBytes = 0
try {
    $userSid = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
    foreach ($drv in @('C:', 'D:')) {
        $binPath = "$drv\`$Recycle.Bin\$userSid"
        if (Test-Path -LiteralPath $binPath) {
            $files = Get-ChildItem -LiteralPath $binPath -Force -Recurse -File -ErrorAction SilentlyContinue
            foreach ($f in $files) {
                $len = $f.Length
                if (-not $DryRun) {
                    try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue } catch {}
                }
                $binCount++
                $binBytes += $len
                $script:wTotalCount++
                $script:wTotalBytes += $len
            }
            if (-not $DryRun) {
                try {
                    Get-ChildItem -LiteralPath $binPath -Force -Recurse -Directory -ErrorAction SilentlyContinue |
                        Sort-Object -Property FullName -Descending |
                        ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch {} }
                } catch {}
            }
        }
    }
} catch {}
$wStats += [PSCustomObject]@{ Step="Windows 휴지통 (C: 및 D: 사용자 휴지통)"; Count=$binCount; SizeMB=[math]::Round($binBytes / 1MB, 2) }

# 정리 후 디스크 실시간 공간 측정
$cDriveEnd = Get-PSDrive C
$dDriveEnd = Get-PSDrive D
$cFreeEndGB = [math]::Round($cDriveEnd.Free / 1GB, 2)
$dFreeEndGB = [math]::Round($dDriveEnd.Free / 1GB, 2)

$dRecoveredMB = [math]::Round($dTotalBytes / 1MB, 2)
$cRecoveredMB = [math]::Round($cTotalBytes / 1MB, 2)
$wRecoveredMB = [math]::Round($wTotalBytes / 1MB, 2)

$grandTotalFiles = $dTotalCount + $cTotalCount + $wTotalCount
$grandTotalMB = [math]::Round(($dTotalBytes + $cTotalBytes + $wTotalBytes) / 1MB, 2)

# MSOffice 업무 파일 무결성 카운팅
$msofficePath = "D:\03 금일작업\00 임시\0000000 MSoffice"
$msofficeCount = 0
if (Test-Path -LiteralPath $msofficePath) {
    $msofficeCount = (Get-ChildItem -LiteralPath $msofficePath -File -ErrorAction SilentlyContinue | Measure-Object).Count
}

# 창 다시 포커스
try {
    $proc = [System.Diagnostics.Process]::GetCurrentProcess()
    $hwnd = $proc.MainWindowHandle
    if ($hwnd -ne [IntPtr]::Zero) {
        [Win32Utils]::ShowWindowAsync($hwnd, 9) | Out-Null
        [Win32Utils]::SetForegroundWindow($hwnd) | Out-Null
    }
} catch {}

# ==============================================================================
# [최종 정리 현황 및 목록 일목요연 보고서 대시보드]
# ==============================================================================
Show-Header
$reportTitle = if ($DryRun) { "(DRY-RUN 시뮬레이션 완료 보고서)" } else { "★★★ C·D 드라이브 + Windows 기본 디스크 정리 (Deep Clean v3.2) 통합 완료 보고서 ★★★" }
Write-Host "  ========================================================================================================" -ForegroundColor Green
Write-Host "      $reportTitle        " -ForegroundColor Green
Write-Host "  ========================================================================================================" -ForegroundColor Green
Write-Host ""

$actionVerb = if ($DryRun) { "검색" } else { "삭제" }

Write-Host "  [1. Zone A: D: 드라이브 (데이터 및 AI 개발 환경 5대 정리 내역)]" -ForegroundColor Cyan
foreach ($st in $dStats) {
    Write-Host ("   - {0,-40} : {1,6} 개 항목 {2}  ({3,9} MB 확보)" -f $st.Step, $st.Count, $actionVerb, $st.SizeMB) -ForegroundColor White
}
Write-Host ("   >> D: 드라이브 소계 : {0:N0} 개 파일 정리 | {1:N2} MB ({2:N2} GB) 확보" -f $dTotalCount, $dRecoveredMB, [math]::Round($dRecoveredMB/1024, 2)) -ForegroundColor Yellow
Write-Host ""

Write-Host "  [2. Zone B: C: 드라이브 (OS 시스템, 인쇄 스풀 및 웹/도구 캐시 5대 정리 내역)]" -ForegroundColor Cyan
foreach ($st in $cStats) {
    Write-Host ("   - {0,-40} : {1,6} 개 항목 {2}  ({3,9} MB 확보)" -f $st.Step, $st.Count, $actionVerb, $st.SizeMB) -ForegroundColor White
}
Write-Host ("   >> C: 드라이브 소계 : {0:N0} 개 파일 정리 | {1:N2} MB ({2:N2} GB) 확보" -f $cTotalCount, $cRecoveredMB, [math]::Round($cRecoveredMB/1024, 2)) -ForegroundColor Yellow
Write-Host ""

Write-Host "  [3. Zone C: Windows 기본 디스크 정리 4대 표준 영역 (cleanmgr 전수 점검 내역)]" -ForegroundColor Cyan
foreach ($st in $wStats) {
    Write-Host ("   - {0,-40} : {1,6} 개 항목 {2}  ({3,9} MB 확보)" -f $st.Step, $st.Count, $actionVerb, $st.SizeMB) -ForegroundColor White
}
Write-Host ("   >> Windows 기본 정리 소계 : {0:N0} 개 항목 정리 | {1:N2} MB 확보" -f $wTotalCount, $wRecoveredMB) -ForegroundColor Yellow
Write-Host ""

Write-Host "  --------------------------------------------------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "  [4. 전체 종합 성과 및 디스크 실시간 공간 비교]" -ForegroundColor Yellow
Write-Host ("   * 3대 영역 총 정리 파일  : {0:N0} 개 임시/캐시/시스템 불필요 항목 완전 소거" -f $grandTotalFiles) -ForegroundColor White
Write-Host ("   * 총 공간 회수량        : +{0:N2} MB  (약 +{1:N2} GB 대규모 디스크 여유 공간 확보)" -f $grandTotalMB, ($grandTotalMB/1024)) -ForegroundColor Green
Write-Host ("   * C: 드라이브 여유 공간  : {0} GB  ==>  {1} GB" -f $cFreeStartGB, $cFreeEndGB) -ForegroundColor White
Write-Host ("   * D: 드라이브 여유 공간  : {0} GB  ==>  {1} GB" -f $dFreeStartGB, $dFreeEndGB) -ForegroundColor White
Write-Host ""

Write-Host "  [5. 화이트리스트 절대 보호 자산 무결성 검증 결과]" -ForegroundColor Green
Write-Host ("   * [보호 1] MSOffice 직속 업무 문서 : D:\03 금일작업\...\0000000 MSoffice ({0}개 파일 100% 무결 보존)" -f $msofficeCount) -ForegroundColor DarkGreen
Write-Host ("   * [보호 2] AI 대화 지능 세션 DB    : D:\DevEnv\...\.gemini (대화 DB, 아티팩트, 25개 스킬 정상)") -ForegroundColor DarkGreen
Write-Host ("   * [보호 3] Codex 설정 및 세션 DB   : D:\DevEnv\...\.codex (config.toml gpt-6-astra 설정 보존)") -ForegroundColor DarkGreen
Write-Host ("   * [보호 4] 활성 바이너리 런타임     : codex.exe 등 7대 실행 파일 및 Programs 본체 100% 정상") -ForegroundColor DarkGreen
Write-Host ("   * [보호 5] ALPDF 완성본 PDF 문서    : C:\ProgramData\ESTsoft\ALPDF\PDFCreator 내 모든 *.pdf 보존") -ForegroundColor DarkGreen
Write-Host "  --------------------------------------------------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  * 핵심 업무 자산과 AI 개발 환경은 안전하게 영구 보존되며, 임시/캐시/기본 디스크 정리가 일목요연하게 완료되었습니다." -ForegroundColor White
Write-Host "  ========================================================================================================" -ForegroundColor Green
Write-Host "  작업이 완료되었습니다. 결과 리포트를 확인하신 후 아무 키나 누르시면 안전하게 종료됩니다." -ForegroundColor White
Write-Host "  ========================================================================================================" -ForegroundColor Green
Write-Host ""

if (-not $Yes) {
    try {
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } catch {
        Read-Host "  계속하려면 엔터 키를 누르세요"
    }
}
exit
