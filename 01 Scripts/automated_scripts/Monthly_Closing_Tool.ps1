#Requires -Version 5.1
<#
.SYNOPSIS
    월마감 파일 패턴폴더 복사 및 파일명 치환 자동화 도구 (PowerShell 네이티브)
.DESCRIPTION
    - 대상 폴더 (A)의 2자리 숫자 접두사 파일을 타겟 상위 폴더 (B-1) 내 동일 접두사 서브폴더의 '01 자료'(B-2)로 복사
    - 파일명 앞 2자리를 '표준단가_'로 치환
    - 대상 폴더 및 타겟 폴더의 지속적인 변동성에 대응하여 네이티브 폴더 선택 다이얼로그 제공
    - 작업 결과 상세 검토 후 다이얼로그 박스에서 선택적 후속 조치:
        1) [원본 정리] 대상 폴더(A) 내 복사 성공 파일 삭제 (폴더 비어있을 시 폴더 삭제 옵션)
        2) [롤백] 방금 복사된 타겟 파일 전체 일괄 삭제 (원상 복구)
        3) [유지] 원본 및 복사본 모두 보존 후 완료
    - JSON 등 외부 임시 설정 파일 일체 생성 안 함
    - Windows 네이티브 UTF-8 / 한글 인코딩 완벽 지원
#>

[CmdletBinding()]
param()

$ProgressPreference = 'SilentlyContinue'

# Windows Forms 및 Drawing 어셈블리 로드
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# 콘솔 UTF-8 출력 설정
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Show-Header {
    Clear-Host
    Write-Host "========================================================================================================" -ForegroundColor Cyan
    Write-Host "             [월마감 자동화 도구] 파일 패턴폴더 매칭 복사 및 파일명 치환 v2.5                           " -ForegroundColor Yellow
    Write-Host "========================================================================================================" -ForegroundColor Cyan
    Write-Host ""
}

# 기본 경로 정의 (메모리 변수 - JSON 파일 생성 없음)
$defaultSource = "D:\03 금일작업\00 월마감\일성"
$defaultTarget = "D:\02 기숙사 및 사택\05 기숙사 및 사택 월마감\26년-사택 작업\09월\01 솔루션"
$defaultPrefix = "표준단가_"

Show-Header

# ------------------------------------------------------------------------------
# 1. 초기 폴더 경로 확인 및 선택 다이얼로그 (지속적 변동성 지원)
# ------------------------------------------------------------------------------
Write-Host "  [1. 작업 경로 설정 및 확인]" -ForegroundColor Yellow
Write-Host "   * 기본 대상 폴더 (A)       : $defaultSource" -ForegroundColor White
Write-Host "   * 기본 타겟 상위 폴더 (B-1): $defaultTarget" -ForegroundColor White
Write-Host "   * 적용 치환 접두사         : $defaultPrefix" -ForegroundColor White
Write-Host ""

$srcPath = $defaultSource
$tgtPath = $defaultTarget
$prefixText = $defaultPrefix

# 경로 선택 GUI 창
$configForm = New-Object System.Windows.Forms.Form
$configForm.Text = "월마감 파일 패턴폴더 매칭 & 복사 설정"
$configForm.Size = New-Object System.Drawing.Size(680, 370)
$configForm.StartPosition = "CenterScreen"
$configForm.TopMost = $true
$configForm.FormBorderStyle = "FixedDialog"
$configForm.MaximizeBox = $false

# 안내 레이블
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "작업할 폴더 경로를 확인하시고, 변경이 필요하면 [찾아보기...]를 클릭하세요."
$lblInfo.Font = New-Object System.Drawing.Font("맑은 고딕", 9, [System.Drawing.FontStyle]::Bold)
$lblInfo.Location = New-Object System.Drawing.Point(20, 15)
$lblInfo.Size = New-Object System.Drawing.Size(620, 25)
$configForm.Controls.Add($lblInfo)

# 1) 대상 폴더 (A)
$lblSrc = New-Object System.Windows.Forms.Label
$lblSrc.Text = "📂 대상 폴더 (A) [원본 파일 위치]:"
$lblSrc.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$lblSrc.Location = New-Object System.Drawing.Point(20, 50)
$lblSrc.Size = New-Object System.Drawing.Size(500, 20)
$configForm.Controls.Add($lblSrc)

$txtSrc = New-Object System.Windows.Forms.TextBox
$txtSrc.Text = $srcPath
$txtSrc.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$txtSrc.Location = New-Object System.Drawing.Point(20, 75)
$txtSrc.Size = New-Object System.Drawing.Size(520, 25)
$configForm.Controls.Add($txtSrc)

$btnBrowseSrc = New-Object System.Windows.Forms.Button
$btnBrowseSrc.Text = "찾아보기..."
$btnBrowseSrc.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$btnBrowseSrc.Location = New-Object System.Drawing.Point(550, 73)
$btnBrowseSrc.Size = New-Object System.Drawing.Size(95, 27)
$btnBrowseSrc.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "대상 폴더 (A)를 선택하세요"
    if (Test-Path $txtSrc.Text) { $dlg.SelectedPath = $txtSrc.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSrc.Text = $dlg.SelectedPath
    }
})
$configForm.Controls.Add($btnBrowseSrc)

# 2) 타겟 상위 폴더 (B-1)
$lblTgt = New-Object System.Windows.Forms.Label
$lblTgt.Text = "🎯 타겟 상위 폴더 (B-1) [01~99 서브폴더 위치]:"
$lblTgt.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$lblTgt.Location = New-Object System.Drawing.Point(20, 115)
$lblTgt.Size = New-Object System.Drawing.Size(500, 20)
$configForm.Controls.Add($lblTgt)

$txtTgt = New-Object System.Windows.Forms.TextBox
$txtTgt.Text = $tgtPath
$txtTgt.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$txtTgt.Location = New-Object System.Drawing.Point(20, 140)
$txtTgt.Size = New-Object System.Drawing.Size(520, 25)
$configForm.Controls.Add($txtTgt)

$btnBrowseTgt = New-Object System.Windows.Forms.Button
$btnBrowseTgt.Text = "찾아보기..."
$btnBrowseTgt.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$btnBrowseTgt.Location = New-Object System.Drawing.Point(550, 138)
$btnBrowseTgt.Size = New-Object System.Drawing.Size(95, 27)
$btnBrowseTgt.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "타겟 상위 폴더 (B-1)를 선택하세요"
    if (Test-Path $txtTgt.Text) { $dlg.SelectedPath = $txtTgt.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtTgt.Text = $dlg.SelectedPath
    }
})
$configForm.Controls.Add($btnBrowseTgt)

# 3) 접두사
$lblPrefix = New-Object System.Windows.Forms.Label
$lblPrefix.Text = "✏️ 치환할 접두사:"
$lblPrefix.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$lblPrefix.Location = New-Object System.Drawing.Point(20, 180)
$lblPrefix.Size = New-Object System.Drawing.Size(120, 20)
$configForm.Controls.Add($lblPrefix)

$txtPrefix = New-Object System.Windows.Forms.TextBox
$txtPrefix.Text = $prefixText
$txtPrefix.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$txtPrefix.Location = New-Object System.Drawing.Point(145, 177)
$txtPrefix.Size = New-Object System.Drawing.Size(150, 25)
$configForm.Controls.Add($txtPrefix)

# 하단 실행 버튼
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "🚀 분석 및 복사/치환 실행"
$btnStart.Font = New-Object System.Drawing.Font("맑은 고딕", 10, [System.Drawing.FontStyle]::Bold)
$btnStart.BackColor = [System.Drawing.Color]::FromArgb(41, 128, 185)
$btnStart.ForeColor = [System.Drawing.Color]::White
$btnStart.Location = New-Object System.Drawing.Point(20, 230)
$btnStart.Size = New-Object System.Drawing.Size(625, 45)
$btnStart.DialogResult = [System.Windows.Forms.DialogResult]::OK
$configForm.Controls.Add($btnStart)

$dialogRes = $configForm.ShowDialog()
if ($dialogRes -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "  [알림] 사용자가 작업을 취소하였습니다." -ForegroundColor Gray
    exit
}

$srcPath = $txtSrc.Text.Trim()
$tgtPath = $txtTgt.Text.Trim()
$prefixText = $txtPrefix.Text.Trim()
$configForm.Dispose()

# ------------------------------------------------------------------------------
# 2. 경로 유효성 검증
# ------------------------------------------------------------------------------
if (-not (Test-Path $srcPath)) {
    [System.Windows.Forms.MessageBox]::Show("대상 폴더 (A)가 존재하지 않습니다:`n$srcPath", "경로 오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}
if (-not (Test-Path $tgtPath)) {
    [System.Windows.Forms.MessageBox]::Show("타겟 상위 폴더 (B-1)가 존재하지 않습니다:`n$tgtPath", "경로 오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

Write-Host "  [2. 타겟 서브폴더 스캔 및 매핑]" -ForegroundColor Yellow
$targetSubfolders = @{}
Get-ChildItem -Path $tgtPath -Directory | ForEach-Object {
    if ($_.Name -match '^(\d{2})') {
        $prefix = $matches[1]
        $targetSubfolders[$prefix] = $_.FullName
    }
}
Write-Host "   * 감지된 타겟 접두사 폴더: $($targetSubfolders.Count)개" -ForegroundColor Green

# ------------------------------------------------------------------------------
# 3. 소스 파일 스캔, 복사 및 접두사 치환 실행
# ------------------------------------------------------------------------------
Write-Host ""
Write-Host "  [3. 파일 복사 및 치환 처리 시작]" -ForegroundColor Yellow

$successFiles = [System.Collections.Generic.List[PSCustomObject]]::new()
$skippedFiles = [System.Collections.Generic.List[PSCustomObject]]::new()

$sourceFiles = Get-ChildItem -Path $srcPath -File | Sort-Object Name
foreach ($file in $sourceFiles) {
    $fn = $file.Name
    if ($fn -notmatch '^(\d{2})') {
        $skippedFiles.Add([PSCustomObject]@{
            FileName = $fn
            Reason   = "파일명이 2자리 숫자로 시작하지 않음"
        })
        continue
    }

    $prefix = $matches[1]
    if (-not $targetSubfolders.ContainsKey($prefix)) {
        $skippedFiles.Add([PSCustomObject]@{
            FileName = $fn
            Reason   = "타겟 상위 폴더에 접두사 '$prefix' 폴더 없음"
        })
        continue
    }

    $tgtSubDir = $targetSubfolders[$prefix]
    $b2Dir = Join-Path $tgtSubDir "01 자료"

    if (-not (Test-Path $b2Dir)) {
        $folderName = Split-Path -Leaf $tgtSubDir
        $skippedFiles.Add([PSCustomObject]@{
            FileName = $fn
            Reason   = "'$folderName' 폴더 내 '01 자료' 폴더 없음"
        })
        continue
    }

    # 파일명 치환: 앞 2자리 제거 후 접두사 추가
    $newFn = $prefixText + $fn.Substring(2)
    $destFilePath = Join-Path $b2Dir $newFn

    try {
        Copy-Item -Path $file.FullName -Destination $destFilePath -Force
        $destItem = Get-Item $destFilePath
        if ($file.Length -eq $destItem.Length) {
            $successFiles.Add([PSCustomObject]@{
                FileName     = $fn
                OrigPath     = $file.FullName
                DestPath     = $destFilePath
                TargetFolder = (Split-Path -Leaf $tgtSubDir)
                NewFileName  = $newFn
                Size         = $file.Length
            })
            Write-Host "   [성공] $fn -> $newFn" -ForegroundColor Green
        } else {
            $skippedFiles.Add([PSCustomObject]@{
                FileName = $fn
                Reason   = "파일 크기 불일치"
            })
        }
    } catch {
        $skippedFiles.Add([PSCustomObject]@{
            FileName = $fn
            Reason   = "복사 에러: $_"
        })
    }
}

Write-Host ""
Write-Host "========================================================================================================" -ForegroundColor Cyan
Write-Host "   작업 완료 결과: 복사 성공 $($successFiles.Count)건 / 제외 $($skippedFiles.Count)건" -ForegroundColor Yellow
Write-Host "========================================================================================================" -ForegroundColor Cyan

# ------------------------------------------------------------------------------
# 4. 사용자 검토 및 선택적 삭제/롤백 다이얼로그 (Review & Action Dialog)
# ------------------------------------------------------------------------------
$reviewForm = New-Object System.Windows.Forms.Form
$reviewForm.Text = "📊 작업 결과 검토 및 후속 조치 선택 (02마감 v2.5)"
$reviewForm.Size = New-Object System.Drawing.Size(760, 640)
$reviewForm.StartPosition = "CenterScreen"
$reviewForm.TopMost = $true
$reviewForm.FormBorderStyle = "FixedDialog"
$reviewForm.MaximizeBox = $false

# 상단 배너 패널
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.BackColor = [System.Drawing.Color]::FromArgb(44, 62, 80)
$topPanel.Dock = "Top"
$topPanel.Height = 70
$reviewForm.Controls.Add($topPanel)

$lblBannerTitle = New-Object System.Windows.Forms.Label
$lblBannerTitle.Text = "월마감 파일 패턴폴더 복사 및 치환 결과 검토"
$lblBannerTitle.Font = New-Object System.Drawing.Font("맑은 고딕", 12, [System.Drawing.FontStyle]::Bold)
$lblBannerTitle.ForeColor = [System.Drawing.Color]::White
$lblBannerTitle.Location = New-Object System.Drawing.Point(15, 10)
$lblBannerTitle.Size = New-Object System.Drawing.Size(650, 25)
$topPanel.Controls.Add($lblBannerTitle)

$lblBannerSub = New-Object System.Windows.Forms.Label
$lblBannerSub.Text = "총 처리: $($successFiles.Count + $skippedFiles.Count)건  |  ✅ 복사 성공: $($successFiles.Count)건  |  ⚠️ 복사 제외: $($skippedFiles.Count)건"
$lblBannerSub.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$lblBannerSub.ForeColor = [System.Drawing.Color]::FromArgb(236, 240, 241)
$lblBannerSub.Location = New-Object System.Drawing.Point(15, 38)
$lblBannerSub.Size = New-Object System.Drawing.Size(650, 20)
$topPanel.Controls.Add($lblBannerSub)

# 중앙 보고서 텍스트 박스
$txtReport = New-Object System.Windows.Forms.TextBox
$txtReport.Multiline = $true
$txtReport.ScrollBars = "Vertical"
$txtReport.ReadOnly = $true
$txtReport.BackColor = [System.Drawing.Color]::FromArgb(252, 252, 252)
$txtReport.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtReport.Location = New-Object System.Drawing.Point(15, 85)
$txtReport.Size = New-Object System.Drawing.Size(715, 350)
$reviewForm.Controls.Add($txtReport)

# 보고서 내용 조립
$nl = [Environment]::NewLine
$rep = "=======================================================================" + $nl
$rep += " 월마감 파일 패턴폴더 복사 및 치환 상세 결과 보고서" + $nl
$rep += "=======================================================================" + $nl
$rep += "• 대상 폴더 (A)       : $srcPath" + $nl
$rep += "• 타겟 상위 폴더 (B-1): $tgtPath" + $nl
$rep += "• 복사 성공           : $($successFiles.Count)건" + $nl
$rep += "• 복사 제외(생략)     : $($skippedFiles.Count)건" + $nl + $nl

if ($successFiles.Count -gt 0) {
    $rep += "[✅ 복사 및 치환 성공 목록 ($($successFiles.Count)건)]" + $nl
    $rep += "-----------------------------------------------------------------------" + $nl
    $i = 1
    foreach ($item in $successFiles) {
        $rep += ("{0,2}. {1}" -f $i, $item.FileName) + $nl
        $rep += "    ➔ 저장: $($item.TargetFolder)\01 자료\$($item.NewFileName)" + $nl
        $rep += ("    ➔ 크기: {0:N0} bytes" -f $item.Size) + $nl
        $i++
    }
    $rep += $nl
}

if ($skippedFiles.Count -gt 0) {
    $rep += "[⚠️ 복사 제외/미매칭 목록 ($($skippedFiles.Count)건)]" + $nl
    $rep += "-----------------------------------------------------------------------" + $nl
    $i = 1
    foreach ($item in $skippedFiles) {
        $rep += ("{0,2}. {1}" -f $i, $item.FileName) + $nl
        $rep += "    ➔ 제외 사유: $($item.Reason)" + $nl
        $i++
    }
    $rep += $nl
}
$txtReport.Text = $rep

# 바로가기 버튼들
$btnOpenSrc = New-Object System.Windows.Forms.Button
$btnOpenSrc.Text = "📂 대상 폴더(A) 열기"
$btnOpenSrc.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$btnOpenSrc.Location = New-Object System.Drawing.Point(15, 442)
$btnOpenSrc.Size = New-Object System.Drawing.Size(160, 28)
$btnOpenSrc.Add_Click({ [System.Diagnostics.Process]::Start("explorer.exe", $srcPath) })
$reviewForm.Controls.Add($btnOpenSrc)

$btnOpenTgt = New-Object System.Windows.Forms.Button
$btnOpenTgt.Text = "📂 타겟 폴더(B-1) 열기"
$btnOpenTgt.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$btnOpenTgt.Location = New-Object System.Drawing.Point(185, 442)
$btnOpenTgt.Size = New-Object System.Drawing.Size(160, 28)
$btnOpenTgt.Add_Click({ [System.Diagnostics.Process]::Start("explorer.exe", $tgtPath) })
$reviewForm.Controls.Add($btnOpenTgt)

# 하단 액션 버튼 그룹 박스
$grpAction = New-Object System.Windows.Forms.GroupBox
$grpAction.Text = "⚡ 후속 조치 선택 (검토 후 원하시는 작업을 선택하세요)"
$grpAction.Font = New-Object System.Drawing.Font("맑은 고딕", 9, [System.Drawing.FontStyle]::Bold)
$grpAction.ForeColor = [System.Drawing.Color]::FromArgb(41, 128, 185)
$grpAction.Location = New-Object System.Drawing.Point(15, 480)
$grpAction.Size = New-Object System.Drawing.Size(715, 105)
$reviewForm.Controls.Add($grpAction)

# 1) 원본 삭제 버튼
$btnDelSrc = New-Object System.Windows.Forms.Button
$btnDelSrc.Text = "🗑️ [원본 정리]`n성공 파일 삭제"
$btnDelSrc.Font = New-Object System.Drawing.Font("맑은 고딕", 9, [System.Drawing.FontStyle]::Bold)
$btnDelSrc.BackColor = [System.Drawing.Color]::FromArgb(231, 76, 60)
$btnDelSrc.ForeColor = [System.Drawing.Color]::White
$btnDelSrc.Location = New-Object System.Drawing.Point(15, 25)
$btnDelSrc.Size = New-Object System.Drawing.Size(220, 65)
$btnDelSrc.Add_Click({
    if ($successFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("삭제할 성공한 원본 파일이 없습니다.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show("대상 폴더(A)에서 성공적으로 복사된 원본 파일 $($successFiles.Count)개를 삭제하시겠습니까?`n`n⚠️ 삭제된 파일은 복구되지 않습니다.`n대상 폴더: $srcPath", "원본 파일 삭제 확인", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        $delCount = 0
        foreach ($item in $successFiles) {
            if (Test-Path $item.OrigPath) {
                Remove-Item -Path $item.OrigPath -Force -ErrorAction SilentlyContinue
                $delCount++
            }
        }
        $remaining = (Get-ChildItem -Path $srcPath -File).Count
        if ($remaining -eq 0) {
            $delFolder = [System.Windows.Forms.MessageBox]::Show("원본 파일 $delCount개가 삭제되었습니다.`n`n현재 대상 폴더(A)가 비어 있습니다. 폴더 자체도 삭제하시겠습니까?", "대상 폴더 정리", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($delFolder -eq [System.Windows.Forms.DialogResult]::Yes) {
                Remove-Item -Path $srcPath -Recurse -Force -ErrorAction SilentlyContinue
                [System.Windows.Forms.MessageBox]::Show("대상 폴더(A)가 완전히 삭제되었습니다.", "완료", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("성공 원본 파일 $delCount개가 삭제되었습니다.`n(미매칭/제외 파일 $remaining개는 보존됨)", "완료", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        $reviewForm.Close()
    }
})
$grpAction.Controls.Add($btnDelSrc)

# 2) 롤백 버튼
$btnRollback = New-Object System.Windows.Forms.Button
$btnRollback.Text = "↩️ [롤백 / 되돌리기]`n복사본 파일 전체 삭제"
$btnRollback.Font = New-Object System.Drawing.Font("맑은 고딕", 9, [System.Drawing.FontStyle]::Bold)
$btnRollback.BackColor = [System.Drawing.Color]::FromArgb(230, 126, 34)
$btnRollback.ForeColor = [System.Drawing.Color]::White
$btnRollback.Location = New-Object System.Drawing.Point(245, 25)
$btnRollback.Size = New-Object System.Drawing.Size(220, 65)
$btnRollback.Add_Click({
    if ($successFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("롤백할 복사본 파일이 없습니다.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show("방금 목적지(01 자료)로 복사된 파일 $($successFiles.Count)개를 모두 삭제하여 작업 전 상태로 되돌리시겠습니까?`n`n타겟 상위 폴더: $tgtPath", "복사본 롤백 확인", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        $rbCount = 0
        foreach ($item in $successFiles) {
            if (Test-Path $item.DestPath) {
                Remove-Item -Path $item.DestPath -Force -ErrorAction SilentlyContinue
                $rbCount++
            }
        }
        [System.Windows.Forms.MessageBox]::Show("복사본 파일 $rbCount개가 안전하게 삭제(롤백)되었습니다.", "롤백 완료", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $reviewForm.Close()
    }
})
$grpAction.Controls.Add($btnRollback)

# 3) 유지 버튼
$btnKeep = New-Object System.Windows.Forms.Button
$btnKeep.Text = "✅ [작업 완료 및 유지]`n원본/복사본 모두 보존"
$btnKeep.Font = New-Object System.Drawing.Font("맑은 고딕", 9, [System.Drawing.FontStyle]::Bold)
$btnKeep.BackColor = [System.Drawing.Color]::FromArgb(39, 174, 96)
$btnKeep.ForeColor = [System.Drawing.Color]::White
$btnKeep.Location = New-Object System.Drawing.Point(475, 25)
$btnKeep.Size = New-Object System.Drawing.Size(225, 65)
$btnKeep.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("모든 파일이 원본 및 대상 폴더에 안전하게 보존되었습니다.", "작업 완료", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    $reviewForm.Close()
})
$grpAction.Controls.Add($btnKeep)

$reviewForm.ShowDialog() | Out-Null
$reviewForm.Dispose()

Write-Host "  [종료] 모든 처리가 정상 완료되었습니다." -ForegroundColor Green
Write-Host ""
