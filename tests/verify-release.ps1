param([switch]$TemporaryTrust)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$releaseDir = Join-Path $root 'release'
$securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$excel = $null
$book = $null
$changedTrust = $false
$log = New-Object System.Collections.Generic.List[string]

try {
    if ($TemporaryTrust) {
        # 沿用已获授权的临时 VBA 访问；不改变 PowerShell 执行策略。
        $settings = Get-ItemProperty -LiteralPath $securityPath -ErrorAction SilentlyContinue
        $hadTrust = $null -ne $settings -and $settings.PSObject.Properties.Name -contains 'AccessVBOM'
        if ($hadTrust) { $oldTrust = $settings.AccessVBOM }
        if (-not (Test-Path -LiteralPath $securityPath)) { New-Item -Path $securityPath -Force | Out-Null }
        New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null
        $changedTrust = $true
    }
    $testText = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'modVATTests.bas'), [Text.Encoding]::UTF8)
    $testText = $testText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
    foreach ($mode in @('REOPEN_XLSM', 'IMPORT_BAS_FRM_FRX')) {
        # 每种发布形式使用独立实例，避免已关闭的同名 VBA 工程仍被 COM 引用。
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        if ($mode -eq 'REOPEN_XLSM') {
            $book = $excel.Workbooks.Open((Join-Path $releaseDir 'VATCheck.xlsm'), 0, $true)
        } else {
            $book = $excel.Workbooks.Add(-4167)
            $book.VBProject.VBComponents.Import((Join-Path $releaseDir 'modVATCheck.bas')) | Out-Null
            $book.VBProject.VBComponents.Import((Join-Path $releaseDir 'frmVATCheck.frm')) | Out-Null
        }
        # 两种发布形式都运行完整回归；不保存对发布文件的内存测试改动。
        $component = $book.VBProject.VBComponents.Add(1)
        $component.Name = 'modVATTests'
        $component.CodeModule.AddFromString($testText)
        $result = [string]$excel.Run('VATCheck_SelfTest')
        $log.Add($mode + ': ' + ($result -split "`r?`n")[0])
        if (-not $result.StartsWith('PASS:')) {
            $log.Add($result)
            throw "$mode 验证失败"
        }
        $log.Add($result)
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
        $component = $null
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
        $book = $null
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        $excel = $null
    }
    $log.Add('PASS: 发布版重新打开和原生模块/窗体导入均可运行；测试改动未保存。')
} finally {
    if ($null -ne $book) { try { $book.Close($false) } catch { Write-Warning $_ } }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { Write-Warning $_ }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($changedTrust) {
        if ($hadTrust) {
            New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value $oldTrust -PropertyType DWord -Force | Out-Null
        } else {
            Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction Stop
        }
        $log.Add('VBA_TRUST_RESTORED')
    }
    [IO.File]::WriteAllText((Join-Path $PSScriptRoot 'release-verification.txt'), ($log -join "`r`n"), [Text.UTF8Encoding]::new($true))
    $log | Where-Object { $_ -notmatch '^PASS: \d+ assertions[\r\n]' } | Write-Output
}
