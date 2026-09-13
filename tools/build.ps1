param(
    [switch]$TemporaryTrust,
    [switch]$Test,
    [switch]$TestOnly,
    [switch]$ParserTestOnly,
    [switch]$MatcherTestOnly,
    [switch]$AIntegrityTestOnly,
    [switch]$BRowMatchTestOnly,
    [switch]$BConflictTestOnly,
    [switch]$AmountTestOnly,
    [switch]$GroupAmountTestOnly,
    [switch]$SnapshotTestOnly,
    [switch]$CompletedScopeTestOnly,
    [switch]$FullFallbackTestOnly,
    [switch]$EffectiveRelationsTestOnly,
    [switch]$EffectiveAmountTestOnly,
    [switch]$ShortSuffixPolicyTestOnly,
    [switch]$AggregateGateTestOnly,
    [switch]$AuditFindingsTestOnly,
    [switch]$FinalDecisionTestOnly
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$outputDir = Join-Path $root 'release'
$reportDir = Join-Path $root 'tests'
$securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$excel = $null
$book = $null
$oldTrustExists = $false
$oldTrust = $null
$changedTrust = $false
$securityKeyExisted = Test-Path -LiteralPath $securityPath

# 只创建本项目发布物，不安装到用户的 Personal.xlsb，也不覆盖已有发布文件。
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
$releaseFile = Join-Path $outputDir 'VATCheck.xlsm'
if ((Test-Path -LiteralPath $releaseFile) -and -not $TestOnly -and -not $ParserTestOnly -and -not $MatcherTestOnly -and -not $AIntegrityTestOnly -and -not $BRowMatchTestOnly -and -not $BConflictTestOnly -and -not $AmountTestOnly -and -not $GroupAmountTestOnly -and -not $SnapshotTestOnly -and -not $CompletedScopeTestOnly -and -not $FullFallbackTestOnly -and -not $EffectiveRelationsTestOnly -and -not $EffectiveAmountTestOnly -and -not $ShortSuffixPolicyTestOnly -and -not $AggregateGateTestOnly -and -not $AuditFindingsTestOnly -and -not $FinalDecisionTestOnly) { throw "发布文件已存在，请先移动或重命名：$releaseFile" }

try {
    if ($TemporaryTrust) {
        # 此开关必须由调用者明确授权；finally 精确恢复旧值或删除新增值。
        $settings = Get-ItemProperty -LiteralPath $securityPath -ErrorAction SilentlyContinue
        $oldTrustExists = $null -ne $settings -and $settings.PSObject.Properties.Name -contains 'AccessVBOM'
        if ($oldTrustExists) { $oldTrust = $settings.AccessVBOM }
        if (-not $securityKeyExisted) { New-Item -Path $securityPath -Force | Out-Null }
        New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null
        $changedTrust = $true
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Add(-4167)
    $project = $book.VBProject
    if ($null -eq $project) { throw 'Excel 不允许访问 VBA 工程。可按说明手动安装，或经授权使用 -TemporaryTrust 临时构建。' }

    if ($AggregateGateTestOnly) {
        # C6.1 只调用冻结VATScan；加载窗体仅满足Stage1编译依赖，不显示UI或导出release。
        foreach ($moduleName in @('modVATCheck', 'modVATStage2AggregateGate', 'modVATStage2AggregateTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2AggregateTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        $form = $project.VBComponents.Add(3)
        $form.Name = 'frmVATCheck'
        $form.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $root 'src/frmVATCheck.vba'), [Text.Encoding]::UTF8))
        $result = [string]$excel.Run('VATStage2AggregateGate_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-aggregate-gate-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C6.1 总体金额健康门自测失败，请读取 tests/stage2-aggregate-gate-test-results.txt。' }
        return
    }

    if ($FinalDecisionTestOnly) {
        # C6.3 测试夹具使用冻结上游，判定层只读取三个结构化摘要，不接入UI或release。
        foreach ($moduleName in @('modVATStage2Parser', 'modVATStage2Matcher', 'modVATStage2AIntegrity', 'modVATStage2BRowMatch', 'modVATStage2BConflict', 'modVATStage2ExcelSnapshot', 'modVATStage2CompletedScope', 'modVATStage2FullFallback', 'modVATStage2EffectiveRelations', 'modVATStage2Amount', 'modVATStage2GroupAmount', 'modVATStage2EffectiveAmount', 'modVATStage2ShortSuffixPolicy', 'modVATStage2AuditFindings', 'modVATCheck', 'modVATStage2AggregateGate', 'modVATStage2FinalDecision', 'modVATStage2FinalTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2FinalTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        # AggregateGate类型所在模块引用冻结Stage1窗体；只满足编译依赖，不显示UI。
        $form = $project.VBComponents.Add(3)
        $form.Name = 'frmVATCheck'
        $form.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $root 'src/frmVATCheck.vba'), [Text.Encoding]::UTF8))
        $result = [string]$excel.Run('VATStage2FinalDecision_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-final-decision-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C6.3 最终判定自测失败，请读取 tests/stage2-final-decision-test-results.txt。' }
        return
    }

    if ($AuditFindingsTestOnly) {
        # C6.2 测试夹具使用冻结上游，事项层只读取上游证据，不接入UI或release。
        foreach ($moduleName in @('modVATStage2Parser', 'modVATStage2Matcher', 'modVATStage2AIntegrity', 'modVATStage2BRowMatch', 'modVATStage2BConflict', 'modVATStage2ExcelSnapshot', 'modVATStage2CompletedScope', 'modVATStage2FullFallback', 'modVATStage2EffectiveRelations', 'modVATStage2Amount', 'modVATStage2GroupAmount', 'modVATStage2EffectiveAmount', 'modVATStage2ShortSuffixPolicy', 'modVATStage2AuditFindings', 'modVATStage2AuditTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2AuditTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        $result = [string]$excel.Run('VATStage2AuditFindings_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-audit-findings-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C6.2 审核事项自测失败，请读取 tests/stage2-audit-findings-test-results.txt。' }
        return
    }

    if ($ShortSuffixPolicyTestOnly) {
        # C5.3 测试夹具使用冻结上游，策略层不重算关系/金额，不接入UI或release。
        foreach ($moduleName in @('modVATStage2Parser', 'modVATStage2Matcher', 'modVATStage2AIntegrity', 'modVATStage2BRowMatch', 'modVATStage2BConflict', 'modVATStage2ExcelSnapshot', 'modVATStage2CompletedScope', 'modVATStage2FullFallback', 'modVATStage2EffectiveRelations', 'modVATStage2Amount', 'modVATStage2GroupAmount', 'modVATStage2EffectiveAmount', 'modVATStage2ShortSuffixPolicy', 'modVATStage2ShortPolicyTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2ShortPolicyTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        $result = [string]$excel.Run('VATStage2ShortSuffixPolicy_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-short-suffix-policy-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C5.3 短尾号策略自测失败，请读取 tests/stage2-short-suffix-policy-test-results.txt。' }
        return
    }

    if ($EffectiveAmountTestOnly) {
        # C5.2 只用内存快照测试统一组金额诊断，不加载 UI 或导出 release。
        foreach ($moduleName in @('modVATStage2Parser', 'modVATStage2Matcher', 'modVATStage2AIntegrity', 'modVATStage2BRowMatch', 'modVATStage2BConflict', 'modVATStage2ExcelSnapshot', 'modVATStage2CompletedScope', 'modVATStage2FullFallback', 'modVATStage2EffectiveRelations', 'modVATStage2Amount', 'modVATStage2GroupAmount', 'modVATStage2EffectiveAmount', 'modVATStage2EffAmountTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2EffAmountTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        $result = [string]$excel.Run('VATStage2EffectiveAmount_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-effective-amount-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C5.2 统一组金额自测失败，请读取 tests/stage2-effective-amount-test-results.txt。' }
        return
    }

    if ($EffectiveRelationsTestOnly) {
        # C5.1 只测试关系整合、完整性及冲突重扫，不加载金额或 UI。
        foreach ($moduleName in @('modVATStage2Parser', 'modVATStage2Matcher', 'modVATStage2AIntegrity', 'modVATStage2BRowMatch', 'modVATStage2BConflict', 'modVATStage2ExcelSnapshot', 'modVATStage2CompletedScope', 'modVATStage2FullFallback', 'modVATStage2EffectiveRelations', 'modVATStage2EffectiveTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2EffectiveTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        $result = [string]$excel.Run('VATStage2EffectiveRelations_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-effective-relations-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C5.1 关系整合自测失败，请读取 tests/stage2-effective-relations-test-results.txt。' }
        return
    }

    if ($FullFallbackTestOnly) {
        # C4.3 只使用临时宿主及内存快照，不加载金额/UI，不导出 release。
        foreach ($moduleName in @('modVATStage2Parser', 'modVATStage2Matcher', 'modVATStage2BRowMatch', 'modVATStage2BConflict', 'modVATStage2ExcelSnapshot', 'modVATStage2CompletedScope', 'modVATStage2FullFallback', 'modVATStage2FullFallbackTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2FullFallbackTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        $result = [string]$excel.Run('VATStage2FullFallback_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-full-fallback-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C4.3 全表fallback自测失败，请读取 tests/stage2-full-fallback-test-results.txt。' }
        return
    }

    if ($CompletedScopeTestOnly) {
        # C4.2 用内存快照测试完成范围映射，不读取业务工作簿，不加载金额或 UI。
        foreach ($moduleName in @('modVATStage2Parser', 'modVATStage2Matcher', 'modVATStage2BRowMatch', 'modVATStage2BConflict', 'modVATStage2ExcelSnapshot', 'modVATStage2CompletedScope', 'modVATStage2CompletedScopeTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2CompletedScopeTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        $result = [string]$excel.Run('VATStage2CompletedScope_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-completed-scope-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C4.2 完成范围自测失败，请读取 tests/stage2-completed-scope-test-results.txt。' }
        return
    }

    if ($SnapshotTestOnly) {
        # C4.1 只用临时工作簿；Stage 1 仅用于完成色计数 parity，不接入发布物。
        foreach ($moduleName in @('modVATCheck', 'modVATStage2ExcelSnapshot', 'modVATStage2SnapshotTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2SnapshotTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        $form = $project.VBComponents.Add(3)
        $form.Name = 'frmVATCheck'
        $form.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $root 'src/frmVATCheck.vba'), [Text.Encoding]::UTF8))
        $result = [string]$excel.Run('VATStage2Snapshot_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-snapshot-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C4.1 快照自测失败，请读取 tests/stage2-snapshot-test-results.txt。' }
        return
    }

    if ($GroupAmountTestOnly) {
        # C3.2 只加载关系/冲突/金额依赖与协调器测试，不接入 release。
        foreach ($moduleName in @('modVATStage2Parser', 'modVATStage2Matcher', 'modVATStage2BRowMatch', 'modVATStage2BConflict', 'modVATStage2Amount', 'modVATStage2GroupAmount', 'modVATStage2GroupAmountTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2GroupAmountTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        $result = [string]$excel.Run('VATStage2GroupAmount_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-group-amount-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C3.2 组金额自测失败，请读取 tests/stage2-group-amount-test-results.txt。' }
        return
    }

    if ($AmountTestOnly) {
        # C3.1 仅加载纯金额引擎和独立测试，不加载号码关系层或 Stage 1。
        foreach ($moduleName in @('modVATStage2Amount', 'modVATStage2AmountTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2AmountTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        $result = [string]$excel.Run('VATStage2Amount_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-amount-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C3.1 金额自测失败，请读取 tests/stage2-amount-test-results.txt。' }
        return
    }

    if ($BConflictTestOnly) {
        # C2.2c 仅测试已唯一匹配关联的冲突；不接入 release。
        foreach ($moduleName in @('modVATStage2Parser', 'modVATStage2Matcher', 'modVATStage2BRowMatch', 'modVATStage2BConflict', 'modVATStage2BConflictTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2BConflictTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        $result = [string]$excel.Run('VATStage2BConflict_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-b-conflict-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C2.2c 冲突自测失败，请读取 tests/stage2-b-conflict-test-results.txt。' }
        return
    }

    if ($BRowMatchTestOnly) {
        # C2.2b 只组合冻结 Parser/Matcher 与行汇总模块，不加载 Stage 1 或发布功能。
        foreach ($moduleName in @('modVATStage2Parser', 'modVATStage2Matcher', 'modVATStage2BRowMatch', 'modVATStage2BRowMatchTests')) {
            $core = $project.VBComponents.Add(1)
            $core.Name = $moduleName
            $moduleDir = if ($moduleName -eq 'modVATStage2BRowMatchTests') { 'tests' } else { 'src' }
            $moduleText = [IO.File]::ReadAllText((Join-Path $root "$moduleDir/$moduleName.bas"), [Text.Encoding]::UTF8)
            $core.CodeModule.AddFromString(($moduleText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        }
        $result = [string]$excel.Run('VATStage2BRowMatch_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-b-row-match-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C2.2b B 行匹配自测失败，请读取 tests/stage2-b-row-match-test-results.txt。' }
        return
    }

    if ($AIntegrityTestOnly) {
        # C2.2a 仅导入纯内存 A 扫描及独立测试，不加载或调度 B 相关模块。
        $core = $project.VBComponents.Add(1)
        $core.Name = 'modVATStage2AIntegrity'
        $integrityText = [IO.File]::ReadAllText((Join-Path $root 'src/modVATStage2AIntegrity.bas'), [Text.Encoding]::UTF8)
        $core.CodeModule.AddFromString(($integrityText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        $testModule = $project.VBComponents.Add(1)
        $testModule.Name = 'modVATStage2AIntegrityTests'
        $testText = [IO.File]::ReadAllText((Join-Path $root 'tests/modVATStage2AIntegrityTests.bas'), [Text.Encoding]::UTF8)
        $testModule.CodeModule.AddFromString(($testText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        $result = [string]$excel.Run('VATStage2AIntegrity_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-a-integrity-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C2.2a A 完整性自测失败，请读取 tests/stage2-a-integrity-test-results.txt。' }
        return
    }

    if ($MatcherTestOnly) {
        # C2.1 只加载 Matcher 和独立测试，不调度 Parser，不进入正式发布物。
        $core = $project.VBComponents.Add(1)
        $core.Name = 'modVATStage2Matcher'
        $matcherText = [IO.File]::ReadAllText((Join-Path $root 'src/modVATStage2Matcher.bas'), [Text.Encoding]::UTF8)
        $core.CodeModule.AddFromString(($matcherText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        $testModule = $project.VBComponents.Add(1)
        $testModule.Name = 'modVATStage2MatcherTests'
        $testText = [IO.File]::ReadAllText((Join-Path $root 'tests/modVATStage2MatcherTests.bas'), [Text.Encoding]::UTF8)
        $testModule.CodeModule.AddFromString(($testText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        $result = [string]$excel.Run('VATStage2Matcher_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-matcher-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C2.1 Matcher 自测失败，请读取 tests/stage2-matcher-test-results.txt。' }
        return
    }

    if ($ParserTestOnly) {
        # C1 只加载独立 Parser 和测试，不依赖 Stage 1，也不进入正式发布物。
        $core = $project.VBComponents.Add(1)
        $core.Name = 'modVATStage2Parser'
        $parserText = [IO.File]::ReadAllText((Join-Path $root 'src/modVATStage2Parser.bas'), [Text.Encoding]::UTF8)
        $core.CodeModule.AddFromString(($parserText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        $testModule = $project.VBComponents.Add(1)
        $testModule.Name = 'modVATStage2ParserTests'
        $testText = [IO.File]::ReadAllText((Join-Path $root 'tests/modVATStage2ParserTests.bas'), [Text.Encoding]::UTF8)
        $testModule.CodeModule.AddFromString(($testText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        $result = [string]$excel.Run('VATStage2Parser_SelfTest')
        [IO.File]::WriteAllText((Join-Path $reportDir 'stage2-parser-test-results.txt'), $result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'C1 Parser 自测失败，请读取 tests/stage2-parser-test-results.txt。' }
        return
    }

    $core = $project.VBComponents.Add(1)
    $core.Name = 'modVATCheck'
    $coreText = [IO.File]::ReadAllText((Join-Path $root 'src/modVATCheck.bas'), [Text.Encoding]::UTF8)
    $core.CodeModule.AddFromString(($coreText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
    $form = $project.VBComponents.Add(3)
    $form.Name = 'frmVATCheck'
    $form.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $root 'src/frmVATCheck.vba'), [Text.Encoding]::UTF8))

    if ($Test -or $TestOnly) {
        $testModule = $project.VBComponents.Add(1)
        $testModule.Name = 'modVATTests'
        $testText = [IO.File]::ReadAllText((Join-Path $root 'tests/modVATTests.bas'), [Text.Encoding]::UTF8)
        $testModule.CodeModule.AddFromString(($testText -replace '(?m)^Attribute VB_Name = .*\r?\n', ''))
        # 本机新建工作簿默认名为 Sheet1，带工作簿限定会被误解析为工作表引用；新实例中该宏唯一，直接用宏名调用。
        $result = $excel.Run("VATCheck_SelfTest")
        [IO.File]::WriteAllText((Join-Path $reportDir 'excel-test-results.txt'), [string]$result, [Text.UTF8Encoding]::new($true))
        Write-Output $result
        if (-not ([string]$result).StartsWith('PASS:')) { throw 'Excel 原生自测失败，请读取 tests/excel-test-results.txt。' }
        $project.VBComponents.Remove($testModule)
    }
    if ($TestOnly) { return }

    # 由真实 Excel 导出窗体和资源文件，避免手工拼接 FRX 格式。
    $core.Export((Join-Path $outputDir 'modVATCheck.bas'))
    $form.Export((Join-Path $outputDir 'frmVATCheck.frm'))
    $sheet = $book.Worksheets(1)
    $sheet.Name = '使用说明'
    $sheet.Range('A1').Value2 = '增值税专票双表一键核对工具'
    $sheet.Range('A3').Value2 = '按 Alt+F8，选择 VATCheck_Start，点击运行。'
    $sheet.Range('A4').Value2 = '点击两侧“选择 Excel”选择文件；当前实例已打开的文件直接复用，再选择 Sheet。'
    $sheet.Range('A5').Value2 = '选中一个正确填色的单元格，回到工具读取颜色，然后点击一键核对。'
    $sheet.Range('A6').Value2 = '完成颜色两表共用；只读取目标列，金额与颜色修改后需再次点击核对。'
    $sheet.Range('A7').Value2 = '长期使用：将 release 内的 bas 和 frm 导入 Personal.xlsb（frx 必须与 frm 同目录）。'
    $sheet.Range('A8').Value2 = '核对功能不修改数据、颜色，不保存或关闭工作簿。'
    $sheet.Range('A10').Value2 = '详细安装与第一次测试：请阅读随附 README.md。'
    $sheet.Columns.Item(1).ColumnWidth = 110
    $sheet.Range('A1:A10').Font.Name = 'Microsoft YaHei UI'
    $sheet.Range('A1:A10').Font.Size = 12
    $sheet.Range('A1').Font.Size = 20
    $sheet.Range('A1').Font.Bold = $true
    $sheet.Rows('1:10').RowHeight = 28
    $book.SaveAs($releaseFile, 52)
    Write-Output "RELEASE=$releaseFile"
} finally {
    # 仅清理由此脚本创建的独立实例及临时工作簿。
    if ($null -ne $book) { try { $book.Close($false) } catch { Write-Warning $_ } }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { Write-Warning $_ }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    # 释放构建产生的 COM 引用后再恢复设置，避免测试实例延迟退出。
    foreach ($reference in @($sheet, $form, $core, $testModule, $project, $book)) {
        if ($null -ne $reference -and [Runtime.InteropServices.Marshal]::IsComObject($reference)) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference) } catch { Write-Warning $_ }
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($changedTrust) {
        if ($oldTrustExists) {
            New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value $oldTrust -PropertyType DWord -Force | Out-Null
        } else {
            Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction Stop
        }
        Write-Output 'VBA_TRUST_RESTORED'
    }
}
