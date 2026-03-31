Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function U {
    param(
        [string]$EscapedText
    )

    return [regex]::Unescape($EscapedText)
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$cliScriptPath = Join-Path -Path $scriptRoot -ChildPath "Reconcile-AlipayWorkbook.ps1"
$outputRoot = Join-Path -Path $scriptRoot -ChildPath (U '\u8F93\u51FA\u7ED3\u679C')
$processTimeoutSeconds = 600

if (-not (Test-Path -LiteralPath $outputRoot)) {
    New-Item -ItemType Directory -Path $outputRoot | Out-Null
}

if (-not (Test-Path -LiteralPath $cliScriptPath)) {
    [System.Windows.Forms.MessageBox]::Show(
        ("CLI script not found: {0}" -f $cliScriptPath),
        "Reconcile Tool",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

$form = New-Object System.Windows.Forms.Form
$form.Text = U '\u6E38\u620F\u8D22\u52A1\u5BF9\u8D26\u5DE5\u5177'
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(860, 730)
$form.MinimumSize = New-Object System.Drawing.Size(860, 730)
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)

$iconPath = Join-Path -Path $scriptRoot -ChildPath "Start-ReconcileTool.ico"
if (Test-Path -LiteralPath $iconPath) {
    try {
        $form.Icon = New-Object System.Drawing.Icon($iconPath)
    }
    catch {
    }
}

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = U '\u672C\u5730\u5BF9\u8D26\u5DE5\u5177'
$titleLabel.Location = New-Object System.Drawing.Point(20, 18)
$titleLabel.Size = New-Object System.Drawing.Size(220, 28)
$titleLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 15, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = U '\u9009\u62E9\u539F\u59CB\u8D26\u5355\u548C\u5206\u6210\u660E\u7EC6\u8868\uFF0C\u7ED3\u679C\u9ED8\u8BA4\u4FDD\u5B58\u5230\u201C\u8F93\u51FA\u7ED3\u679C\u201D\u6587\u4EF6\u5939'
$subtitleLabel.Location = New-Object System.Drawing.Point(22, 50)
$subtitleLabel.Size = New-Object System.Drawing.Size(800, 24)
$subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$form.Controls.Add($subtitleLabel)

function Add-FieldRow {
    param(
        [string]$LabelText,
        [int]$Top,
        [string]$ButtonText
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $LabelText
    $label.Location = New-Object System.Drawing.Point(22, $Top)
    $label.Size = New-Object System.Drawing.Size(130, 24)
    $label.TextAlign = "MiddleLeft"
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(160, $Top)
    $textBox.Size = New-Object System.Drawing.Size(560, 26)
    $form.Controls.Add($textBox)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $ButtonText
    $button.Location = New-Object System.Drawing.Point(732, ($Top - 1))
    $button.Size = New-Object System.Drawing.Size(90, 28)
    $form.Controls.Add($button)

    return [pscustomobject]@{
        Label   = $label
        TextBox = $textBox
        Button  = $button
    }
}

$inputRow = Add-FieldRow -LabelText (U '\u539F\u59CB\u652F\u4ED8\u5B9D\u8D26\u5355') -Top 100 -ButtonText (U '\u9009\u62E9')
$shareStatementRow = Add-FieldRow -LabelText (U '\u8D22\u52A1\u5206\u6210\u660E\u7EC6\u8868') -Top 146 -ButtonText (U '\u9009\u62E9')
$outputRow = Add-FieldRow -LabelText (U '\u5BF9\u8D26\u7ED3\u679C\u4FDD\u5B58\u4F4D\u7F6E') -Top 192 -ButtonText (U '\u53E6\u5B58\u4E3A')

$hintLabel = New-Object System.Windows.Forms.Label
$hintLabel.Text = U '\u5206\u6210\u4EE5\u8D22\u52A1\u5206\u6210\u660E\u7EC6\u8868\u4E3A\u51C6\uFF1B\u82E5\u53D1\u73B0\u201C\u9000\u6B3E\u7F3A\u5C11\u8D1F\u5206\u6210\u51B2\u56DE\u201D\uFF0C\u5DE5\u5177\u4F1A\u5728\u201C\u63A8\u7406\u610F\u89C1\u201Dsheet\u7ED9\u51FA\u53C2\u8003\u7ED3\u679C\uFF0C\u5E76\u53EF\u53C2\u8003\u201C\u8F93\u51FA\u7ED3\u679C\u201D\u4E2D\u7684\u5386\u53F2\u660E\u7EC6\uFF1B\u4E24\u4EFD\u539F\u59CB\u8868\u90FD\u652F\u6301 zip \u81EA\u52A8\u89E3\u538B'
$hintLabel.Location = New-Object System.Drawing.Point(160, 238)
$hintLabel.Size = New-Object System.Drawing.Size(640, 42)
$hintLabel.ForeColor = [System.Drawing.Color]::FromArgb(95, 95, 95)
$form.Controls.Add($hintLabel)

$openAfterRun = New-Object System.Windows.Forms.CheckBox
$openAfterRun.Text = U '\u5B8C\u6210\u540E\u81EA\u52A8\u6253\u5F00\u7ED3\u679C\u6587\u4EF6'
$openAfterRun.Location = New-Object System.Drawing.Point(160, 290)
$openAfterRun.Size = New-Object System.Drawing.Size(250, 28)
$openAfterRun.Checked = $true
$form.Controls.Add($openAfterRun)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = U '\u5F00\u59CB\u5BF9\u8D26'
$runButton.Location = New-Object System.Drawing.Point(22, 330)
$runButton.Size = New-Object System.Drawing.Size(120, 36)
$runButton.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
$runButton.ForeColor = [System.Drawing.Color]::White
$runButton.FlatStyle = "Flat"
$form.Controls.Add($runButton)

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Text = U '\u505C\u6B62\u5904\u7406'
$stopButton.Location = New-Object System.Drawing.Point(156, 330)
$stopButton.Size = New-Object System.Drawing.Size(120, 36)
$stopButton.BackColor = [System.Drawing.Color]::FromArgb(196, 64, 44)
$stopButton.ForeColor = [System.Drawing.Color]::White
$stopButton.FlatStyle = "Flat"
$form.Controls.Add($stopButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = U '\u5C31\u7EEA'
$statusLabel.Location = New-Object System.Drawing.Point(292, 337)
$statusLabel.Size = New-Object System.Drawing.Size(390, 24)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$form.Controls.Add($statusLabel)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(22, 385)
$logBox.Size = New-Object System.Drawing.Size(800, 290)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 248)
$form.Controls.Add($logBox)

function Append-Log {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return
    }

    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $logBox.AppendText(("[{0}] {1}{2}" -f $timestamp, $Text, [Environment]::NewLine))
}

function Quote-Argument {
    param(
        [string]$Value
    )

    if ($null -eq $Value) {
        return '""'
    }

    return '"' + $Value.Replace('"', '\"') + '"'
}

function Stop-RunningProcess {
    if ($null -eq $script:activeProcess) {
        return
    }

    try {
        if (-not $script:activeProcess.HasExited) {
            $script:activeProcess.Kill()
            $script:activeProcess.WaitForExit(5000) | Out-Null
        }
    }
    catch {
    }
}

function Request-StopProcess {
    if (-not $script:isRunning) {
        return
    }

    if ($script:stopRequested) {
        return
    }

    $script:stopRequested = $true
    $statusLabel.Text = U '\u6B63\u5728\u505C\u6B62\uff0C\u8BF7\u7A0D\u5019...'
    Append-Log (U '\u6B63\u5728\u624B\u52A8\u505C\u6B62\u5F53\u524D\u5BF9\u8D26\u4EFB\u52A1...')
    Stop-RunningProcess
}

function Get-ProcessFailureMessage {
    param(
        [pscustomobject]$Result
    )

    $streams = @()
    if ($null -ne $Result) {
        $streams += @($Result.StdOut, $Result.StdErr)
    }

    foreach ($stream in $streams) {
        if ([string]::IsNullOrWhiteSpace($stream)) {
            continue
        }

        $lines = @($stream -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        foreach ($line in $lines) {
            $trimmed = $line.Trim()
            if ($trimmed.StartsWith('ERROR:', [System.StringComparison]::OrdinalIgnoreCase)) {
                return $trimmed.Substring(6).Trim()
            }
        }
    }

    foreach ($stream in $streams) {
        if ([string]::IsNullOrWhiteSpace($stream)) {
            continue
        }

        $lines = @($stream -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $userFacingLines = @(
            $lines | ForEach-Object { $_.Trim() } | Where-Object {
                ($_ -notmatch '^(At line:|At I:|At [A-Z]:|\\+ )') -and
                ($_ -notmatch '^(CategoryInfo|FullyQualifiedErrorId)') -and
                ($_ -notmatch '^[~]+$') -and
                ($_ -notmatch '^ERROR:$')
            }
        )

        if ($userFacingLines.Count -gt 0) {
            return $userFacingLines[-1]
        }
    }

    return (U '\u5BF9\u8D26\u672A\u80FD\u5B8C\u6210\uff0c\u8BF7\u68C0\u67E5\u4E0A\u4F20\u7684\u6587\u4EF6\u662F\u5426\u6B63\u786E')
}

function Normalize-UserFacingErrorMessage {
    param(
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return (U '\u5BF9\u8D26\u672A\u80FD\u5B8C\u6210\uff0c\u8BF7\u68C0\u67E5\u4E0A\u4F20\u7684\u6587\u4EF6\u662F\u5426\u6B63\u786E')
    }

    $trimmed = $Message.Trim()
    if ($trimmed -match 'HRESULT E_FAIL|0x800A03EC|RPC_E_SERVERCALL_RETRYLATER|^\?+ COM ') {
        return (U '\u672C\u6B21 Excel \u5904\u7406\u5931\u8D25\u3002\u8BF7\u5148\u5173\u95ED\u76F8\u5173 Excel \u6587\u4EF6\u548C Excel \u7A97\u53E3\u540E\u91CD\u8BD5\uff1B\u5982\u679C\u4ECD\u5931\u8D25\uff0C\u8BF7\u5206\u522B\u624B\u52A8\u6253\u5F00\u4E0A\u4F20\u7684\u8D26\u5355\u548C\u5206\u6210\u660E\u7EC6\u8868\uff0C\u786E\u8BA4\u6CA1\u6709\u4FEE\u590D\u6216\u4FDD\u62A4\u63D0\u793A\u3002')
    }

    return $trimmed
}

function Get-DefaultOutputBaseName {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return (U '\u5BF9\u8D26\u7ED3\u679C')
    }

    $name = [System.IO.Path]::GetFileName($Path)
    if ($name.EndsWith(".zip", [System.StringComparison]::OrdinalIgnoreCase)) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($name)
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($name)
    $cleanupPatterns = @(
        '(?i)(?:[-_ ]reconcile(?:-(?:check|fixed))?)$',
        '(?i)(?:[-_ ](?:result|output|summary))$',
        '(?:[-_ ]?(?:\u5BF9\u8D26\u7ED3\u679C|\u5BF9\u8D26\u6C47\u603B|\u5BF9\u8D26|\u6C47\u603B|\u7ED3\u679C))$'
    )

    do {
        $changed = $false
        foreach ($pattern in $cleanupPatterns) {
            $trimmedName = [regex]::Replace($baseName, $pattern, '')
            if ($trimmedName -ne $baseName) {
                $baseName = $trimmedName.TrimEnd('-', '_', ' ')
                $changed = $true
            }
        }
    } while ($changed)

    if ([string]::IsNullOrWhiteSpace($baseName)) {
        return (U '\u5BF9\u8D26\u7ED3\u679C')
    }

    return $baseName
}

function Test-IsLikelyAutoManagedOutputPath {
    param(
        [string]$OutputPath
    )

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        return $true
    }

    $trimmed = $OutputPath.Trim()
    if ($trimmed -eq $script:lastAutoOutputPath) {
        return $true
    }

    $fileName = [System.IO.Path]::GetFileName($trimmed)
    if ($fileName -match '^(?i:reconcile(?:-\d{4}-\d{2})?\.xlsx)$') {
        return $true
    }

    if ($fileName -match '^(?:\u5BF9\u8D26\u7ED3\u679C(?:-\d{4}\u5E74\d{2}\u6708)?)\.xlsx$') {
        return $true
    }

    return $false
}

function Get-OutputPeriodInfo {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return [pscustomobject]@{
            FolderName   = (U '\u672A\u5206\u7C7B')
            DisplayLabel = $null
        }
    }

    $name = Get-DefaultOutputBaseName -Path $Path
    $patterns = @(
        '(?<!\d)(20\d{2})[-_.年](\d{1,2})(?:月)?',
        '(?<!\d)(\d{2})[-_.](\d{1,2})(?:月)?',
        '(?<!\d)(20\d{2})(\d{2})(?!\d)'
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($name, $pattern)
        if (-not $match.Success) {
            continue
        }

        $year = [int]$match.Groups[1].Value
        $month = [int]$match.Groups[2].Value

        if ($year -lt 100) {
            $year += 2000
        }

        if ($month -lt 1 -or $month -gt 12) {
            continue
        }

        return [pscustomobject]@{
            FolderName   = ('{0:D4}-{1:D2}' -f $year, $month)
            DisplayLabel = ('{0:D4}{1}{2:D2}{3}' -f $year, (U '\u5E74'), $month, (U '\u6708'))
        }
    }

    return [pscustomobject]@{
        FolderName   = (U '\u672A\u5206\u7C7B')
        DisplayLabel = $null
    }
}

function Get-OutputPeriodInfoFixed {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return [pscustomobject]@{
            FolderName   = 'unknown'
            DisplayLabel = $null
        }
    }

    $name = Get-DefaultOutputBaseName -Path $Path
    $patterns = @(
        '(?<!\d)(20\d{2})[年._-]?(\d{1,2})月?(?!\d)',
        '(?<!\d)(\d{2})[年._-]?(\d{1,2})月(?!\d)',
        '(?<!\d)(\d{2})[-_.](\d{1,2})(?!\d)',
        '(?<!\d)(20\d{2})(\d{2})(?!\d)'
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($name, $pattern)
        if (-not $match.Success) {
            continue
        }

        $year = [int]$match.Groups[1].Value
        $month = [int]$match.Groups[2].Value

        if ($year -lt 100) {
            $year += 2000
        }

        if ($month -lt 1 -or $month -gt 12) {
            continue
        }

        return [pscustomobject]@{
            FolderName   = ('{0:D4}-{1:D2}' -f $year, $month)
            DisplayLabel = ('{0:D4}{1}{2:D2}{3}' -f $year, (U '\u5E74'), $month, (U '\u6708'))
        }
    }

    return [pscustomobject]@{
        FolderName   = 'unknown'
        DisplayLabel = $null
    }
}

function Get-DefaultOutputPath {
    param(
        [string]$InputPath
    )

    $periodInfo = Get-OutputPeriodInfoFixed -Path $InputPath
    $monthFolderPath = Join-Path -Path $outputRoot -ChildPath $periodInfo.FolderName
    if (-not (Test-Path -LiteralPath $monthFolderPath)) {
        New-Item -ItemType Directory -Path $monthFolderPath | Out-Null
    }

    $fileName = if (-not [string]::IsNullOrWhiteSpace($periodInfo.DisplayLabel)) {
        ('reconcile-{0}.xlsx' -f $periodInfo.FolderName)
    } else {
        'reconcile.xlsx'
    }

    return Join-Path -Path $monthFolderPath -ChildPath $fileName
}

function Get-AutoOutputPath {
    param(
        [string]$InputPath,
        [string]$ShareStatementPath
    )

    $periodInfo = Get-OutputPeriodInfoFixed -Path $InputPath
    if ($periodInfo.FolderName -eq 'unknown' -and -not [string]::IsNullOrWhiteSpace($ShareStatementPath)) {
        $sharePeriodInfo = Get-OutputPeriodInfoFixed -Path $ShareStatementPath
        if ($sharePeriodInfo.FolderName -ne 'unknown') {
            $periodInfo = $sharePeriodInfo
        }
    }

    $monthFolderPath = Join-Path -Path $outputRoot -ChildPath $periodInfo.FolderName
    if (-not (Test-Path -LiteralPath $monthFolderPath)) {
        New-Item -ItemType Directory -Path $monthFolderPath | Out-Null
    }

    $fileName = if (-not [string]::IsNullOrWhiteSpace($periodInfo.DisplayLabel)) {
        ('reconcile-{0}.xlsx' -f $periodInfo.FolderName)
    } else {
        'reconcile.xlsx'
    }

    return Join-Path -Path $monthFolderPath -ChildPath $fileName
}

function Set-AutoOutputPath {
    param(
        [string]$InputPath,
        [string]$ShareStatementPath
    )

    $newPath = Get-AutoOutputPath -InputPath $InputPath -ShareStatementPath $ShareStatementPath
    $script:lastAutoOutputPath = $newPath
    $outputRow.TextBox.Text = $newPath
}

function Invoke-ReconcileProcess {
    param(
        [pscustomobject]$Payload
    )

    $argList = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $Payload.CliScriptPath,
        "-InputPath", $Payload.InputPath,
        "-OutputPath", $Payload.OutputPath
    )

    if (-not [string]::IsNullOrWhiteSpace($Payload.ShareStatementPath)) {
        $argList += @("-ShareStatementPath", $Payload.ShareStatementPath)
    }

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = "powershell.exe"
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $startInfo.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    $startInfo.Arguments = ($argList | ForEach-Object { Quote-Argument $_ }) -join " "

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo

    try {
        [void]$process.Start()
        $script:activeProcess = $process
        $startedAt = Get-Date

        while (-not $process.HasExited) {
            $elapsedSeconds = [int]((Get-Date) - $startedAt).TotalSeconds
            if ($elapsedSeconds -ge $processTimeoutSeconds) {
                Stop-RunningProcess
                throw ((U '\u5904\u7406\u8D85\u65F6\uff0C\u5DF2\u81EA\u52A8\u505C\u6B62\u3002\u8BF7\u68C0\u67E5\u4E0A\u4F20\u7684\u6587\u4EF6\u662F\u5426\u6B63\u786E\uff0C\u6216\u8005\u6362\u4E00\u4EFD\u65B0\u6587\u4EF6\u518D\u8BD5\u3002') + " (" + $processTimeoutSeconds + "s)")
            }

            $statusLabel.Text = ((U '\u6B63\u5728\u5904\u7406\uff0C\u8BF7\u7A0D\u5019...') + " " + $elapsedSeconds + "s")
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 200
        }

        $result = [pscustomobject]@{
            ExitCode = $process.ExitCode
            StdOut   = $process.StandardOutput.ReadToEnd()
            StdErr   = $process.StandardError.ReadToEnd()
            Output   = $Payload.OutputPath
            OpenFile = $Payload.OpenAfterRun
            Canceled = $script:stopRequested
        }

        if ($result.Canceled) {
            return $result
        }

        if ($result.ExitCode -eq 0 -and -not (Test-Path -LiteralPath $result.Output)) {
            throw (U '\u7A0B\u5E8F\u5DF2\u7ED3\u675F\uff0C\u4F46\u672A\u751F\u6210\u5BF9\u8D26\u7ED3\u679C\u6587\u4EF6\u3002\u8BF7\u68C0\u67E5\u4E0A\u4F20\u7684\u6587\u4EF6\u662F\u5426\u6B63\u786E\u3002')
        }

        return $result
    }
    finally {
        $script:activeProcess = $null
        $process.Dispose()
    }
}

function Update-RunButtonState {
    if ($script:isRunning) {
        $runButton.Enabled = $false
        $stopButton.Enabled = $true
        $inputRow.Button.Enabled = $false
        $outputRow.Button.Enabled = $false
        $shareStatementRow.Button.Enabled = $false
        return
    }

    $runButton.Enabled = $true
    $stopButton.Enabled = $false
    $inputRow.Button.Enabled = $true
    $outputRow.Button.Enabled = $true
    $shareStatementRow.Button.Enabled = $true
}

$inputDialog = New-Object System.Windows.Forms.OpenFileDialog
$inputDialog.Filter = "Excel Or Zip|*.xls;*.xlsx;*.xlsm;*.zip|Excel Files|*.xls;*.xlsx;*.xlsm|Zip Files|*.zip|All Files|*.*"
$inputDialog.Title = U '\u9009\u62E9\u539F\u59CB\u652F\u4ED8\u5B9D\u8D26\u5355'

$shareStatementDialog = New-Object System.Windows.Forms.OpenFileDialog
$shareStatementDialog.Filter = "Excel Or Zip|*.xls;*.xlsx;*.xlsm;*.zip|Excel Files|*.xls;*.xlsx;*.xlsm|Zip Files|*.zip|All Files|*.*"
$shareStatementDialog.Title = U '\u9009\u62E9\u8D22\u52A1\u5206\u6210\u660E\u7EC6\u8868'

$saveDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveDialog.Filter = "Excel Workbook|*.xlsx"
$saveDialog.Title = U '\u8BBE\u7F6E\u5BF9\u8D26\u7ED3\u679C\u4FDD\u5B58\u4F4D\u7F6E'
$saveDialog.InitialDirectory = $outputRoot
$saveDialog.FileName = "reconcile.xlsx"

$inputRow.TextBox.Add_TextChanged({
    if ($script:isRunning) {
        return
    }

    if (-not (
        [string]::IsNullOrWhiteSpace($outputRow.TextBox.Text) -or
        $script:outputPathAutoManaged -or
        (Test-IsLikelyAutoManagedOutputPath -OutputPath $outputRow.TextBox.Text)
    )) {
        return
    }

    $script:outputPathAutoManaged = $true
    Set-AutoOutputPath -InputPath $inputRow.TextBox.Text -ShareStatementPath $shareStatementRow.TextBox.Text
})

$shareStatementRow.TextBox.Add_TextChanged({
    if ($script:isRunning) {
        return
    }

    if (-not (
        [string]::IsNullOrWhiteSpace($outputRow.TextBox.Text) -or
        $script:outputPathAutoManaged -or
        (Test-IsLikelyAutoManagedOutputPath -OutputPath $outputRow.TextBox.Text)
    )) {
        return
    }

    $script:outputPathAutoManaged = $true
    Set-AutoOutputPath -InputPath $inputRow.TextBox.Text -ShareStatementPath $shareStatementRow.TextBox.Text
})

$inputRow.Button.Add_Click({
    if ($inputDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputRow.TextBox.Text = $inputDialog.FileName
        if (
            [string]::IsNullOrWhiteSpace($outputRow.TextBox.Text) -or
            $script:outputPathAutoManaged -or
            (Test-IsLikelyAutoManagedOutputPath -OutputPath $outputRow.TextBox.Text)
        ) {
            $script:outputPathAutoManaged = $true
            Set-AutoOutputPath -InputPath $inputDialog.FileName -ShareStatementPath $shareStatementRow.TextBox.Text
        }
    }
})

$outputRow.Button.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($outputRow.TextBox.Text)) {
        $saveDialog.FileName = [System.IO.Path]::GetFileName($outputRow.TextBox.Text)
        $initialDirectory = Split-Path -Parent $outputRow.TextBox.Text
        if ([string]::IsNullOrWhiteSpace($initialDirectory)) {
            $initialDirectory = $outputRoot
        }
        $saveDialog.InitialDirectory = $initialDirectory
    }

    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputRow.TextBox.Text = $saveDialog.FileName
        $script:outputPathAutoManaged = $false
    }
})

$shareStatementRow.Button.Add_Click({
    if ($shareStatementDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $shareStatementRow.TextBox.Text = $shareStatementDialog.FileName
    }
})

$stopButton.Add_Click({
    if (-not $script:isRunning) {
        return
    }

    Request-StopProcess
})

$runButton.Add_Click({
    if ($script:isRunning) {
        return
    }

    $inputPath = $inputRow.TextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($inputPath) -or -not (Test-Path -LiteralPath $inputPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            (U '\u8BF7\u9009\u62E9\u6709\u6548\u7684\u539F\u59CB\u652F\u4ED8\u5B9D\u8D26\u5355'),
            $form.Text,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $outputPath = $outputRow.TextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($outputPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            (U '\u8BF7\u8BBE\u7F6E\u5BF9\u8D26\u7ED3\u679C\u4FDD\u5B58\u4F4D\u7F6E'),
            $form.Text,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $shareStatementPath = $shareStatementRow.TextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($shareStatementPath)) {
        $shareStatementPath = $null
    }
    elseif (-not (Test-Path -LiteralPath $shareStatementPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            (U '\u8D22\u52A1\u5206\u6210\u660E\u7EC6\u8868\u6587\u4EF6\u4E0D\u5B58\u5728'),
            $form.Text,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if ($script:outputPathAutoManaged -or (Test-IsLikelyAutoManagedOutputPath -OutputPath $outputPath)) {
        $outputPath = Get-AutoOutputPath -InputPath $inputPath -ShareStatementPath $shareStatementPath
        $outputRow.TextBox.Text = $outputPath
        $script:lastAutoOutputPath = $outputPath
        $script:outputPathAutoManaged = $true
        Append-Log ((U '\u5DF2\u6309\u8D26\u5355\u6587\u4EF6\u540D\u91CD\u65B0\u751F\u6210\u5BFC\u51FA\u8DEF\u5F84') + ' ' + $outputPath)
    }

    $script:isRunning = $true
    $script:stopRequested = $false
    Update-RunButtonState
    $statusLabel.Text = U '\u6B63\u5728\u5904\u7406\uff0C\u8BF7\u7A0D\u5019...'
    Append-Log ((U '\u5F00\u59CB\u5904\u7406') + " " + $inputPath)
    if (-not [string]::IsNullOrWhiteSpace($shareStatementPath)) {
        Append-Log ((U '\u5206\u6210\u5C06\u4F7F\u7528\u8D22\u52A1\u5206\u6210\u660E\u7EC6\u8868') + " " + $shareStatementPath)
    } else {
        Append-Log (U '\u672A\u63D0\u4F9B\u8D22\u52A1\u5206\u6210\u660E\u7EC6\u8868\uff0c\u5206\u6210\u5217\u5C06\u4FDD\u6301\u7A7A\u767D')
    }

    $payload = [pscustomobject]@{
        CliScriptPath      = $cliScriptPath
        InputPath          = $inputPath
        OutputPath         = $outputPath
        ShareStatementPath = $shareStatementPath
        OpenAfterRun       = $openAfterRun.Checked
    }

    try {
        $result = Invoke-ReconcileProcess -Payload $payload

        if (-not [string]::IsNullOrWhiteSpace($result.StdOut)) {
            foreach ($line in ($result.StdOut -split "`r?`n")) {
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    Append-Log $line
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($result.StdErr)) {
            foreach ($line in ($result.StdErr -split "`r?`n")) {
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    Append-Log $line
                }
            }
        }

        if ($result.Canceled) {
            $statusLabel.Text = U '\u5DF2\u624B\u52A8\u505C\u6B62'
            [System.Windows.Forms.MessageBox]::Show(
                (U '\u5DF2\u624B\u52A8\u505C\u6B62\u5F53\u524D\u5BF9\u8D26\u4EFB\u52A1'),
                $form.Text,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
        elseif ($result.ExitCode -eq 0) {
            $statusLabel.Text = U '\u5BF9\u8D26\u5B8C\u6210'
            if ($result.OpenFile -and (Test-Path -LiteralPath $result.Output)) {
                Start-Process -FilePath $result.Output | Out-Null
            }
            [System.Windows.Forms.MessageBox]::Show(
                ((U '\u5DF2\u751F\u6210\u5BF9\u8D26\u6587\u4EF6') + [Environment]::NewLine + $result.Output),
                $form.Text,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
        else {
            $statusLabel.Text = U '\u8FD0\u884C\u5931\u8D25'
            $failureMessage = Normalize-UserFacingErrorMessage -Message (Get-ProcessFailureMessage -Result $result)
            Append-Log $failureMessage
            [System.Windows.Forms.MessageBox]::Show(
                $failureMessage,
                $form.Text,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    }
    catch {
        $statusLabel.Text = U '\u8FD0\u884C\u5931\u8D25'
        $failureMessage = Normalize-UserFacingErrorMessage -Message $_.Exception.Message
        Append-Log $failureMessage
        [System.Windows.Forms.MessageBox]::Show(
            $failureMessage,
            $form.Text,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        $script:isRunning = $false
        $script:stopRequested = $false
        Update-RunButtonState
    }
})

$form.Add_Shown({
    if ([string]::IsNullOrWhiteSpace($outputRow.TextBox.Text)) {
        Set-AutoOutputPath -InputPath $null -ShareStatementPath $null
    }
    $form.Activate()
})

$form.Add_FormClosing({
    if ($script:isRunning) {
        Append-Log (U '\u6B63\u5728\u7ED3\u675F\u5F53\u524D\u5BF9\u8D26\u8FDB\u7A0B...')
        Request-StopProcess
        $script:isRunning = $false
        $statusLabel.Text = U '\u5DF2\u505C\u6B62'
        Update-RunButtonState
    }
})

$script:isRunning = $false
$script:stopRequested = $false
$script:activeProcess = $null
$script:outputPathAutoManaged = $true
$script:lastAutoOutputPath = $null
Update-RunButtonState
[void]$form.ShowDialog()
