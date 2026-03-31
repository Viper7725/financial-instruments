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

$bgColor = [System.Drawing.Color]::FromArgb(244, 247, 245)
$cardColor = [System.Drawing.Color]::White
$cardBorderColor = [System.Drawing.Color]::FromArgb(220, 228, 223)
$primaryColor = [System.Drawing.Color]::FromArgb(24, 121, 92)
$primarySoftColor = [System.Drawing.Color]::FromArgb(227, 243, 238)
$dangerColor = [System.Drawing.Color]::FromArgb(191, 74, 62)
$dangerSoftColor = [System.Drawing.Color]::FromArgb(250, 232, 228)
$textColor = [System.Drawing.Color]::FromArgb(37, 49, 45)
$mutedTextColor = [System.Drawing.Color]::FromArgb(100, 112, 108)
$inputBackColor = [System.Drawing.Color]::FromArgb(251, 252, 251)

function Set-PrimaryButtonStyle {
    param(
        [System.Windows.Forms.Button]$Button
    )

    $Button.BackColor = $primaryColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(20, 103, 78)
    $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(18, 92, 70)
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function Set-SecondaryButtonStyle {
    param(
        [System.Windows.Forms.Button]$Button
    )

    $Button.BackColor = [System.Drawing.Color]::White
    $Button.ForeColor = $textColor
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = $cardBorderColor
    $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(245, 248, 246)
    $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(236, 242, 239)
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function Set-DangerButtonStyle {
    param(
        [System.Windows.Forms.Button]$Button
    )

    $Button.BackColor = $dangerColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(171, 63, 51)
    $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(156, 57, 47)
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function Set-TextBoxStyle {
    param(
        [System.Windows.Forms.TextBox]$TextBox
    )

    $TextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $TextBox.BackColor = $inputBackColor
    $TextBox.ForeColor = $textColor
}

function Set-CardStyle {
    param(
        [System.Windows.Forms.Panel]$Panel
    )

    $Panel.BackColor = $cardColor
    $Panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
}

function Set-StatusTheme {
    param(
        [ValidateSet('idle', 'running', 'success', 'error', 'stopping')]
        [string]$Theme
    )

    switch ($Theme) {
        'running' {
            $statusPanel.BackColor = $primarySoftColor
            $statusLabel.ForeColor = $primaryColor
        }
        'success' {
            $statusPanel.BackColor = $primarySoftColor
            $statusLabel.ForeColor = $primaryColor
        }
        'error' {
            $statusPanel.BackColor = $dangerSoftColor
            $statusLabel.ForeColor = $dangerColor
        }
        'stopping' {
            $statusPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 243, 224)
            $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(163, 95, 26)
        }
        default {
            $statusPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 245, 242)
            $statusLabel.ForeColor = $mutedTextColor
        }
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = U '\u6E38\u620F\u8D22\u52A1\u5BF9\u8D26\u5DE5\u5177'
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(920, 760)
$form.MinimumSize = New-Object System.Drawing.Size(920, 760)
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$form.BackColor = $bgColor
$form.MaximizeBox = $false

$iconPath = Join-Path -Path $scriptRoot -ChildPath "Start-ReconcileTool.ico"
if (Test-Path -LiteralPath $iconPath) {
    try {
        $form.Icon = New-Object System.Drawing.Icon($iconPath)
    }
    catch {
    }
}

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(18, 16)
$headerPanel.Size = New-Object System.Drawing.Size(868, 96)
$headerPanel.BackColor = $primaryColor
$form.Controls.Add($headerPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = U '\u672C\u5730\u5BF9\u8D26\u5DE5\u5177'
$titleLabel.Location = New-Object System.Drawing.Point(22, 18)
$titleLabel.Size = New-Object System.Drawing.Size(260, 32)
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 18, [System.Drawing.FontStyle]::Bold)
$headerPanel.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = U '\u9009\u62E9\u539F\u59CB\u8D26\u5355\u548C\u5206\u6210\u660E\u7EC6\u8868\uff0c\u7ED3\u679C\u4F1A\u81EA\u52A8\u6309\u6708\u4EFD\u5B58\u5165\u201C\u8F93\u51FA\u7ED3\u679C\u201D\u6587\u4EF6\u5939'
$subtitleLabel.Location = New-Object System.Drawing.Point(24, 56)
$subtitleLabel.Size = New-Object System.Drawing.Size(620, 22)
$subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(233, 245, 240)
$headerPanel.Controls.Add($subtitleLabel)

$headerTag = New-Object System.Windows.Forms.Label
$headerTag.Text = U '\u672C\u5730\u8FD0\u884C'
$headerTag.Location = New-Object System.Drawing.Point(748, 24)
$headerTag.Size = New-Object System.Drawing.Size(92, 30)
$headerTag.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$headerTag.BackColor = [System.Drawing.Color]::FromArgb(236, 246, 241)
$headerTag.ForeColor = $primaryColor
$headerTag.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9, [System.Drawing.FontStyle]::Bold)
$headerPanel.Controls.Add($headerTag)

function Add-FieldRow {
    param(
        [System.Windows.Forms.Control]$Parent,
        [string]$LabelText,
        [int]$Top,
        [string]$ButtonText
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $LabelText
    $label.Location = New-Object System.Drawing.Point(18, $Top)
    $label.Size = New-Object System.Drawing.Size(140, 24)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $label.ForeColor = $textColor
    $Parent.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(162, ($Top - 1))
    $textBox.Size = New-Object System.Drawing.Size(540, 28)
    Set-TextBoxStyle -TextBox $textBox
    $Parent.Controls.Add($textBox)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $ButtonText
    $button.Location = New-Object System.Drawing.Point(714, ($Top - 2))
    $button.Size = New-Object System.Drawing.Size(92, 30)
    Set-SecondaryButtonStyle -Button $button
    $Parent.Controls.Add($button)

    return [pscustomobject]@{
        Label   = $label
        TextBox = $textBox
        Button  = $button
    }
}

$fileCard = New-Object System.Windows.Forms.Panel
$fileCard.Location = New-Object System.Drawing.Point(18, 126)
$fileCard.Size = New-Object System.Drawing.Size(868, 218)
Set-CardStyle -Panel $fileCard
$form.Controls.Add($fileCard)

$fileCardTitle = New-Object System.Windows.Forms.Label
$fileCardTitle.Text = U '\u8D26\u5355\u6587\u4EF6'
$fileCardTitle.Location = New-Object System.Drawing.Point(18, 16)
$fileCardTitle.Size = New-Object System.Drawing.Size(180, 24)
$fileCardTitle.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 11, [System.Drawing.FontStyle]::Bold)
$fileCardTitle.ForeColor = $textColor
$fileCard.Controls.Add($fileCardTitle)

$fileCardIntro = New-Object System.Windows.Forms.Label
$fileCardIntro.Text = U '\u652F\u6301 Excel \u548C zip \u81EA\u52A8\u89E3\u538B\uff0c\u4FDD\u6301\u73B0\u6709\u5BF9\u8D26\u903B\u8F91\u4E0D\u53D8\u3002'
$fileCardIntro.Location = New-Object System.Drawing.Point(18, 42)
$fileCardIntro.Size = New-Object System.Drawing.Size(500, 20)
$fileCardIntro.ForeColor = $mutedTextColor
$fileCard.Controls.Add($fileCardIntro)

$inputRow = Add-FieldRow -Parent $fileCard -LabelText (U '\u539F\u59CB\u652F\u4ED8\u5B9D\u8D26\u5355') -Top 74 -ButtonText (U '\u9009\u62E9')
$shareStatementRow = Add-FieldRow -Parent $fileCard -LabelText (U '\u8D22\u52A1\u5206\u6210\u660E\u7EC6\u8868') -Top 116 -ButtonText (U '\u9009\u62E9')
$outputRow = Add-FieldRow -Parent $fileCard -LabelText (U '\u5BF9\u8D26\u7ED3\u679C\u4FDD\u5B58\u4F4D\u7F6E') -Top 158 -ButtonText (U '\u53E6\u5B58\u4E3A')

$hintLabel = New-Object System.Windows.Forms.Label
$hintLabel.Text = U '\u5206\u6210\u4EE5\u8D22\u52A1\u5206\u6210\u660E\u7EC6\u8868\u4E3A\u51C6\uff1B\u8F93\u51FA\u7ED3\u679C\u4F1A\u9ED8\u8BA4\u6309\u6708\u4EFD\u5F52\u6863\u3002'
$hintLabel.Location = New-Object System.Drawing.Point(18, 190)
$hintLabel.Size = New-Object System.Drawing.Size(600, 18)
$hintLabel.ForeColor = $mutedTextColor
$fileCard.Controls.Add($hintLabel)

$actionCard = New-Object System.Windows.Forms.Panel
$actionCard.Location = New-Object System.Drawing.Point(18, 358)
$actionCard.Size = New-Object System.Drawing.Size(868, 100)
Set-CardStyle -Panel $actionCard
$form.Controls.Add($actionCard)

$openAfterRun = New-Object System.Windows.Forms.CheckBox
$openAfterRun.Text = U '\u5B8C\u6210\u540E\u81EA\u52A8\u6253\u5F00\u7ED3\u679C\u6587\u4EF6'
$openAfterRun.Location = New-Object System.Drawing.Point(18, 16)
$openAfterRun.Size = New-Object System.Drawing.Size(240, 24)
$openAfterRun.Checked = $true
$openAfterRun.ForeColor = $textColor
$actionCard.Controls.Add($openAfterRun)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = U '\u5F00\u59CB\u5BF9\u8D26'
$runButton.Location = New-Object System.Drawing.Point(18, 50)
$runButton.Size = New-Object System.Drawing.Size(126, 34)
Set-PrimaryButtonStyle -Button $runButton
$actionCard.Controls.Add($runButton)

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Text = U '\u505C\u6B62\u5904\u7406'
$stopButton.Location = New-Object System.Drawing.Point(156, 50)
$stopButton.Size = New-Object System.Drawing.Size(126, 34)
Set-DangerButtonStyle -Button $stopButton
$actionCard.Controls.Add($stopButton)

$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Location = New-Object System.Drawing.Point(304, 18)
$statusPanel.Size = New-Object System.Drawing.Size(540, 58)
$statusPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$actionCard.Controls.Add($statusPanel)

$statusCaption = New-Object System.Windows.Forms.Label
$statusCaption.Text = U '\u5F53\u524D\u72B6\u6001'
$statusCaption.Location = New-Object System.Drawing.Point(16, 9)
$statusCaption.Size = New-Object System.Drawing.Size(90, 18)
$statusCaption.ForeColor = $mutedTextColor
$statusPanel.Controls.Add($statusCaption)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = U '\u5C31\u7EEA'
$statusLabel.Location = New-Object System.Drawing.Point(16, 27)
$statusLabel.Size = New-Object System.Drawing.Size(500, 22)
$statusLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10, [System.Drawing.FontStyle]::Bold)
$statusPanel.Controls.Add($statusLabel)

$logCard = New-Object System.Windows.Forms.Panel
$logCard.Location = New-Object System.Drawing.Point(18, 472)
$logCard.Size = New-Object System.Drawing.Size(868, 236)
Set-CardStyle -Panel $logCard
$form.Controls.Add($logCard)

$logTitle = New-Object System.Windows.Forms.Label
$logTitle.Text = U '\u8FD0\u884C\u65E5\u5FD7'
$logTitle.Location = New-Object System.Drawing.Point(18, 14)
$logTitle.Size = New-Object System.Drawing.Size(140, 22)
$logTitle.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 11, [System.Drawing.FontStyle]::Bold)
$logTitle.ForeColor = $textColor
$logCard.Controls.Add($logTitle)

$logSubtitle = New-Object System.Windows.Forms.Label
$logSubtitle.Text = U '\u8FD9\u91CC\u4F1A\u8BB0\u5F55\u672C\u6B21\u5BF9\u8D26\u7684\u5173\u952E\u8FDB\u5EA6\u548C\u5F02\u5E38\u4FE1\u606F\u3002'
$logSubtitle.Location = New-Object System.Drawing.Point(18, 38)
$logSubtitle.Size = New-Object System.Drawing.Size(420, 18)
$logSubtitle.ForeColor = $mutedTextColor
$logCard.Controls.Add($logSubtitle)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(18, 66)
$logBox.Size = New-Object System.Drawing.Size(830, 150)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 250)
$logBox.ForeColor = $textColor
$logBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$logBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$logCard.Controls.Add($logBox)

Set-StatusTheme -Theme 'idle'

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
    Set-StatusTheme -Theme 'stopping'
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
    if (-not $script:stopRequested) {
        Set-StatusTheme -Theme 'idle'
    }
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
    Set-StatusTheme -Theme 'running'
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
            Set-StatusTheme -Theme 'idle'
            [System.Windows.Forms.MessageBox]::Show(
                (U '\u5DF2\u624B\u52A8\u505C\u6B62\u5F53\u524D\u5BF9\u8D26\u4EFB\u52A1'),
                $form.Text,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
        elseif ($result.ExitCode -eq 0) {
            $statusLabel.Text = U '\u5BF9\u8D26\u5B8C\u6210'
            Set-StatusTheme -Theme 'success'
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
            Set-StatusTheme -Theme 'error'
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
        Set-StatusTheme -Theme 'error'
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
        Set-StatusTheme -Theme 'idle'
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
