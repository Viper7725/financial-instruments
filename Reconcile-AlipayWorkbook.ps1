param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath 'reconcile.xlsx'),

    [string]$ShareStatementPath
)

Set-StrictMode -Version 3
$ErrorActionPreference = "Stop"
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

try {
    $utf8Encoding = New-Object System.Text.UTF8Encoding($false)
    [Console]::InputEncoding = $utf8Encoding
    [Console]::OutputEncoding = $utf8Encoding
    $OutputEncoding = $utf8Encoding
}
catch {
}

$ProductPattern = '^(.*?)(\d+(?:\.\d+)?)\u5143(?:\u6E38\u620F)?\u793C\u5305$'
$TradeIncomePrefix = '0010001|'
$RefundPrefix = '0070002|'
$FeePrefix = '0030003|'
$BaseFeePrefix = '0030130|'

function U {
    param(
        [string]$EscapedText
    )

    return [regex]::Unescape($EscapedText)
}

function Get-ReadableExceptionDetail {
    param(
        [System.Exception]$Exception
    )

    if ($null -eq $Exception -or [string]::IsNullOrWhiteSpace($Exception.Message)) {
        return $null
    }

    $message = $Exception.Message.Trim()
    if ($message -match 'HRESULT|COM|0x800A03EC|RPC_E_SERVERCALL_RETRYLATER') {
        return $null
    }

    return $message
}

function Get-FriendlyExcelOperationMessage {
    param(
        [string]$Operation,
        [string]$Path,
        [System.Exception]$Exception
    )

    $pathSuffix = if ([string]::IsNullOrWhiteSpace($Path)) { '' } else { ' ' + (U '\u6587\u4EF6') + ': ' + $Path }
    $detail = Get-ReadableExceptionDetail -Exception $Exception
    $detailSuffix = if ([string]::IsNullOrWhiteSpace($detail)) { '' } else { ' ' + (U '\u8BE6\u7EC6\u4FE1\u606F') + ': ' + $detail }

    switch ($Operation) {
        'CreateExcelApplication' {
            return (U '\u65E0\u6CD5\u542F\u52A8 Excel\u3002\u8BF7\u786E\u8BA4\u672C\u673A\u5DF2\u5B89\u88C5 Microsoft Excel\uff0c\u5E76\u5148\u5173\u95ED\u5DF2\u5361\u4F4F\u7684 Excel \u8FDB\u7A0B\u540E\u91CD\u8BD5\u3002') + $detailSuffix
        }
        'OpenInputWorkbook' {
            return (U '\u65E0\u6CD5\u6253\u5F00\u539F\u59CB\u652F\u4ED8\u5B9D\u8D26\u5355\u3002\u8BF7\u5148\u5173\u95ED\u8BE5\u6587\u4EF6\u548C\u6240\u6709 Excel \u7A97\u53E3\u540E\u91CD\u8BD5\uff1B\u5982\u679C\u6587\u4EF6\u521A\u4ECE\u538B\u7F29\u5305\u89E3\u51FA\uff0C\u4E5F\u53EF\u4EE5\u5148\u624B\u52A8\u7528 Excel \u6253\u5F00\u4E00\u6B21\uff0C\u786E\u8BA4\u6CA1\u6709\u4FEE\u590D\u6216\u4FDD\u62A4\u63D0\u793A\u3002') + $pathSuffix + $detailSuffix
        }
        'OpenShareWorkbook' {
            return (U '\u65E0\u6CD5\u6253\u5F00\u8D22\u52A1\u5206\u6210\u660E\u7EC6\u8868\u3002\u8BF7\u5148\u5173\u95ED\u8BE5\u6587\u4EF6\u548C\u6240\u6709 Excel \u7A97\u53E3\u540E\u91CD\u8BD5\uff0C\u5E76\u786E\u8BA4\u8BE5\u6587\u4EF6\u80FD\u591F\u88AB Excel \u6B63\u5E38\u6253\u5F00\u3002') + $pathSuffix + $detailSuffix
        }
        'OpenHistoricalWorkbook' {
            return (U '\u65E0\u6CD5\u6253\u5F00\u5386\u53F2\u5BF9\u8D26\u53C2\u8003\u6587\u4EF6') + $pathSuffix + $detailSuffix
        }
        'SaveWorkbook' {
            return (U '\u65E0\u6CD5\u4FDD\u5B58\u5BF9\u8D26\u7ED3\u679C\u3002\u8BF7\u786E\u8BA4\u76EE\u6807\u6587\u4EF6\u6CA1\u6709\u88AB Excel \u6216\u5176\u4ED6\u7A0B\u5E8F\u5360\u7528\uff0C\u4E14\u4FDD\u5B58\u76EE\u5F55\u5177\u6709\u5199\u5165\u6743\u9650\u3002') + $pathSuffix + $detailSuffix
        }
        default {
            return (U '\u5BF9\u8D26\u5904\u7406\u5931\u8D25\uff0C\u8BF7\u68C0\u67E5 Excel \u6587\u4EF6\u662F\u5426\u53EF\u4EE5\u6B63\u5E38\u6253\u5F00\u3002') + $pathSuffix + $detailSuffix
        }
    }
}

function Open-WorkbookSafely {
    param(
        $ExcelApp,
        [string]$Path,
        [string]$Operation
    )

    $lastException = $null
    foreach ($attempt in 1..2) {
        try {
            return $ExcelApp.Workbooks.Open($Path, 0, $true)
        }
        catch {
            $lastException = $_.Exception
            if ($attempt -lt 2) {
                Start-Sleep -Milliseconds 500
            }
        }
    }

    throw (Get-FriendlyExcelOperationMessage -Operation $Operation -Path $Path -Exception $lastException)
}

function Save-WorkbookSafely {
    param(
        $Workbook,
        [string]$OutputPath
    )

    $directory = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $lastException = $null
    foreach ($attempt in 1..2) {
        try {
            $Workbook.SaveAs($OutputPath, 51)
            return
        }
        catch {
            $lastException = $_.Exception
            if ($attempt -lt 2) {
                Start-Sleep -Milliseconds 500
            }
        }
    }

    throw (Get-FriendlyExcelOperationMessage -Operation 'SaveWorkbook' -Path $OutputPath -Exception $lastException)
}

trap {
    $message = $null
    if ($null -ne $_ -and $null -ne $_.Exception) {
        $message = Get-ReadableExceptionDetail -Exception $_.Exception
        if ([string]::IsNullOrWhiteSpace($message)) {
            $message = Get-FriendlyExcelOperationMessage -Operation 'Generic' -Path $null -Exception $_.Exception
        }
    }

    if ([string]::IsNullOrWhiteSpace($message)) {
        $message = U '\u5BF9\u8D26\u5904\u7406\u5931\u8D25\uff0C\u8BF7\u68C0\u67E5\u4E0A\u4F20\u7684\u6587\u4EF6\u662F\u5426\u6B63\u786E\u3002'
    }

    Write-Output ("ERROR: " + $message)
    exit 1
}

$HistoricalOutputRoot = Join-Path -Path $ScriptRoot -ChildPath (U '\u8F93\u51FA\u7ED3\u679C')

$OtherCategoryLabel = U '\u5176\u4ED6'
$TransferToWangshangLabel = U '\u8F6C\u51FA\u5230\u7F51\u5546\u94F6\u884C'
$TransferToWangshangKeyword = U '\u8F6C\u51FA\u5230\u7F51\u5546\u94F6\u884C'
$CommissionRemarkKeyword = U '\u5929\u732B\u4F63\u91D1'
$BaseSoftwareFeeRemarkKeyword = U '\u57FA\u7840\u8F6F\u4EF6\u670D\u52A1\u8D39'
$CategorySystemKeyword = U '\u7C7B\u76EE\u7CFB\u7EDF'
$DeductKeyword = U '\u6263\u6B3E'

function Normalize-GameName {
    param(
        [string]$GameName
    )

    if ([string]::IsNullOrWhiteSpace($GameName)) {
        return $GameName
    }

    $normalized = $GameName.Trim()
    $normalized = [regex]::Replace($normalized, '(?:\s*[\(\[]?(?:test|TEST|娴嬭瘯)\s*\d+(?:\.\d+)*[\)\]]?)$', '')
    $testKeyword = [regex]::Escape((U '\u6D4B\u8BD5'))
    $normalized = [regex]::Replace($normalized, ("(?:\s*[\(\[]?(?:test|TEST|{0})\s*\d+(?:\.\d+)*[\)\]]?)$" -f $testKeyword), '')
    $normalized = [regex]::Replace($normalized, '\s+', ' ')
    $normalized = $normalized.Trim()

    return $normalized
}

function Get-NumericValue {
    param(
        [Parameter(ValueFromPipeline = $true)]
        $Value
    )

    if ($null -eq $Value) {
        return 0.0
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return 0.0
    }

    return [double]$text.Trim()
}

function Get-ProductInfo {
    param(
        [string]$ProductName
    )

    if ([string]::IsNullOrWhiteSpace($ProductName)) {
        return $null
    }

    $trimmed = $ProductName.Trim()
    $match = [regex]::Match($trimmed, $ProductPattern)
    if ($match.Success) {
        return [pscustomobject]@{
            Game  = Normalize-GameName -GameName $match.Groups[1].Value
            Price = [double]$match.Groups[2].Value
        }
    }

    return [pscustomobject]@{
        Game  = Normalize-GameName -GameName $trimmed
        Price = $null
    }
}

function Get-OrderIdFromRemark {
    param(
        [string]$Remark
    )

    if ([string]::IsNullOrWhiteSpace($Remark)) {
        return $null
    }

    $match = [regex]::Match($Remark, '\{(?<OrderId>\d{10,})\}')
    if ($match.Success) {
        return $match.Groups['OrderId'].Value
    }

    $match = [regex]::Match($Remark, '[\(\uFF08](?<OrderId>\d{10,})[\)\uFF09]')
    if ($match.Success) {
        return $match.Groups['OrderId'].Value
    }

    return $null
}

function Test-IsRemarkDerivedFee {
    param(
        [string]$BusinessDesc,
        [string]$Remark,
        [string]$RemarkOrderId
    )

    if (-not [string]::IsNullOrWhiteSpace($BusinessDesc)) {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($Remark) -or [string]::IsNullOrWhiteSpace($RemarkOrderId)) {
        return $false
    }

    return (
        $Remark.Contains($CommissionRemarkKeyword) -and
        $Remark.Contains($CategorySystemKeyword) -and
        $Remark.Contains($DeductKeyword)
    )
}

function Test-IsRemarkDerivedBaseFee {
    param(
        [string]$BusinessDesc,
        [string]$Remark
    )

    if (-not [string]::IsNullOrWhiteSpace($BusinessDesc)) {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($Remark)) {
        return $false
    }

    return (
        $Remark.Contains($BaseSoftwareFeeRemarkKeyword) -and
        $Remark.Contains($DeductKeyword)
    )
}

function Get-ServiceFeeHourBucketFromRemark {
    param(
        [string]$Remark
    )

    if ([string]::IsNullOrWhiteSpace($Remark)) {
        return $null
    }

    $match = [regex]::Match($Remark, '\{(?<Bucket>\d{10})\}')
    if ($match.Success) {
        return $match.Groups['Bucket'].Value
    }

    return $null
}

function Get-TimeHourBucket {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $match = [regex]::Match($Text.Trim(), '^(?<Year>\d{4})[-/](?<Month>\d{1,2})[-/](?<Day>\d{1,2})\s+(?<Hour>\d{1,2})')
    if (-not $match.Success) {
        return $null
    }

    return ('{0}{1:D2}{2:D2}{3:D2}' -f [int]$match.Groups['Year'].Value, [int]$match.Groups['Month'].Value, [int]$match.Groups['Day'].Value, [int]$match.Groups['Hour'].Value)
}

function Add-BucketAmount {
    param(
        [hashtable]$BucketAmounts,
        [string]$Bucket,
        [double]$Amount
    )

    if ([string]::IsNullOrWhiteSpace($Bucket) -or [math]::Abs($Amount) -lt 0.000001) {
        return
    }

    if (-not $BucketAmounts.ContainsKey($Bucket)) {
        $BucketAmounts[$Bucket] = @{}
    }

    $amountKey = ('{0:F2}' -f [double]$Amount)
    if (-not $BucketAmounts[$Bucket].ContainsKey($amountKey)) {
        $BucketAmounts[$Bucket][$amountKey] = 0
    }

    $BucketAmounts[$Bucket][$amountKey] += 1
}

function Resolve-UploadPath {
    param(
        [string]$Path,
        [string[]]$AllowedExtensions,
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Label not found: $Path"
    }

    $item = Get-Item -LiteralPath $Path
    if ($item.PSIsContainer) {
        throw "$Label must be a file: $Path"
    }

    $extension = [System.IO.Path]::GetExtension($item.Name).ToLowerInvariant()
    if ($AllowedExtensions -contains $extension) {
        return [pscustomobject]@{
            OriginalPath = $item.FullName
            ResolvedPath = $item.FullName
            CleanupPath  = $null
            Notice       = $null
        }
    }

    if ($extension -ne '.zip') {
        throw ("{0} must be one of: {1}, or a .zip archive" -f $Label, ($AllowedExtensions -join ', '))
    }

    $tempRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "XiaoManReconcile"
    if (-not (Test-Path -LiteralPath $tempRoot)) {
        New-Item -ItemType Directory -Path $tempRoot | Out-Null
    }

    $extractPath = Join-Path -Path $tempRoot -ChildPath ([Guid]::NewGuid().ToString("N"))
    Expand-Archive -LiteralPath $item.FullName -DestinationPath $extractPath -Force

    $candidates = @(Get-ChildItem -LiteralPath $extractPath -Recurse -File | Where-Object {
        $candidateExtension = [System.IO.Path]::GetExtension($_.Name).ToLowerInvariant()
        ($AllowedExtensions -contains $candidateExtension) -and
        ($_.Name -notlike '~$*') -and
        ($_.FullName -notmatch '\\__MACOSX(\\|$)')
    })

    if ($candidates.Count -eq 0) {
        Remove-Item -LiteralPath $extractPath -Recurse -Force -ErrorAction SilentlyContinue
        throw ("No supported file found inside archive for {0}: {1}" -f $Label, $Path)
    }

    $selected = $candidates | Sort-Object -Property @{ Expression = 'Length'; Descending = $true }, @{ Expression = 'FullName'; Descending = $false } | Select-Object -First 1

    $notice = if ($candidates.Count -gt 1) {
        "{0}压缩包中存在多个候选文件，已自动使用体积最大的文件：{1}" -f $Label, $selected.FullName
    } else {
        "已从压缩包中解压{0}：{1}" -f $Label, $selected.FullName
    }

    return [pscustomobject]@{
        OriginalPath = $item.FullName
        ResolvedPath = $selected.FullName
        CleanupPath  = $extractPath
        Notice       = $notice
    }
}

function Ensure-SummaryRow {
    param(
        [hashtable]$Summary,
        [string]$Game
    )

    if (-not $Summary.ContainsKey($Game)) {
        $Summary[$Game] = [ordered]@{
            Category = $Game
            Income   = 0.0
            Expense  = 0.0
            Share    = $null
            Fee      = 0.0
            BaseFee  = 0.0
        }
    }
}

function Get-SortedCategoryNames {
    param(
        [hashtable]$Summary
    )

    $specialOrder = @(
        $OtherCategoryLabel,
        $TransferToWangshangLabel
    )

    $all = @($Summary.Keys)
    $normal = $all | Where-Object { $specialOrder -notcontains $_ } | Sort-Object
    $special = foreach ($name in $specialOrder) {
        if ($Summary.ContainsKey($name)) {
            $name
        }
    }

    return @($normal + $special)
}

function Load-ShareStatementSummary {
    param(
        [string]$Path,
        $ExcelApp
    )

    $result = @{
        Totals      = @{}
        Games       = @{}
        Rows        = New-Object 'System.Collections.Generic.List[object]'
        OrderToGame = @{}
        PositiveAmountsByOrder = @{}
        NegativeAmountsByOrder = @{}
        GamesByOrder = @{}
        PositiveAmountsByGamePrice = @{}
        MatchedWorksheetCount = 0
        MatchedRowCount = 0
    }

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $result
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "财务分成明细表文件不存在：$Path"
    }

    $shareWorkbook = $null
    try {
        $shareWorkbook = Open-WorkbookSafely -ExcelApp $ExcelApp -Path $Path -Operation 'OpenShareWorkbook'

        foreach ($worksheet in $shareWorkbook.Worksheets) {
            $usedRange = $worksheet.UsedRange
            try {
                $values = $usedRange.Value2
                $rowCount = $usedRange.Rows.Count
                $columnCount = $usedRange.Columns.Count

                if ($rowCount -lt 2 -or $columnCount -lt 2) {
                    continue
                }

                $headerMap = @{}
                for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
                    $headerText = [string]$values[1, $columnIndex]
                    if (-not [string]::IsNullOrWhiteSpace($headerText)) {
                        $headerMap[$headerText.Trim()] = $columnIndex
                    }
                }

                if (-not $headerMap.ContainsKey((U '\u5546\u54C1\u540D\u79F0')) -or -not $headerMap.ContainsKey((U '\u6263\u8D39\u91D1\u989D'))) {
                    continue
                }

                $result.MatchedWorksheetCount += 1

                $productColumn = $headerMap[(U '\u5546\u54C1\u540D\u79F0')]
                $shareAmountColumn = $headerMap[(U '\u6263\u8D39\u91D1\u989D')]
                $orderColumn = if ($headerMap.ContainsKey((U '\u4EA4\u6613\u4E3B\u8BA2\u5355\u53F7'))) {
                    $headerMap[(U '\u4EA4\u6613\u4E3B\u8BA2\u5355\u53F7')]
                } else {
                    $null
                }

                for ($rowIndex = 2; $rowIndex -le $rowCount; $rowIndex++) {
                    $rawProductName = [string]$values[$rowIndex, $productColumn]
                    $productInfo = Get-ProductInfo -ProductName $rawProductName
                    if ($null -eq $productInfo -or [string]::IsNullOrWhiteSpace($productInfo.Game)) {
                        continue
                    }

                    $shareAmount = Get-NumericValue $values[$rowIndex, $shareAmountColumn]
                    $orderId = if ($null -ne $orderColumn) {
                        $worksheet.Cells.Item($rowIndex, $orderColumn).Text.Trim()
                    } else {
                        $null
                    }

                    if (-not $result.Totals.ContainsKey($productInfo.Game)) {
                        $result.Totals[$productInfo.Game] = 0.0
                    }

                    $result.Totals[$productInfo.Game] += $shareAmount
                    $result.Games[$productInfo.Game] = $true
                    $result.MatchedRowCount += 1
                    $result.Rows.Add([pscustomobject]@{
                        Worksheet   = $worksheet.Name
                        RowIndex    = $rowIndex
                        OrderId     = $orderId
                        ProductName = $rawProductName
                        Game        = $productInfo.Game
                        Price       = $productInfo.Price
                        Amount      = $shareAmount
                    }) | Out-Null

                    if (-not [string]::IsNullOrWhiteSpace($orderId) -and -not $result.OrderToGame.ContainsKey($orderId)) {
                        $result.OrderToGame[$orderId] = $productInfo.Game
                    }

                    if (-not [string]::IsNullOrWhiteSpace($orderId)) {
                        if (-not $result.GamesByOrder.ContainsKey($orderId)) {
                            $result.GamesByOrder[$orderId] = New-Object 'System.Collections.Generic.HashSet[string]'
                        }

                        [void]$result.GamesByOrder[$orderId].Add($productInfo.Game)
                    }

                    if (-not [string]::IsNullOrWhiteSpace($orderId) -and $shareAmount -gt 0.0) {
                        if (-not $result.PositiveAmountsByOrder.ContainsKey($orderId)) {
                            $result.PositiveAmountsByOrder[$orderId] = 0.0
                        }

                        $result.PositiveAmountsByOrder[$orderId] += $shareAmount
                    }

                    if ($shareAmount -gt 0.0 -and $null -ne $productInfo.Price) {
                        $gamePriceKey = ('{0}|{1:F2}' -f $productInfo.Game, [double]$productInfo.Price)
                        if (-not $result.PositiveAmountsByGamePrice.ContainsKey($gamePriceKey)) {
                            $result.PositiveAmountsByGamePrice[$gamePriceKey] = [ordered]@{
                                Count   = 0
                                Amounts = @{}
                            }
                        }

                        $amountKey = ('{0:F2}' -f [double]$shareAmount)
                        $result.PositiveAmountsByGamePrice[$gamePriceKey].Count += 1
                        $result.PositiveAmountsByGamePrice[$gamePriceKey].Amounts[$amountKey] = [double]$shareAmount
                    }

                    if (-not [string]::IsNullOrWhiteSpace($orderId) -and $shareAmount -lt 0.0) {
                        if (-not $result.NegativeAmountsByOrder.ContainsKey($orderId)) {
                            $result.NegativeAmountsByOrder[$orderId] = 0.0
                        }

                        $result.NegativeAmountsByOrder[$orderId] += $shareAmount
                    }
                }
            }
            finally {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null
            }
        }
    }
    finally {
        if ($shareWorkbook) {
            $shareWorkbook.Close($false) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shareWorkbook) | Out-Null
        }
    }

    return $result
}

function Add-ShareAmount {
    param(
        [hashtable]$Summary,
        [string]$Game,
        [double]$Amount
    )

    Ensure-SummaryRow -Summary $Summary -Game $Game
    if ($null -eq $Summary[$Game].Share) {
        $Summary[$Game].Share = 0.0
    }

    $Summary[$Game].Share += $Amount
}

function Add-Diagnostic {
    param(
        $Diagnostics,
        [string]$Type,
        [string]$Status,
        [string]$Game,
        [string]$OrderId,
        $Amount,
        [string]$Source,
        [string]$Message
    )

    $Diagnostics.Add([pscustomobject]@{
        Type    = $Type
        Status  = $Status
        Game    = $Game
        OrderId = $OrderId
        Amount  = $Amount
        Source  = $Source
        Message = $Message
    }) | Out-Null
}

function Copy-Summary {
    param(
        [hashtable]$Summary
    )

    $copy = @{}
    foreach ($game in $Summary.Keys) {
        $row = $Summary[$game]
        $copy[$game] = [ordered]@{
            Category = [string]$row['Category']
            Income   = [double]$row['Income']
            Expense  = [double]$row['Expense']
            Share    = if ($null -eq $row['Share']) { $null } else { [double]$row['Share'] }
            Fee      = [double]$row['Fee']
            BaseFee  = [double]$row['BaseFee']
        }
    }

    return $copy
}

function Add-ReferenceShareProfile {
    param(
        [hashtable]$Profiles,
        [string]$Game,
        [double]$Price,
        [double]$ShareAmount,
        [string]$Source
    )

    if ([string]::IsNullOrWhiteSpace($Game) -or $Price -le 0.0 -or $ShareAmount -le 0.0) {
        return
    }

    $key = ('{0}|{1:F2}' -f (Normalize-GameName -GameName $Game), $Price)
    if (-not $Profiles.ContainsKey($key)) {
        $Profiles[$key] = [ordered]@{
            Count   = 0
            Amounts = @{}
            Sources = New-Object 'System.Collections.Generic.List[string]'
        }
    }

    $amountKey = ('{0:F2}' -f $ShareAmount)
    $Profiles[$key].Count += 1
    if (-not $Profiles[$key].Amounts.ContainsKey($amountKey)) {
        $Profiles[$key].Amounts[$amountKey] = [double]$ShareAmount
    }

    if (-not [string]::IsNullOrWhiteSpace($Source)) {
        $Profiles[$key].Sources.Add($Source) | Out-Null
    }
}

function Load-HistoricalInferenceProfiles {
    param(
        [string]$RootPath,
        [string]$CurrentOutputPath,
        $ExcelApp
    )

    $profiles = @{}

    if ([string]::IsNullOrWhiteSpace($RootPath) -or -not (Test-Path -LiteralPath $RootPath)) {
        return $profiles
    }

    $currentOutputFullPath = if ([string]::IsNullOrWhiteSpace($CurrentOutputPath)) { $null } else { [System.IO.Path]::GetFullPath($CurrentOutputPath) }
    $candidateFiles = @(Get-ChildItem -LiteralPath $RootPath -Recurse -File -Filter *.xlsx | Where-Object {
        if ($null -eq $currentOutputFullPath) {
            return $true
        }

        return [System.IO.Path]::GetFullPath($_.FullName) -ne $currentOutputFullPath
    })

    foreach ($file in $candidateFiles) {
        $workbook = $null
        try {
            $workbook = Open-WorkbookSafely -ExcelApp $ExcelApp -Path $file.FullName -Operation 'OpenHistoricalWorkbook'
            $inferenceSheet = $null
            foreach ($worksheet in $workbook.Worksheets) {
                if ($worksheet.Name -eq (U '\u63A8\u7406\u610F\u89C1')) {
                    $inferenceSheet = $worksheet
                    break
                }
            }

            if ($null -eq $inferenceSheet) {
                continue
            }

            $usedRange = $inferenceSheet.UsedRange
            try {
                $rowCount = $usedRange.Rows.Count
                $headerRow = $null
                for ($rowIndex = 1; $rowIndex -le $rowCount; $rowIndex++) {
                    $firstCell = $inferenceSheet.Cells.Item($rowIndex, 1).Text.Trim()
                    $secondCell = $inferenceSheet.Cells.Item($rowIndex, 2).Text.Trim()
                    if ($firstCell -eq (U '\u6E38\u620F') -and $secondCell -eq (U '\u8BA2\u5355\u53F7')) {
                        $headerRow = $rowIndex
                        break
                    }
                }

                if ($null -eq $headerRow) {
                    continue
                }

                for ($rowIndex = $headerRow + 1; $rowIndex -le $rowCount; $rowIndex++) {
                    $game = Normalize-GameName -GameName $inferenceSheet.Cells.Item($rowIndex, 1).Text.Trim()
                    $refundAmountText = $inferenceSheet.Cells.Item($rowIndex, 3).Text.Trim()
                    $positiveShareText = $inferenceSheet.Cells.Item($rowIndex, 4).Text.Trim()
                    if ([string]::IsNullOrWhiteSpace($game) -or [string]::IsNullOrWhiteSpace($refundAmountText) -or [string]::IsNullOrWhiteSpace($positiveShareText)) {
                        continue
                    }

                    $refundAmount = Get-NumericValue $refundAmountText
                    $positiveShare = Get-NumericValue $positiveShareText
                    Add-ReferenceShareProfile -Profiles $profiles -Game $game -Price $refundAmount -ShareAmount $positiveShare -Source $file.FullName
                }
            }
            finally {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null
            }
        }
        catch {
            $message = Get-ReadableExceptionDetail -Exception $_.Exception
            if ([string]::IsNullOrWhiteSpace($message)) {
                $message = Get-FriendlyExcelOperationMessage -Operation 'OpenHistoricalWorkbook' -Path $file.FullName -Exception $_.Exception
            }
            Write-Output ("已跳过历史结果参考文件：{0}" -f $message)
        }
        finally {
            if ($workbook) {
                $workbook.Close($false) | Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
        }
    }

    return $profiles
}

function Test-IsRecognizedBusinessDescription {
    param(
        [string]$BusinessDesc
    )

    if ([string]::IsNullOrWhiteSpace($BusinessDesc)) {
        return $false
    }

    return (
        $BusinessDesc.StartsWith($TradeIncomePrefix, [System.StringComparison]::Ordinal) -or
        $BusinessDesc.StartsWith($RefundPrefix, [System.StringComparison]::Ordinal) -or
        $BusinessDesc.StartsWith($FeePrefix, [System.StringComparison]::Ordinal) -or
        $BusinessDesc.StartsWith($BaseFeePrefix, [System.StringComparison]::Ordinal)
    )
}

function Save-SummaryWorkbook {
    param(
        [hashtable]$Summary,
        [hashtable]$InferredSummary,
        [hashtable]$InferredAdjustmentsByGame,
        $InferenceEntries,
        [string]$OutputPath,
        $Diagnostics,
        $ExcelApp
    )

    $headers = @(
        (U '\u5206\u7C7B'),
        (U '\u6C42\u548C\u9879:\u6536\u5165\uff08+\u5143\uff09'),
        (U '\u6C42\u548C\u9879:\u652F\u51FA\uff08-\u5143\uff09'),
        (U '\u6C42\u548C\u9879:\u5206\u6210'),
        (U '\u6C42\u548C\u9879:\u624B\u7EED\u8D39'),
        (U '\u6C42\u548C\u9879:\u57FA\u7840\u8F6F\u4EF6\u8D39')
    )

    $directory = Split-Path -Path $OutputPath -Parent
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory | Out-Null
    }

    $workbook = $ExcelApp.Workbooks.Add()
    try {
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = U '\u5BF9\u8D26\u6C47\u603B'

        for ($col = 0; $col -lt $headers.Count; $col++) {
            $sheet.Cells.Item(1, $col + 1).Value2 = $headers[$col]
        }

        $gameNames = Get-SortedCategoryNames -Summary $Summary
        $rowIndex = 2
        foreach ($gameName in $gameNames) {
            $row = $Summary[$gameName]
            try {
                $sheet.Cells.Item($rowIndex, 1).Value2 = [string]$row['Category']
                $sheet.Cells.Item($rowIndex, 2).Value2 = [double]$row['Income']
                $sheet.Cells.Item($rowIndex, 3).Value2 = [double]$row['Expense']
                if ($null -ne $row['Share']) {
                    $sheet.Cells.Item($rowIndex, 4).Value2 = [double]$row['Share']
                }
                $sheet.Cells.Item($rowIndex, 5).Value2 = [double]$row['Fee']
                $sheet.Cells.Item($rowIndex, 6).Value2 = [double]$row['BaseFee']
            }
            catch {
                Write-Output ("SAVE_ROW_ERROR game={0} row={1} categoryType={2} incomeType={3} expenseType={4} feeType={5} baseFeeType={6}" -f $gameName, $rowIndex, $row['Category'].GetType().FullName, $row['Income'].GetType().FullName, $row['Expense'].GetType().FullName, $row['Fee'].GetType().FullName, $row['BaseFee'].GetType().FullName)
                throw
            }
            $rowIndex++
        }

        $sheet.Range("A1:F1").Font.Bold = $true
        $sheet.Range("B:F").NumberFormat = "0.00"
        $sheet.Columns.AutoFit() | Out-Null

        if ($null -ne $InferredSummary) {
            $inferenceHeaders = @(
                (U '\u5206\u7C7B'),
                (U '\u6C42\u548C\u9879:\u6536\u5165\uff08+\u5143\uff09'),
                (U '\u6C42\u548C\u9879:\u652F\u51FA\uff08-\u5143\uff09'),
                (U '\u6C42\u548C\u9879:\u5206\u6210\uff08\u63A8\u7406\u540E\uff09'),
                (U '\u6C42\u548C\u9879:\u624B\u7EED\u8D39'),
                (U '\u6C42\u548C\u9879:\u57FA\u7840\u8F6F\u4EF6\u8D39'),
                (U '\u539F\u59CB\u5206\u6210'),
                (U '\u63A8\u7406\u8C03\u6574:\u5206\u6210'),
                (U '\u63A8\u7406\u8BF4\u660E')
            )

            $inferenceSheet = $workbook.Worksheets.Add([System.Type]::Missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
            $inferenceSheet.Name = U '\u63A8\u7406\u610F\u89C1'
            $inferenceSheet.Cells.Item(1, 1).Value2 = U '\u4EE5\u4E0B\u662F\u6839\u636E\u5206\u6210\u660E\u7EC6\u7F3A\u6F0F\u6216\u7F3A\u5C11\u51B2\u56DE\u7B49\u60C5\u51B5\u81EA\u52A8\u63A8\u7406\u7684\u53C2\u8003\u7ED3\u679C\uFF0C\u539F\u59CB\u7ED3\u679C\u4ECD\u4FDD\u7559\u5728\u201C\u5BF9\u8D26\u6C47\u603B\u201D\u3002'

            for ($col = 0; $col -lt $inferenceHeaders.Count; $col++) {
                $inferenceSheet.Cells.Item(2, $col + 1).Value2 = $inferenceHeaders[$col]
            }

            $gameNames = Get-SortedCategoryNames -Summary $InferredSummary
            $rowIndex = 3
            foreach ($gameName in $gameNames) {
                $row = $InferredSummary[$gameName]
                $originalRow = if ($Summary.ContainsKey($gameName)) { $Summary[$gameName] } else { $null }
                $originalShare = if ($null -ne $originalRow) { $originalRow['Share'] } else { $null }
                $adjustmentAmount = 0.0
                $explanation = $null

                if ($InferredAdjustmentsByGame.ContainsKey($gameName)) {
                    $adjustmentAmount = [double]$InferredAdjustmentsByGame[$gameName].Amount
                    $explanation = ('{0}{1}{2}' -f (U '\u5DF2\u81EA\u52A8\u63A8\u7406 '), $InferredAdjustmentsByGame[$gameName].Count, (U ' \u7B14\u5206\u6210\u8C03\u6574'))
                }

                $inferenceSheet.Cells.Item($rowIndex, 1).Value2 = [string]$row['Category']
                $inferenceSheet.Cells.Item($rowIndex, 2).Value2 = [double]$row['Income']
                $inferenceSheet.Cells.Item($rowIndex, 3).Value2 = [double]$row['Expense']
                if ($null -ne $row['Share']) {
                    $inferenceSheet.Cells.Item($rowIndex, 4).Value2 = [double]$row['Share']
                }
                $inferenceSheet.Cells.Item($rowIndex, 5).Value2 = [double]$row['Fee']
                $inferenceSheet.Cells.Item($rowIndex, 6).Value2 = [double]$row['BaseFee']
                if ($null -ne $originalShare) {
                    $inferenceSheet.Cells.Item($rowIndex, 7).Value2 = [double]$originalShare
                }
                $inferenceSheet.Cells.Item($rowIndex, 8).Value2 = [double]$adjustmentAmount
                if (-not [string]::IsNullOrWhiteSpace($explanation)) {
                    $inferenceSheet.Cells.Item($rowIndex, 9).Value2 = $explanation
                }
                $rowIndex++
            }

            $detailTitleRow = $rowIndex + 2
            $detailHeaderRow = $detailTitleRow + 1
            $inferenceSheet.Cells.Item($detailTitleRow, 1).Value2 = U '\u81EA\u52A8\u63A8\u7406\u660E\u7EC6'

            $detailHeaders = @(
                (U '\u6E38\u620F'),
                (U '\u8BA2\u5355\u53F7'),
                (U '\u8BA2\u5355/\u8C03\u6574\u57FA\u51C6\u91D1\u989D'),
                (U '\u53C2\u8003\u5206\u6210'),
                (U '\u63A8\u7406\u8C03\u6574:\u5206\u6210'),
                (U '\u5224\u65AD\u4F9D\u636E'),
                (U '\u6765\u6E90')
            )

            for ($col = 0; $col -lt $detailHeaders.Count; $col++) {
                $inferenceSheet.Cells.Item($detailHeaderRow, $col + 1).Value2 = $detailHeaders[$col]
            }

            $detailRowIndex = $detailHeaderRow + 1
            if ($null -ne $InferenceEntries -and $InferenceEntries.Count -gt 0) {
                foreach ($entry in $InferenceEntries) {
                    $inferenceSheet.Cells.Item($detailRowIndex, 1).Value2 = [string]$entry.Game
                    $inferenceSheet.Cells.Item($detailRowIndex, 2).Value2 = ("'" + [string]$entry.OrderId)
                    $inferenceSheet.Cells.Item($detailRowIndex, 3).Value2 = [double]$entry.BasisAmount
                    $inferenceSheet.Cells.Item($detailRowIndex, 4).Value2 = [double]$entry.ReferenceShare
                    $inferenceSheet.Cells.Item($detailRowIndex, 5).Value2 = [double]$entry.InferredAdjustment
                    $inferenceSheet.Cells.Item($detailRowIndex, 6).Value2 = [string]$entry.Reason
                    $inferenceSheet.Cells.Item($detailRowIndex, 7).Value2 = [string]$entry.Source
                    $detailRowIndex++
                }
            } else {
                $inferenceSheet.Cells.Item($detailRowIndex, 1).Value2 = U '\u672C\u6B21\u672A\u53D1\u73B0\u53EF\u81EA\u52A8\u63A8\u7406\u7684\u5206\u6210\u8C03\u6574\u60C5\u51B5\u3002'
                $detailRowIndex++
            }

            $summaryLastRow = [Math]::Max(3, $rowIndex - 1)
            $detailDataStartRow = $detailHeaderRow + 1
            $detailDataEndRow = [Math]::Max($detailDataStartRow, $detailRowIndex - 1)

            $inferenceSheet.Range("A1:I1").Font.Bold = $true
            $inferenceSheet.Range("A2:I2").Font.Bold = $true
            $inferenceSheet.Range(("A{0}:G{1}" -f $detailTitleRow, $detailHeaderRow)).Font.Bold = $true
            $inferenceSheet.Range(("B3:H{0}" -f $summaryLastRow)).NumberFormat = "0.00"
            if ($detailRowIndex -gt $detailDataStartRow -or ($null -ne $InferenceEntries -and $InferenceEntries.Count -gt 0)) {
                $inferenceSheet.Range(("B{0}:B{1}" -f $detailDataStartRow, $detailDataEndRow)).NumberFormat = "@"
                $inferenceSheet.Range(("C{0}:E{1}" -f $detailDataStartRow, $detailDataEndRow)).NumberFormat = "0.00"
            }
            $inferenceSheet.Columns.AutoFit() | Out-Null
        }

        if ($null -ne $Diagnostics -and $Diagnostics.Count -gt 0) {
            $diagnosticHeaders = @(
                (U '\u5F02\u5E38\u7C7B\u578B'),
                (U '\u72B6\u6001'),
                (U '\u6E38\u620F'),
                (U '\u8BA2\u5355\u53F7'),
                (U '\u91D1\u989D'),
                (U '\u6765\u6E90'),
                (U '\u8BF4\u660E')
            )

            $diagnosticSheet = $workbook.Worksheets.Add([System.Type]::Missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
            $diagnosticSheet.Name = U '\u5F02\u5E38\u68C0\u67E5'

            for ($col = 0; $col -lt $diagnosticHeaders.Count; $col++) {
                $diagnosticSheet.Cells.Item(1, $col + 1).Value2 = $diagnosticHeaders[$col]
            }

            $diagnosticSheet.Range("D:D").NumberFormat = "@"

            $diagnosticRowIndex = 2
            foreach ($diagnostic in $Diagnostics) {
                $diagnosticSheet.Cells.Item($diagnosticRowIndex, 1).Value2 = [string]$diagnostic.Type
                $diagnosticSheet.Cells.Item($diagnosticRowIndex, 2).Value2 = [string]$diagnostic.Status
                if (-not [string]::IsNullOrWhiteSpace([string]$diagnostic.Game)) {
                    $diagnosticSheet.Cells.Item($diagnosticRowIndex, 3).Value2 = [string]$diagnostic.Game
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$diagnostic.OrderId)) {
                    $diagnosticSheet.Cells.Item($diagnosticRowIndex, 4).Value2 = ("'" + [string]$diagnostic.OrderId)
                }
                if ($null -ne $diagnostic.Amount -and -not [string]::IsNullOrWhiteSpace([string]$diagnostic.Amount)) {
                    $diagnosticSheet.Cells.Item($diagnosticRowIndex, 5).Value2 = [double]$diagnostic.Amount
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$diagnostic.Source)) {
                    $diagnosticSheet.Cells.Item($diagnosticRowIndex, 6).Value2 = [string]$diagnostic.Source
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$diagnostic.Message)) {
                    $diagnosticSheet.Cells.Item($diagnosticRowIndex, 7).Value2 = [string]$diagnostic.Message
                }
                $diagnosticRowIndex++
            }

            $diagnosticSheet.Range("A1:G1").Font.Bold = $true
            $diagnosticSheet.Range("E:E").NumberFormat = "0.00"
            $diagnosticSheet.Columns.AutoFit() | Out-Null
        }

        Save-WorkbookSafely -Workbook $workbook -OutputPath $OutputPath
    }
    finally {
        $workbook.Close($false) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
}

$cleanupPaths = @()
$summary = @{}
$inferredSummary = $null
$inferredAdjustmentsByGame = @{}
$historicalInferenceProfiles = @{}
$diagnostics = New-Object 'System.Collections.Generic.List[object]'
$inferenceEntries = New-Object 'System.Collections.Generic.List[object]'
$orderToGame = @{}
$inputWorkbook = $null
$excel = $null
$refundOrders = @{}
$tradeOrders = @{}
$serviceFeeTransferAmountsByTime = @{}
$serviceFeeTransferAmountsByBucket = @{}
$inputCompatibleWorksheetCount = 0
$inputRecognizedWorksheetCount = 0
$inputRecognizedRowCount = 0

try {
    $resolvedInput = Resolve-UploadPath -Path $InputPath -AllowedExtensions @('.xls', '.xlsx', '.xlsm') -Label '原始支付宝账单'
    $resolvedShareStatement = Resolve-UploadPath -Path $ShareStatementPath -AllowedExtensions @('.xls', '.xlsx', '.xlsm') -Label '财务分成明细表'

    foreach ($resolvedFile in @($resolvedInput, $resolvedShareStatement)) {
        if ($null -ne $resolvedFile -and -not [string]::IsNullOrWhiteSpace($resolvedFile.Notice)) {
            Write-Output $resolvedFile.Notice
        }

        if ($null -ne $resolvedFile -and -not [string]::IsNullOrWhiteSpace($resolvedFile.CleanupPath)) {
            $cleanupPaths += $resolvedFile.CleanupPath
        }
    }

    $InputPath = $resolvedInput.ResolvedPath
    $ShareStatementPath = if ($null -ne $resolvedShareStatement) { $resolvedShareStatement.ResolvedPath } else { $null }

    try {
        $excel = New-Object -ComObject Excel.Application
    }
    catch {
        throw (Get-FriendlyExcelOperationMessage -Operation 'CreateExcelApplication' -Path $null -Exception $_.Exception)
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false

    Ensure-SummaryRow -Summary $summary -Game $OtherCategoryLabel
    Ensure-SummaryRow -Summary $summary -Game $TransferToWangshangLabel

    $inputWorkbook = Open-WorkbookSafely -ExcelApp $excel -Path $InputPath -Operation 'OpenInputWorkbook'

    foreach ($worksheet in $inputWorkbook.Worksheets) {
        $usedRange = $worksheet.UsedRange
        try {
            $values = $usedRange.Value2
            $rowCount = $usedRange.Rows.Count
            $columnCount = $usedRange.Columns.Count
            $startRow = if ($worksheet.Index -eq 1) { 4 } else { 1 }

            if ($columnCount -lt 21) {
                continue
            }

            $inputCompatibleWorksheetCount++

            for ($rowIndex = $startRow; $rowIndex -le $rowCount; $rowIndex++) {
                $productName = [string]$values[$rowIndex, 16]
                $baseOrderId = [string]$values[$rowIndex, 18]
                if ([string]::IsNullOrWhiteSpace($productName) -or [string]::IsNullOrWhiteSpace($baseOrderId)) {
                    continue
                }

                $productInfo = Get-ProductInfo -ProductName $productName
                if ($null -ne $productInfo) {
                    $orderToGame[$baseOrderId] = $productInfo.Game
                }
            }
        }
        finally {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null
        }
    }

    if ($inputCompatibleWorksheetCount -eq 0) {
        throw (U '\u539F\u59CB\u652F\u4ED8\u5B9D\u8D26\u5355\u683C\u5F0F\u4E0D\u6B63\u786E\uff0C\u672A\u627E\u5230\u7B26\u5408\u652F\u4ED8\u5B9D\u8D26\u52A1\u8868\u7ED3\u6784\u7684\u5DE5\u4F5C\u8868\u3002')
    }

    $shareStatement = Load-ShareStatementSummary -Path $ShareStatementPath -ExcelApp $excel
    $historicalInferenceProfiles = Load-HistoricalInferenceProfiles -RootPath $HistoricalOutputRoot -CurrentOutputPath $OutputPath -ExcelApp $excel
    $hasShareStatement = ($shareStatement.Games.Count -gt 0)
    if (-not [string]::IsNullOrWhiteSpace($ShareStatementPath)) {
        if ($shareStatement.MatchedWorksheetCount -eq 0) {
            throw (U '\u8D22\u52A1\u5206\u6210\u660E\u7EC6\u8868\u683C\u5F0F\u4E0D\u6B63\u786E\uff0C\u672A\u627E\u5230\u5305\u542B\u201C\u5546\u54C1\u540D\u79F0\u201D\u548C\u201C\u6263\u8D39\u91D1\u989D\u201D\u7684\u5DE5\u4F5C\u8868\u3002')
        }

        if ($shareStatement.MatchedRowCount -eq 0) {
            throw (U '\u8D22\u52A1\u5206\u6210\u660E\u7EC6\u8868\u4E2D\u672A\u627E\u5230\u53EF\u7528\u7684\u5206\u6210\u660E\u7EC6\u884C\uff0C\u8BF7\u68C0\u67E5\u4E0A\u4F20\u7684\u6587\u4EF6\u662F\u5426\u6B63\u786E\u3002')
        }
    }

    if ($hasShareStatement) {
        Write-Output ("分成将使用财务分成明细表：{0}" -f $ShareStatementPath)
    } else {
        Write-Output "未提供财务分成明细表，分成列将保持空白。"
    }
    if ($historicalInferenceProfiles.Count -gt 0) {
        Write-Output ("已加载历史推理参考：{0} 份" -f $historicalInferenceProfiles.Count)
    }

    foreach ($worksheet in $inputWorkbook.Worksheets) {
        $usedRange = $worksheet.UsedRange
        try {
            $values = $usedRange.Value2
            $rowCount = $usedRange.Rows.Count
            $columnCount = $usedRange.Columns.Count
            $startRow = if ($worksheet.Index -eq 1) { 4 } else { 1 }
            $unmatchedFeeLogCount = 0
            $unmatchedBaseFeeLogCount = 0
            $unmatchedFeeTotal = 0.0
            $unmatchedFeeCount = 0
            $unmatchedBaseFeeTotal = 0.0
            $unmatchedBaseFeeCount = 0
            $worksheetRecognizedRowCount = 0

            if ($columnCount -lt 21) {
                continue
            }

            for ($rowIndex = $startRow; $rowIndex -le $rowCount; $rowIndex++) {
                $entryTime = [string]$values[$rowIndex, 2]
                $businessDesc = [string]$values[$rowIndex, 21]
                $productName = [string]$values[$rowIndex, 16]
                $baseOrderId = [string]$values[$rowIndex, 18]
                $remark = [string]$values[$rowIndex, 17]
                $account = [string]$values[$rowIndex, 13]
                $remarkOrderId = Get-OrderIdFromRemark -Remark $remark
                $serviceFeeHourBucket = Get-ServiceFeeHourBucketFromRemark -Remark $remark
                $entryHourBucket = Get-TimeHourBucket -Text $entryTime
                if ([string]::IsNullOrWhiteSpace($baseOrderId) -and -not [string]::IsNullOrWhiteSpace($remarkOrderId)) {
                    $baseOrderId = $remarkOrderId
                }
                $isRemarkDerivedFee = Test-IsRemarkDerivedFee -BusinessDesc $businessDesc -Remark $remark -RemarkOrderId $remarkOrderId
                $isRemarkDerivedBaseFee = Test-IsRemarkDerivedBaseFee -BusinessDesc $businessDesc -Remark $remark

                $game = $null
                $productInfo = $null
                if (-not [string]::IsNullOrWhiteSpace($productName)) {
                    $productInfo = Get-ProductInfo -ProductName $productName
                    if ($null -ne $productInfo) {
                        $game = $productInfo.Game
                    }
                } elseif (-not [string]::IsNullOrWhiteSpace($baseOrderId) -and $orderToGame.ContainsKey($baseOrderId)) {
                    $game = $orderToGame[$baseOrderId]
                } elseif (-not [string]::IsNullOrWhiteSpace($baseOrderId) -and $shareStatement.OrderToGame.ContainsKey($baseOrderId)) {
                    $game = $shareStatement.OrderToGame[$baseOrderId]
                }

                $income = Get-NumericValue $values[$rowIndex, 7]
                $expense = Get-NumericValue $values[$rowIndex, 8]
                $hasAmount = ($income -ne 0.0 -or $expense -ne 0.0)

                if ($account -eq 'tbly@service.aliyun.com' -and $hasAmount) {
                    Add-BucketAmount -BucketAmounts $serviceFeeTransferAmountsByTime -Bucket $entryTime.Trim() -Amount ($expense - $income)
                    Add-BucketAmount -BucketAmounts $serviceFeeTransferAmountsByBucket -Bucket $serviceFeeHourBucket -Amount ($expense - $income)
                }

                if ((Test-IsRecognizedBusinessDescription -BusinessDesc $businessDesc) -or
                    $isRemarkDerivedFee -or
                    $isRemarkDerivedBaseFee -or
                    (-not [string]::IsNullOrWhiteSpace($remark) -and $remark.Contains($TransferToWangshangKeyword)) -or
                    ($account -eq 'tbly@service.aliyun.com' -and $hasAmount)) {
                    $worksheetRecognizedRowCount++
                    $inputRecognizedRowCount++
                }

                if (-not [string]::IsNullOrWhiteSpace($remark) -and $remark.Contains($TransferToWangshangKeyword)) {
                    $summary[$TransferToWangshangLabel].Income += $income
                    $summary[$TransferToWangshangLabel].Expense += $expense
                    continue
                }

                if (-not [string]::IsNullOrWhiteSpace($game)) {
                    Ensure-SummaryRow -Summary $summary -Game $game
                }

                $isTradeIncome = -not [string]::IsNullOrWhiteSpace($businessDesc) -and $businessDesc.StartsWith($TradeIncomePrefix, [System.StringComparison]::Ordinal)
                if ($isTradeIncome) {
                    if ([string]::IsNullOrWhiteSpace($game)) {
                        $summary[$OtherCategoryLabel].Income += $income
                        continue
                    }

                    $summary[$game].Income += $income
                    if (-not [string]::IsNullOrWhiteSpace($baseOrderId)) {
                        $tradeOrders[$baseOrderId] = [pscustomobject]@{
                            OrderId      = $baseOrderId
                            Game         = $game
                            Amount       = $income
                            ProductPrice = if ($null -ne $productInfo) { $productInfo.Price } else { $null }
                            EntryTime    = $entryTime.Trim()
                            HourBucket   = $entryHourBucket
                            Source       = ("{0}#{1}" -f $worksheet.Name, $rowIndex)
                        }
                    }
                    continue
                }

                $isRefund = -not [string]::IsNullOrWhiteSpace($businessDesc) -and $businessDesc.StartsWith($RefundPrefix, [System.StringComparison]::Ordinal)
                if ($isRefund) {
                    if ([string]::IsNullOrWhiteSpace($game)) {
                        $summary[$OtherCategoryLabel].Expense += $expense
                        continue
                    }

                    $summary[$game].Expense += $expense
                    if (-not [string]::IsNullOrWhiteSpace($baseOrderId)) {
                        $refundOrders[$baseOrderId] = [pscustomobject]@{
                            OrderId      = $baseOrderId
                            Game         = $game
                            Amount       = $expense
                            ProductPrice = if ($null -ne $productInfo) { $productInfo.Price } else { $null }
                            Source       = ("{0}#{1}" -f $worksheet.Name, $rowIndex)
                        }
                    }
                    continue
                }

                $isFee = (
                    (-not [string]::IsNullOrWhiteSpace($businessDesc) -and $businessDesc.StartsWith($FeePrefix, [System.StringComparison]::Ordinal)) -or
                    $isRemarkDerivedFee
                )
                if ($isFee) {
                    if ([string]::IsNullOrWhiteSpace($game)) {
                        $summary[$OtherCategoryLabel].Fee += ($expense - $income)
                        $unmatchedFeeCount++
                        $unmatchedFeeTotal += ($expense - $income)
                        if ($unmatchedFeeLogCount -lt 10) {
                            Write-Output ("Unmatched fee row in {0} row {1}: order={2}, amount={3}, remark={4}" -f $worksheet.Name, $rowIndex, $baseOrderId, ($expense - $income).ToString('0.00'), $remark)
                            $unmatchedFeeLogCount++
                        }
                        Add-Diagnostic -Diagnostics $diagnostics -Type (U '\u5B64\u7ACB\u624B\u7EED\u8D39') -Status (U '\u5F85\u590D\u6838') -Game $null -OrderId $baseOrderId -Amount ($expense - $income) -Source ("{0}#{1}" -f $worksheet.Name, $rowIndex) -Message $remark
                    } else {
                        $summary[$game].Fee += ($expense - $income)
                    }
                    continue
                }

                $isBaseFee = (
                    (-not [string]::IsNullOrWhiteSpace($businessDesc) -and $businessDesc.StartsWith($BaseFeePrefix, [System.StringComparison]::Ordinal)) -or
                    $isRemarkDerivedBaseFee
                )
                if ($isBaseFee) {
                    if ([string]::IsNullOrWhiteSpace($game)) {
                        $summary[$OtherCategoryLabel].BaseFee += $expense
                        $unmatchedBaseFeeCount++
                        $unmatchedBaseFeeTotal += $expense
                        if ($unmatchedBaseFeeLogCount -lt 10) {
                            Write-Output ("Unmatched base fee row in {0} row {1}: order={2}, amount={3}, remark={4}" -f $worksheet.Name, $rowIndex, $baseOrderId, $expense.ToString('0.00'), $remark)
                            $unmatchedBaseFeeLogCount++
                        }
                        Add-Diagnostic -Diagnostics $diagnostics -Type (U '\u5B64\u7ACB\u57FA\u7840\u8F6F\u4EF6\u8D39') -Status (U '\u5F85\u590D\u6838') -Game $null -OrderId $baseOrderId -Amount $expense -Source ("{0}#{1}" -f $worksheet.Name, $rowIndex) -Message $remark
                    } else {
                        $summary[$game].BaseFee += $expense
                    }
                    continue
                }

                if ($account -eq 'tbly@service.aliyun.com') {
                    continue
                }

                if ($hasAmount) {
                    $summary[$OtherCategoryLabel].Income += $income
                    $summary[$OtherCategoryLabel].Expense += $expense
                }
            }

            if ($unmatchedFeeCount -gt 0) {
                Write-Output ("Unmatched fee summary for {0}: rows={1}, total={2}" -f $worksheet.Name, $unmatchedFeeCount, $unmatchedFeeTotal.ToString('0.00'))
            }

            if ($unmatchedBaseFeeCount -gt 0) {
                Write-Output ("Unmatched base fee summary for {0}: rows={1}, total={2}" -f $worksheet.Name, $unmatchedBaseFeeCount, $unmatchedBaseFeeTotal.ToString('0.00'))
            }

            if ($worksheetRecognizedRowCount -gt 0) {
                $inputRecognizedWorksheetCount++
            }
        }
        finally {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null
        }
    }

    if ($inputRecognizedWorksheetCount -eq 0 -or $inputRecognizedRowCount -eq 0) {
        throw (U '\u539F\u59CB\u652F\u4ED8\u5B9D\u8D26\u5355\u683C\u5F0F\u4E0D\u6B63\u786E\uff0C\u672A\u627E\u5230\u53EF\u8BC6\u522B\u7684\u8D26\u52A1\u6D41\u6C34\u3002\u8BF7\u68C0\u67E5\u662F\u5426\u4E0A\u4F20\u4E86\u6B63\u786E\u7684\u8D26\u52A1\u8868\u3002')
    }

    if ($hasShareStatement) {
        $unmatchedShareCount = 0
        $unmatchedShareTotal = 0.0
        $unmatchedShareLogCount = 0

        foreach ($shareRow in $shareStatement.Rows) {
            $game = $null
            if (-not [string]::IsNullOrWhiteSpace($shareRow.OrderId) -and $orderToGame.ContainsKey($shareRow.OrderId)) {
                $game = $orderToGame[$shareRow.OrderId]
            }

            if ([string]::IsNullOrWhiteSpace($game)) {
                $game = $OtherCategoryLabel
                $unmatchedShareCount++
                $unmatchedShareTotal += $shareRow.Amount
                if ($unmatchedShareLogCount -lt 10) {
                    Write-Output ("Unmatched share row in {0} row {1}: order={2}, product={3}, amount={4}" -f $shareRow.Worksheet, $shareRow.RowIndex, $shareRow.OrderId, $shareRow.ProductName, $shareRow.Amount.ToString('0.00'))
                    $unmatchedShareLogCount++
                }
                Add-Diagnostic -Diagnostics $diagnostics -Type (U '\u5206\u6210\u672A\u5F52\u7C7B') -Status (U '\u5F85\u590D\u6838') -Game $null -OrderId $shareRow.OrderId -Amount $shareRow.Amount -Source ("{0}#{1}" -f $shareRow.Worksheet, $shareRow.RowIndex) -Message $shareRow.ProductName
            }

            Add-ShareAmount -Summary $summary -Game $game -Amount ([double]$shareRow.Amount)
        }

        if ($unmatchedShareCount -gt 0) {
            Write-Output ("Unmatched share summary: rows={0}, total={1}" -f $unmatchedShareCount, $unmatchedShareTotal.ToString('0.00'))
        }

        $inferredSummary = Copy-Summary -Summary $summary

        foreach ($refundOrder in $refundOrders.Values) {
            $hasNegativeShare = $shareStatement.NegativeAmountsByOrder.ContainsKey($refundOrder.OrderId)
            $positiveShare = if ($shareStatement.PositiveAmountsByOrder.ContainsKey($refundOrder.OrderId)) { [double]$shareStatement.PositiveAmountsByOrder[$refundOrder.OrderId] } else { 0.0 }
            $hasTradeInSamePeriod = $tradeOrders.ContainsKey($refundOrder.OrderId)
            $canInfer = $false
            $reason = $null
            $inferredShareBase = 0.0

            if ($hasNegativeShare) {
                continue
            }

            if ($hasTradeInSamePeriod -and $positiveShare -le 0.0) {
                Add-Diagnostic -Diagnostics $diagnostics -Type (U '\u540C\u8BA2\u5355\u6536\u5165\u9000\u6B3E\u5E76\u5B58') -Status (U '\u4EE5\u8D26\u52A1\u660E\u7EC6\u4E3A\u51C6') -Game $refundOrder.Game -OrderId $refundOrder.OrderId -Amount 0.0 -Source $refundOrder.Source -Message (U '\u8BE5\u8BA2\u5355\u5728\u672C\u671F\u539F\u59CB\u8D26\u52A1\u660E\u7EC6\u4E2D\u540C\u65F6\u51FA\u73B0\u4E86\u6536\u5165\u548C\u9000\u6B3E\uff0C\u4F46\u5206\u6210\u660E\u7EC6\u8868\u672A\u627E\u5230\u540C\u8BA2\u5355\u7684\u6B63\u5206\u6210\u8BB0\u5F55\u3002\u4E3A\u907F\u514D\u4E0E\u8D26\u52A1\u660E\u7EC6\u51B2\u7A81\uff0C\u672C\u6B21\u4E0D\u518D\u6309\u5386\u53F2\u6837\u672C\u81EA\u52A8\u63A8\u7406\u8D1F\u5206\u6210\uff0C\u4EE5\u672C\u671F\u539F\u59CB\u8D26\u52A1\u660E\u7EC6\u4E3A\u51C6\u4FDD\u7559\u539F\u7ED3\u679C')
                continue
            }

            if ($positiveShare -gt 0.0 -and $shareStatement.GamesByOrder.ContainsKey($refundOrder.OrderId)) {
                $gamesForOrder = $shareStatement.GamesByOrder[$refundOrder.OrderId]
                $sameGameOnly = ($gamesForOrder.Count -eq 1 -and $gamesForOrder.Contains($refundOrder.Game))
                $priceMatchesRefund = ($null -eq $refundOrder.ProductPrice -or [math]::Abs([double]$refundOrder.ProductPrice - [double]$refundOrder.Amount) -lt 0.01)

                if ($sameGameOnly -and $priceMatchesRefund) {
                    $canInfer = $true
                    $inferredShareBase = $positiveShare
                    $reason = ('{0}{1}{2}' -f (U '\u540C\u8BA2\u5355\u5728\u5206\u6210\u660E\u7EC6\u8868\u4E2D\u5DF2\u5B58\u5728\u6B63\u5206\u6210 '), $positiveShare.ToString('0.00'), (U '\uFF0C\u4F46\u7F3A\u5C11\u8D1F\u5206\u6210\u51B2\u56DE\uFF0C\u5DF2\u6309\u540C\u5355\u91D1\u989D\u81EA\u52A8\u8865\u56DE'))
                }
            }

            if (-not $canInfer -and $null -ne $refundOrder.ProductPrice) {
                $gamePriceKey = ('{0}|{1:F2}' -f $refundOrder.Game, [double]$refundOrder.ProductPrice)
                if ($shareStatement.PositiveAmountsByGamePrice.ContainsKey($gamePriceKey)) {
                    $priceProfile = $shareStatement.PositiveAmountsByGamePrice[$gamePriceKey]
                    $candidateAmounts = @($priceProfile.Amounts.Values)
                    if ($candidateAmounts.Count -eq 1) {
                        $canInfer = $true
                        $inferredShareBase = [double]$candidateAmounts[0]
                        $reason = ('{0}{1}{2}{3}{4}' -f (U '\u5206\u6210\u660E\u7EC6\u8868\u4E2D\u201C'), $refundOrder.Game, $refundOrder.ProductPrice.ToString('0.##'), (U '\u5143\u793C\u5305\u201D\u51FA\u73B0\u4E86\u7A33\u5B9A\u5206\u6210 '), $inferredShareBase.ToString('0.00'))
                        $reason += ('{0}{1}{2}' -f (U '\uFF08\u6837\u672C '), $priceProfile.Count, (U ' \u7B14\uFF09\uFF0C\u5DF2\u6309\u540C\u6E38\u620F\u540C\u91D1\u989D\u6863\u4F4D\u81EA\u52A8\u8865\u56DE'))
                    }
                }
            }

            if (-not $canInfer -and $null -ne $refundOrder.ProductPrice) {
                $gamePriceKey = ('{0}|{1:F2}' -f $refundOrder.Game, [double]$refundOrder.ProductPrice)
                if ($historicalInferenceProfiles.ContainsKey($gamePriceKey)) {
                    $historyProfile = $historicalInferenceProfiles[$gamePriceKey]
                    $candidateAmounts = @($historyProfile.Amounts.Values)
                    if ($candidateAmounts.Count -eq 1) {
                        $canInfer = $true
                        $inferredShareBase = [double]$candidateAmounts[0]
                        $reason = ('{0}{1}{2}{3}{4}' -f (U '\u5DF2\u53C2\u8003\u5386\u53F2\u8F93\u51FA\u7ED3\u679C\u4E2D\u201C'), $refundOrder.Game, $refundOrder.ProductPrice.ToString('0.##'), (U '\u5143\u793C\u5305\u201D\u7684\u7A33\u5B9A\u63A8\u7406\u5206\u6210 '), $inferredShareBase.ToString('0.00'))
                        $reason += ('{0}{1}{2}' -f (U '\uFF08\u5386\u53F2\u6837\u672C '), $historyProfile.Count, (U ' \u7B14\uFF09\uFF0C\u5DF2\u6309\u5386\u53F2\u53C2\u8003\u81EA\u52A8\u8865\u56DE'))
                    }
                }
            }

            if ($canInfer) {
                $inferredAdjustment = -1.0 * $inferredShareBase
                Add-ShareAmount -Summary $inferredSummary -Game $refundOrder.Game -Amount $inferredAdjustment

                if (-not $inferredAdjustmentsByGame.ContainsKey($refundOrder.Game)) {
                    $inferredAdjustmentsByGame[$refundOrder.Game] = [ordered]@{
                        Amount = 0.0
                        Count  = 0
                    }
                }

                $inferredAdjustmentsByGame[$refundOrder.Game].Amount += $inferredAdjustment
                $inferredAdjustmentsByGame[$refundOrder.Game].Count += 1

                $inferenceEntries.Add([pscustomobject]@{
                    Game               = $refundOrder.Game
                    OrderId            = $refundOrder.OrderId
                    BasisAmount        = [double]$refundOrder.Amount
                    ReferenceShare     = [double]$inferredShareBase
                    InferredAdjustment = [double]$inferredAdjustment
                    Reason             = $reason
                    Source             = $refundOrder.Source
                }) | Out-Null

                Add-Diagnostic -Diagnostics $diagnostics -Type (U '\u9000\u6B3E\u7F3A\u5C11\u8D1F\u5206\u6210\u51B2\u56DE') -Status (U '\u5DF2\u63A8\u7406') -Game $refundOrder.Game -OrderId $refundOrder.OrderId -Amount $inferredAdjustment -Source $refundOrder.Source -Message $reason
                Write-Output ("Inferred missing negative share reversal: order={0}, game={1}, inferredAdjustment={2}" -f $refundOrder.OrderId, $refundOrder.Game, $inferredAdjustment.ToString('0.00'))
                continue
            }

            Add-Diagnostic -Diagnostics $diagnostics -Type (U '\u9000\u6B3E\u7F3A\u5C11\u8D1F\u5206\u6210\u51B2\u56DE') -Status (U '\u5F85\u590D\u6838') -Game $refundOrder.Game -OrderId $refundOrder.OrderId -Amount $refundOrder.Amount -Source $refundOrder.Source -Message (U '\u8BE5\u9000\u6B3E\u5355\u5728\u5206\u6210\u660E\u7EC6\u8868\u4E2D\u672A\u627E\u5230\u540C\u8BA2\u5355\u7684\u8D1F\u5206\u6210\u51B2\u56DE\uff0c\u6682\u65E0\u6CD5\u81EA\u52A8\u63A8\u7406\uff0c\u8BF7\u5728\u5BFC\u51FA\u7ED3\u679C\u4E2D\u590D\u6838')
            Write-Output ("Refund is missing negative share reversal: order={0}, game={1}, amount={2}" -f $refundOrder.OrderId, $refundOrder.Game, $refundOrder.Amount.ToString('0.00'))
        }

        foreach ($tradeOrder in $tradeOrders.Values) {
            $hasPositiveShare = $shareStatement.PositiveAmountsByOrder.ContainsKey($tradeOrder.OrderId)
            $negativeShare = if ($shareStatement.NegativeAmountsByOrder.ContainsKey($tradeOrder.OrderId)) { [math]::Abs([double]$shareStatement.NegativeAmountsByOrder[$tradeOrder.OrderId]) } else { 0.0 }
            $hasRefundInSamePeriod = $refundOrders.ContainsKey($tradeOrder.OrderId)
            $canInfer = $false
            $reason = $null
            $inferredShareBase = 0.0
            $candidateAmounts = New-Object 'System.Collections.Generic.List[double]'

            if ($hasPositiveShare) {
                continue
            }

            if ($hasRefundInSamePeriod -and $negativeShare -gt 0.0 -and $shareStatement.GamesByOrder.ContainsKey($tradeOrder.OrderId)) {
                $gamesForOrder = $shareStatement.GamesByOrder[$tradeOrder.OrderId]
                if ($gamesForOrder.Count -eq 1 -and $gamesForOrder.Contains($tradeOrder.Game)) {
                    $canInfer = $true
                    $inferredShareBase = [double]$negativeShare
                    $reason = ('{0}{1}{2}' -f (U '\u540C\u8BA2\u5355\u5728\u5206\u6210\u660E\u7EC6\u8868\u4E2D\u5DF2\u5B58\u5728\u8D1F\u5206\u6210 '), $negativeShare.ToString('0.00'), (U '\uFF0C\u4F46\u7F3A\u5C11\u5BF9\u5E94\u7684\u6B63\u5206\u6210\u3002\u56E0\u539F\u59CB\u8D26\u52A1\u660E\u7EC6\u663E\u793A\u540C\u5355\u540C\u671F\u5148\u6536\u5165\u540E\u9000\u6B3E\uff0C\u5DF2\u6309\u540C\u8BA2\u5355\u8D1F\u5206\u6210\u53CD\u63A8\u8865\u5165\u6B63\u5206\u6210'))
                }
            }

            if ($hasRefundInSamePeriod -and -not $canInfer -and $negativeShare -le 0.0) {
                continue
            }

            if ($null -ne $tradeOrder.ProductPrice) {
                $gamePriceKey = ('{0}|{1:F2}' -f $tradeOrder.Game, [double]$tradeOrder.ProductPrice)
                if ($shareStatement.PositiveAmountsByGamePrice.ContainsKey($gamePriceKey)) {
                    $priceProfile = $shareStatement.PositiveAmountsByGamePrice[$gamePriceKey]
                    foreach ($value in $priceProfile.Amounts.Values) {
                        if (-not $candidateAmounts.Contains([double]$value)) {
                            [void]$candidateAmounts.Add([double]$value)
                        }
                    }
                    if ($candidateAmounts.Count -eq 1) {
                        $canInfer = $true
                        $inferredShareBase = [double]$candidateAmounts[0]
                        $reason = ('{0}{1}{2}{3}{4}' -f (U '\u5206\u6210\u660E\u7EC6\u8868\u4E2D\u201C'), $tradeOrder.Game, $tradeOrder.ProductPrice.ToString('0.##'), (U '\u5143\u793C\u5305\u201D\u51FA\u73B0\u4E86\u7A33\u5B9A\u5206\u6210 '), $inferredShareBase.ToString('0.00'))
                        $reason += ('{0}{1}{2}' -f (U '\uFF08\u6837\u672C '), $priceProfile.Count, (U ' \u7B14\uFF09\uFF0C\u672C\u7B14\u6B63\u5206\u6210\u7F3A\u5931\uff0c\u5DF2\u6309\u540C\u6E38\u620F\u540C\u91D1\u989D\u6863\u4F4D\u81EA\u52A8\u8865\u5165'))
                    }
                }
            }

            if (-not $canInfer -and $null -ne $tradeOrder.ProductPrice) {
                $gamePriceKey = ('{0}|{1:F2}' -f $tradeOrder.Game, [double]$tradeOrder.ProductPrice)
                if ($historicalInferenceProfiles.ContainsKey($gamePriceKey)) {
                    $historyProfile = $historicalInferenceProfiles[$gamePriceKey]
                    $historicalCandidateAmounts = @($historyProfile.Amounts.Values)
                    if ($candidateAmounts.Count -eq 0 -and $historicalCandidateAmounts.Count -eq 1) {
                        $canInfer = $true
                        $inferredShareBase = [double]$historicalCandidateAmounts[0]
                        $reason = ('{0}{1}{2}{3}{4}' -f (U '\u5DF2\u53C2\u8003\u5386\u53F2\u8F93\u51FA\u7ED3\u679C\u4E2D\u201C'), $tradeOrder.Game, $tradeOrder.ProductPrice.ToString('0.##'), (U '\u5143\u793C\u5305\u201D\u7684\u7A33\u5B9A\u63A8\u7406\u5206\u6210 '), $inferredShareBase.ToString('0.00'))
                        $reason += ('{0}{1}{2}' -f (U '\uFF08\u5386\u53F2\u6837\u672C '), $historyProfile.Count, (U ' \u7B14\uFF09\uFF0C\u672C\u7B14\u6B63\u5206\u6210\u7F3A\u5931\uff0c\u5DF2\u6309\u5386\u53F2\u53C2\u8003\u81EA\u52A8\u8865\u5165'))
                    }
                }
            }

            if (-not $canInfer -and $candidateAmounts.Count -gt 1 -and -not [string]::IsNullOrWhiteSpace($tradeOrder.HourBucket) -and $serviceFeeTransferAmountsByBucket.ContainsKey($tradeOrder.HourBucket)) {
                if (-not [string]::IsNullOrWhiteSpace($tradeOrder.EntryTime) -and $serviceFeeTransferAmountsByTime.ContainsKey($tradeOrder.EntryTime)) {
                    $matchedCandidates = New-Object 'System.Collections.Generic.List[double]'
                    foreach ($candidateAmount in $candidateAmounts) {
                        $candidateKey = ('{0:F2}' -f [double]$candidateAmount)
                        if ($serviceFeeTransferAmountsByTime[$tradeOrder.EntryTime].ContainsKey($candidateKey)) {
                            [void]$matchedCandidates.Add([double]$candidateAmount)
                        }
                    }

                    if ($matchedCandidates.Count -eq 1) {
                        $canInfer = $true
                        $inferredShareBase = [double]$matchedCandidates[0]
                        $reason = ('{0}{1}{2}' -f (U '\u540C\u6E38\u620F\u540C\u4EF7\u4F4D\u5B58\u5728\u591A\u79CD\u5206\u6210\u6863\u4F4D\uff0C\u4F46\u5728 '), $tradeOrder.EntryTime, (U ' \u8FD9\u4E2A\u51C6\u786E\u65F6\u70B9\u7684\u539F\u59CB\u8D26\u5355\u4E2D\u51FA\u73B0\u4E86\u552F\u4E00\u5339\u914D\u7684\u4E92\u52A8\u5F00\u653E\u865A\u62DF\u4E1A\u52A1\u8F6F\u4EF6\u670D\u52A1\u8D39\u91D1\u989D '))
                        $reason += $inferredShareBase.ToString('0.00')
                        $reason += (U '\uFF0C\u5DF2\u6309\u540C\u65F6\u70B9\u670D\u52A1\u8D39\u8BB0\u5F55\u81EA\u52A8\u8865\u5165\u8FD9\u7B14\u6B63\u5206\u6210')
                    }
                }
            }

            if (-not $canInfer -and $candidateAmounts.Count -gt 1 -and -not [string]::IsNullOrWhiteSpace($tradeOrder.HourBucket) -and $serviceFeeTransferAmountsByBucket.ContainsKey($tradeOrder.HourBucket)) {
                $matchedCandidates = New-Object 'System.Collections.Generic.List[double]'
                foreach ($candidateAmount in $candidateAmounts) {
                    $candidateKey = ('{0:F2}' -f [double]$candidateAmount)
                    if ($serviceFeeTransferAmountsByBucket[$tradeOrder.HourBucket].ContainsKey($candidateKey)) {
                        [void]$matchedCandidates.Add([double]$candidateAmount)
                    }
                }

                if ($matchedCandidates.Count -eq 1) {
                    $canInfer = $true
                    $inferredShareBase = [double]$matchedCandidates[0]
                    $reason = ('{0}{1}{2}' -f (U '\u540C\u6E38\u620F\u540C\u4EF7\u4F4D\u5B58\u5728\u591A\u79CD\u5206\u6210\u6863\u4F4D\uff0C\u4F46\u5728 '), $tradeOrder.HourBucket, (U ' \u8FD9\u4E2A\u5C0F\u65F6\u6863\u4F4D\u7684\u539F\u59CB\u8D26\u5355\u4E2D\u51FA\u73B0\u4E86\u552F\u4E00\u5339\u914D\u7684\u4E92\u52A8\u5F00\u653E\u865A\u62DF\u4E1A\u52A1\u8F6F\u4EF6\u670D\u52A1\u8D39\u91D1\u989D '))
                    $reason += $inferredShareBase.ToString('0.00')
                    $reason += (U '\uFF0C\u5DF2\u6309\u540C\u5C0F\u65F6\u670D\u52A1\u8D39\u8BB0\u5F55\u81EA\u52A8\u8865\u5165\u8FD9\u7B14\u6B63\u5206\u6210')
                }
            }

            if ($canInfer) {
                $inferredAdjustment = [double]$inferredShareBase
                Add-ShareAmount -Summary $inferredSummary -Game $tradeOrder.Game -Amount $inferredAdjustment

                if (-not $inferredAdjustmentsByGame.ContainsKey($tradeOrder.Game)) {
                    $inferredAdjustmentsByGame[$tradeOrder.Game] = [ordered]@{
                        Amount = 0.0
                        Count  = 0
                    }
                }

                $inferredAdjustmentsByGame[$tradeOrder.Game].Amount += $inferredAdjustment
                $inferredAdjustmentsByGame[$tradeOrder.Game].Count += 1

                $inferenceEntries.Add([pscustomobject]@{
                    Game               = $tradeOrder.Game
                    OrderId            = $tradeOrder.OrderId
                    BasisAmount        = [double]$tradeOrder.Amount
                    ReferenceShare     = [double]$inferredShareBase
                    InferredAdjustment = [double]$inferredAdjustment
                    Reason             = $reason
                    Source             = $tradeOrder.Source
                }) | Out-Null

                Add-Diagnostic -Diagnostics $diagnostics -Type (U '\u6536\u5165\u7F3A\u5C11\u6B63\u5206\u6210') -Status (U '\u5DF2\u63A8\u7406') -Game $tradeOrder.Game -OrderId $tradeOrder.OrderId -Amount $inferredAdjustment -Source $tradeOrder.Source -Message $reason
                Write-Output ("Inferred missing positive share: order={0}, game={1}, inferredAdjustment={2}" -f $tradeOrder.OrderId, $tradeOrder.Game, $inferredAdjustment.ToString('0.00'))
                continue
            }

            Add-Diagnostic -Diagnostics $diagnostics -Type (U '\u6536\u5165\u7F3A\u5C11\u6B63\u5206\u6210') -Status (U '\u5F85\u590D\u6838') -Game $tradeOrder.Game -OrderId $tradeOrder.OrderId -Amount $tradeOrder.Amount -Source $tradeOrder.Source -Message (U '\u8BE5\u6536\u5165\u5355\u5728\u5206\u6210\u660E\u7EC6\u8868\u4E2D\u672A\u627E\u5230\u540C\u8BA2\u5355\u7684\u6B63\u5206\u6210\uff0c\u6682\u65E0\u6CD5\u81EA\u52A8\u63A8\u7406\uff0c\u8BF7\u5728\u5BFC\u51FA\u7ED3\u679C\u4E2D\u590D\u6838')
            Write-Output ("Trade is missing positive share: order={0}, game={1}, amount={2}" -f $tradeOrder.OrderId, $tradeOrder.Game, $tradeOrder.Amount.ToString('0.00'))
        }
    }

    Save-SummaryWorkbook -Summary $summary -InferredSummary $inferredSummary -InferredAdjustmentsByGame $inferredAdjustmentsByGame -InferenceEntries $inferenceEntries -OutputPath $OutputPath -Diagnostics $diagnostics -ExcelApp $excel
}
finally {
    if ($inputWorkbook) {
        $inputWorkbook.Close($false) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($inputWorkbook) | Out-Null
    }

    if ($excel) {
        $excel.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    foreach ($cleanupPath in $cleanupPaths) {
        if (-not [string]::IsNullOrWhiteSpace($cleanupPath) -and (Test-Path -LiteralPath $cleanupPath)) {
            Remove-Item -LiteralPath $cleanupPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Write-Output ("已生成对账结果：{0}" -f $OutputPath)
