Set-StrictMode -Version 3
$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourcePath = Join-Path -Path $scriptRoot -ChildPath "Start-ReconcileToolLauncher.cs"
$outputPath = Join-Path -Path $scriptRoot -ChildPath "Start-ReconcileTool.exe"
$iconBuilderPath = Join-Path -Path $scriptRoot -ChildPath "Generate-Start-ReconcileToolIcon.ps1"
$iconPath = Join-Path -Path $scriptRoot -ChildPath "Start-ReconcileTool.ico"
$cscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"

if (-not (Test-Path -LiteralPath $cscPath)) {
    throw "csc.exe not found: $cscPath"
}

if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "Launcher source not found: $sourcePath"
}

if (-not (Test-Path -LiteralPath $iconBuilderPath)) {
    throw "Icon builder not found: $iconBuilderPath"
}

& $iconBuilderPath -IcoPath $iconPath | Out-Host

& $cscPath `
    /nologo `
    /target:winexe `
    /platform:anycpu `
    /out:$outputPath `
    /win32icon:$iconPath `
    /reference:System.dll `
    /reference:System.Windows.Forms.dll `
    $sourcePath

Write-Output ("Built launcher: {0}" -f $outputPath)
