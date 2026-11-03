# Publishes the app self-contained into out/app and, with -Smoke, starts it once: it must open its window, render and exit with 0.
param(
    [ValidateSet('x64', 'ARM64')]
    [string]$Platform = 'x64',
    [switch]$Smoke
)
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true

$rid = if ($Platform -eq 'ARM64') { 'win-arm64' } else { 'win-x64' }
$out = Join-Path $PSScriptRoot 'out/app'
if (Test-Path $out) { Remove-Item -Recurse -Force $out }

Push-Location $PSScriptRoot
try {
    dotnet publish 'src/SpreadsheetEditor.App/SpreadsheetEditor.App.csproj' -c Release -r $rid "-p:Platform=$Platform" -o $out
}
finally {
    Pop-Location
}

# The app's PRI file holds its compiled XAML; without it the main window cannot load.
if (-not (Get-ChildItem $out -File | Where-Object Name -in 'SpreadsheetEditor.pri', 'resources.pri')) { throw "no app PRI file in $out" }
Write-Host "Published to $out"

if ($Smoke) {
    $exe = Join-Path $out 'SpreadsheetEditor.exe'
    $started = Get-Date
    $process = Start-Process $exe -ArgumentList '--smoke' -PassThru
    if (-not $process.WaitForExit(60000)) {
        $process.Kill()
        throw 'smoke run: the app did not exit within 60 s'
    }
    if ($process.ExitCode -ne 0) {
        # A crash leaves its exception in the Application event log (".NET Runtime" and "Application Error" entries).
        Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = $started } -ErrorAction SilentlyContinue |
            Where-Object ProviderName -in '.NET Runtime', 'Application Error', 'Windows Error Reporting' |
            Format-List TimeCreated, ProviderName, Id, Message | Out-String -Width 300 | Write-Host
        throw "smoke run: exit code $($process.ExitCode) (0x$('{0:X8}' -f $process.ExitCode))"
    }
    Write-Host 'Smoke run passed'
}
