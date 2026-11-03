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
    $log = Join-Path ([IO.Path]::GetTempPath()) 'SpreadsheetEditor-smoke.log'
    Remove-Item $log -ErrorAction SilentlyContinue
    $process = Start-Process $exe -ArgumentList '--smoke' -PassThru
    if (-not $process.WaitForExit(60000)) {
        $process.Kill()
        throw 'smoke run: the app did not exit within 60 s'
    }
    if ($process.ExitCode -ne 0) {
        if (Test-Path $log) { Get-Content $log | Write-Host }
        # Temporary (feat/app-localisation): which .resw files the PRI build indexed, and the app PRI's candidates for one string.
        $obj = Join-Path $PSScriptRoot 'src/SpreadsheetEditor.App/obj'
        Get-ChildItem $obj -Recurse -File -Include 'resources.resfiles', 'priconfig.xml', 'pri.resfiles' -ErrorAction SilentlyContinue |
            ForEach-Object { Write-Host "--- $($_.FullName)"; Get-Content $_.FullName | Write-Host }
        $packages = if ($env:NUGET_PACKAGES) { $env:NUGET_PACKAGES } else { Join-Path $HOME '.nuget/packages' }
        $makepri = Get-ChildItem (Join-Path $packages 'microsoft.windows.sdk.buildtools') -Recurse -Filter 'makepri.exe' -ErrorAction SilentlyContinue |
            Where-Object { $_.Directory.Name -eq 'x64' } | Select-Object -First 1
        try {
            if ($makepri) {
                $dump = Join-Path ([IO.Path]::GetTempPath()) 'SpreadsheetEditor-pri.xml'
                & $makepri.FullName dump /if (Join-Path $out 'SpreadsheetEditor.pri') /of $dump /dt detailed /o | Out-Null
                if (Test-Path $dump) {
                    Write-Host '--- makepri dump: MainWindow_EmptyTitle'
                    Select-String -Path $dump -Pattern 'MainWindow_EmptyTitle' -Context 0, 16 | ForEach-Object { $_.ToString() } | Write-Host
                    Write-Host '--- makepri dump: qualifiers'
                    Select-String -Path $dump -Pattern '<Qualifier ' | Select-Object -First 40 | ForEach-Object { $_.Line.Trim() } | Write-Host
                }
            }
            else { Write-Host 'makepri.exe not found in the NuGet cache' }
        }
        catch { Write-Host "makepri dump failed: $_" }
        # A crash leaves its exception in the Application event log (".NET Runtime" and "Application Error" entries).
        Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = $started } -ErrorAction SilentlyContinue |
            Where-Object ProviderName -in '.NET Runtime', 'Application Error', 'Windows Error Reporting' |
            Format-List TimeCreated, ProviderName, Id, Message | Out-String -Width 300 | Write-Host
        throw "smoke run: exit code $($process.ExitCode) (0x$('{0:X8}' -f $process.ExitCode))"
    }
    Write-Host 'Smoke run passed'
}
