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
    # Every string is in both languages (docs/notes/localisation.md).
    $strings = Join-Path $PSScriptRoot 'src/SpreadsheetEditor.App/Strings'
    function Read-Strings([string]$language) {
        $table = @{}
        ([xml](Get-Content -Raw -Encoding utf8 (Join-Path $strings "$language/Resources.resw"))).root.data | ForEach-Object { $table[$_.name] = $_.value }
        $table
    }
    $english = Read-Strings 'en-US'
    $russian = Read-Strings 'ru-RU'
    $missing = @($english.Keys | Where-Object { -not $russian.ContainsKey($_) }) + @($russian.Keys | Where-Object { -not $english.ContainsKey($_) })
    if ($missing) { throw "strings: not in both Strings/en-US and Strings/ru-RU: $($missing -join ', ')" }

    $exe = Join-Path $out 'SpreadsheetEditor.exe'
    $log = Join-Path ([IO.Path]::GetTempPath()) 'SpreadsheetEditor-smoke.log'
    $texts = Join-Path ([IO.Path]::GetTempPath()) 'SpreadsheetEditor-smoke-texts.txt'
    # One run per interface language: the window must open, render, show its strings in that language and exit with 0.
    foreach ($run in @(@{ Language = 'en'; Strings = $english }, @{ Language = 'ru'; Strings = $russian })) {
        $started = Get-Date
        Remove-Item $log, $texts -ErrorAction SilentlyContinue
        $process = Start-Process $exe -ArgumentList '--smoke', '--smoke-language', $run.Language -PassThru
        if (-not $process.WaitForExit(60000)) {
            $process.Kill()
            throw "smoke run ($($run.Language)): the app did not exit within 60 s"
        }
        if ($process.ExitCode -ne 0) {
            if (Test-Path $log) { Get-Content $log | Write-Host }
            # A crash leaves its exception in the Application event log (".NET Runtime" and "Application Error" entries).
            Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = $started } -ErrorAction SilentlyContinue |
                Where-Object ProviderName -in '.NET Runtime', 'Application Error', 'Windows Error Reporting' |
                Format-List TimeCreated, ProviderName, Id, Message | Out-String -Width 300 | Write-Host
            throw "smoke run ($($run.Language)): exit code $($process.ExitCode) (0x$('{0:X8}' -f $process.ExitCode))"
        }
        # The texts the window showed, one "key<TAB>text" line per string (App.SmokeTextsPath), against the resource file.
        if (-not (Test-Path $texts)) { throw "smoke run ($($run.Language)): the app wrote no texts to $texts" }
        $shown = Get-Content -Encoding utf8 $texts
        if (-not $shown) { throw "smoke run ($($run.Language)): the app wrote no texts" }
        $wrong = foreach ($line in $shown) {
            $key, $text = $line -split "`t", 2
            if ($text -cne $run.Strings[$key]) { "$key shows '$text', expected '$($run.Strings[$key])'" }
        }
        if ($wrong) { throw "smoke run ($($run.Language)): $($wrong -join '; ')" }
        Write-Host "Smoke run passed ($($run.Language)): $($shown.Count) strings shown as in Strings/$(if ($run.Language -eq 'ru') { 'ru-RU' } else { 'en-US' })"
    }
}
