# End-to-end test for #33: the standalone viewer and the Total Commander plugin
# must work at the same time. A WebView2 data folder can only be used by one
# host program at a time; before the fix both used %TEMP%\MDView_WebView2, and
# the one started second failed to load (0x8007139F).
#
# Runs mdview.exe and a TC-like plugin host (tests\security\lister-host.ps1)
# side by side, in both orders, with TEMP set to the same test folder. Each
# check passes only if the document is actually shown (read through UI
# Automation). Backs up and restores HKCU\Software\MDView.
#
# Usage:
#   .\tests\gui\data-folder.ps1 -Exe .\target\debug\mdview.exe -Dll .\target\debug\mdview_wlx.dll
# Exit code = number of failed checks.
param(
    [Parameter(Mandatory)] [string]$Exe,
    [Parameter(Mandatory)] [string]$Dll
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$Exe = (Resolve-Path $Exe).Path
$Dll = (Resolve-Path $Dll).Path
$hostScript = (Resolve-Path (Join-Path $PSScriptRoot '..\security\lister-host.ps1')).Path
$doc = (Resolve-Path (Join-Path $PSScriptRoot '..\..\assets\manual-test\second.md')).Path
$heading = '^Second Test File$'

$work = Join-Path $env:TEMP ('MDView-datafolder-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$temp = Join-Path $work 'temp'
New-Item -ItemType Directory $temp | Out-Null

$key = 'HKCU:\Software\MDView'
$regNames = 'Left', 'Top', 'Right', 'Bottom', 'ShowCmd', 'DontAskRegister'
$regBackup = @{}
$p0 = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
foreach ($n in $regNames) { if ($p0 -and $null -ne $p0.$n) { $regBackup[$n] = $p0.$n } }
function Restore-Registry {
    foreach ($n in $regNames) {
        if ($regBackup.ContainsKey($n)) { Set-ItemProperty -Path $key -Name $n -Value $regBackup[$n] -Type DWord }
        else { Remove-ItemProperty -Path $key -Name $n -ErrorAction SilentlyContinue }
    }
}

$started = New-Object System.Collections.ArrayList
function Start-WithTemp([scriptblock]$start) {
    $savedTemp = $env:TEMP; $savedTmp = $env:TMP; $env:TEMP = $temp; $env:TMP = $temp
    try { $p = & $start } finally { $env:TEMP = $savedTemp; $env:TMP = $savedTmp }
    [void]$started.Add($p)
    $p
}
function Start-Exe { Start-WithTemp { Start-Process -FilePath $Exe -ArgumentList '--gui', "`"$doc`"" -PassThru } }
function Start-PluginHost {
    Start-WithTemp {
        Start-Process -FilePath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -WindowStyle Hidden -PassThru -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$hostScript`"", '-Dll', "`"$Dll`"", '-Doc', "`"$doc`"",
            '-Seconds', 90, '-StatusFile', "`"$(Join-Path $work ('status-' + [guid]::NewGuid().ToString('N') + '.txt'))`"", '-Temp', "`"$temp`"")
    }
}

# True when the process shows the document (its heading is visible through UI Automation).
function Test-ShowsDocument($proc, $seconds = 25) {
    $deadline = (Get-Date).AddSeconds($seconds)
    # PropertyCondition needs a plain Int32, not a PowerShell-wrapped one.
    $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, [int]$proc.Id)
    while ((Get-Date) -lt $deadline -and -not $proc.HasExited) {
        foreach ($top in $AE::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)) {
            try {
                # Open the window by its handle: found through the desktop's list,
                # UI Automation does not show the WebView content.
                $w = $AE::FromHandle([IntPtr]$top.Current.NativeWindowHandle)
                foreach ($e in $w.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                        (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text)))) {
                    if ($e.Current.Name -match $heading) { return $true }
                }
            } catch { }
        }
        Start-Sleep -Milliseconds 400
    }
    $false
}
function Stop-All { foreach ($p in $started) { if (-not $p.HasExited) { $p | Stop-Process -Force; $p.WaitForExit(5000) | Out-Null } }; $started.Clear() }

$failures = 0
function Check($name, $ok) {
    $status = if ($ok) { 'PASS' } else { 'FAIL' }
    if (-not $ok) { $script:failures++ }
    Write-Output ("{0}  {1}" -f $status, $name)
}

try {
    foreach ($kv in @{ Left = 200; Top = 150; Right = 1100; Bottom = 850; ShowCmd = 1; DontAskRegister = 1 }.GetEnumerator()) {
        Set-ItemProperty -Path $key -Name $kv.Key -Value $kv.Value -Type DWord
    }

    # 1. Standalone viewer first, then the plugin (e.g. F3 in TC while a viewer window is open).
    $viewerProc = Start-Exe
    $exeShows = Test-ShowsDocument $viewerProc
    $pluginProc = Start-PluginHost
    $pluginShows = Test-ShowsDocument $pluginProc
    Check 'standalone viewer shows the document' $exeShows
    Check 'plugin started second also shows it' $pluginShows
    Stop-All

    # 2. Plugin first, then the standalone viewer (double-click while TC's Lister is open).
    $pluginProc = Start-PluginHost
    $pluginShows = Test-ShowsDocument $pluginProc
    $viewerProc = Start-Exe
    $exeShows = Test-ShowsDocument $viewerProc
    Check 'plugin shows the document' $pluginShows
    Check 'standalone viewer started second also shows it' $exeShows
    Stop-All
}
finally {
    Stop-All
    Restore-Registry
    Start-Sleep -Seconds 2
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
exit $failures
