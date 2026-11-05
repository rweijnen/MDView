# End-to-end check that normal viewing still works with the security
# hardening, in the standalone viewer and in the Total Commander plugin:
# local images, .md links, #anchor links, the remote-images banner
# (Load images / Always / Never) and the welcome screen.
#
# The "remote" images come from a local server on 127.0.0.1. Backs up and
# restores HKCU\Software\MDView; test windows get their own TEMP (own WebView2
# data folder) and are stopped afterwards. Needs an interactive desktop.
#
# Usage:
#   .\tests\gui\viewer-behaviour.ps1 -Exe .\target\debug\mdview.exe -Dll .\target\debug\mdview_wlx.dll
# Exit code = number of failed checks.
param(
    [Parameter(Mandatory)] [string]$Exe,
    [Parameter(Mandatory)] [string]$Dll
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]
$Exe = (Resolve-Path $Exe).Path
$Dll = (Resolve-Path $Dll).Path
$hostScript = Join-Path $PSScriptRoot '..\security\lister-host.ps1'

$work = Join-Path $env:TEMP ("MDView-behaviour-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory $work | Out-Null
$wv2Temp = Join-Path $work 'wv2-temp'
New-Item -ItemType Directory $wv2Temp | Out-Null

$tcp = $null; $job = $null
$started = New-Object System.Collections.ArrayList
try {
# ---- local "remote" image server
$tcp = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
$tcp.Start()
$port = $tcp.LocalEndpoint.Port
$hits = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$job = Start-ThreadJob -ArgumentList $tcp, $hits -ScriptBlock {
    param($tcp, $hits)
    $png = [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==')
    while ($true) {
        try { $c = $tcp.AcceptTcpClient() } catch { break }
        try {
            $s = $c.GetStream(); $r = [IO.StreamReader]::new($s, [Text.Encoding]::ASCII, $false, 4096, $true)
            $line = $r.ReadLine(); while (($h = $r.ReadLine()) -ne $null -and $h -ne '') { }
            if ($line) { $hits.Enqueue($line) }
            $head = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 200 OK`r`nContent-Type: image/png`r`nContent-Length: $($png.Length)`r`nConnection: close`r`n`r`n")
            $s.Write($head, 0, $head.Length); $s.Write($png, 0, $png.Length); $s.Flush()
        } catch { } finally { $c.Close() }
    }
}
function Get-PixelHits { @($hits.ToArray() | Where-Object { $_ -match '/pixel' }).Count }

# ---- documents
Copy-Item (Join-Path $PSScriptRoot '..\..\assets\mdview-poster.jpg') (Join-Path $work 'local.jpg')
Set-Content (Join-Path $work 'main.md') -NoNewline -Value (@(
    '# Main', '', '[Go to second](second.md)', '', '[Jump to bottom](#bottom)', '',
    '![local image](local.jpg)', '', (('filler line' + "`n`n") * 80), '## Bottom', '', 'end') -join "`n")
Set-Content (Join-Path $work 'second.md') -Value "# Second`n`nsecond document" -NoNewline
Set-Content (Join-Path $work 'remote.md') -Value "# Remote`n`n![remote pixel](http://127.0.0.1:$port/pixel.png)" -NoNewline
# Looks like MDView's bar, with the same classes and buttons, but it comes from the document.
Set-Content (Join-Path $work 'fake-bar.md') -NoNewline -Value ("# Fake bar`n`n" +
    "<div id=`"mdview-bars`"><div class=`"mdview-bar`" role=`"alert`" data-bar=`"remote-images`">" +
    "<span>Remote images in this document were blocked to protect your privacy.</span>" +
    "<button type=`"button`" onclick=`"chrome.webview.postMessage({type:'bar',bar:'remote-images',choice:'always'})`">Always load</button>" +
    "</div></div>`n`nText.")

# ---- registry
$key = 'HKCU:\Software\MDView'
$regNames = 'Left', 'Top', 'Right', 'Bottom', 'ShowCmd', 'DontAskRegister', 'ZoomPercent', 'RemoteImages'
$regBackup = @{}
$p0 = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
foreach ($n in $regNames) { if ($p0 -and $null -ne $p0.$n) { $regBackup[$n] = $p0.$n } }
function Restore-Registry {
    if (-not (Test-Path $key)) { New-Item -Path $key | Out-Null }
    foreach ($n in $regNames) {
        if ($regBackup.ContainsKey($n)) { Set-ItemProperty -Path $key -Name $n -Value $regBackup[$n] -Type DWord }
        else { Remove-ItemProperty -Path $key -Name $n -ErrorAction SilentlyContinue }
    }
}
function Set-TestRegistry($remoteImages) {
    Restore-Registry
    foreach ($kv in @{ Left = 200; Top = 150; Right = 1200; Bottom = 950; ShowCmd = 1; DontAskRegister = 1 }.GetEnumerator()) {
        Set-ItemProperty -Path $key -Name $kv.Key -Value $kv.Value -Type DWord
    }
    if ($null -ne $remoteImages) { Set-ItemProperty -Path $key -Name RemoteImages -Value $remoteImages -Type DWord }
}
function Get-RegistryRemoteImages { (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).RemoteImages }

# ---- the plugin's settings: mdview.ini in the folder TC reports (#34)
$tcIni = Join-Path $work 'tc-ini'
New-Item -ItemType Directory $tcIni | Out-Null
$pluginIni = Join-Path $tcIni 'mdview.ini'
function Get-IniRemoteImages {
    if (-not (Test-Path $pluginIni)) { return $null }
    $inSection = $false
    foreach ($line in Get-Content $pluginIni) {
        if ($line -match '^\s*\[(.+)\]\s*$') { $inSection = ($Matches[1] -eq 'MDView'); continue }
        if ($inSection -and $line -match '^\s*RemoteImages\s*=\s*(\d+)\s*$') { return [int]$Matches[1] }
    }
    $null
}
function Set-IniRemoteImages($value) {
    if ($null -eq $value) { Remove-Item $pluginIni -ErrorAction SilentlyContinue }
    else { Set-Content -Path $pluginIni -Value "[MDView]`r`nRemoteImages=$value`r`n" -NoNewline }
}
# Where each host keeps its remote-images choice: the standalone viewer in the
# registry, the plugin in mdview.ini.
function Get-RemoteImagesSetting($hostKind) { if ($hostKind -eq 'exe') { Get-RegistryRemoteImages } else { Get-IniRemoteImages } }
function Get-OtherRemoteImagesSetting($hostKind) { if ($hostKind -eq 'exe') { Get-IniRemoteImages } else { Get-RegistryRemoteImages } }
function Set-RemoteImagesSetting($hostKind, $value) {
    Set-TestRegistry $null
    Set-IniRemoteImages $null
    if ($hostKind -eq 'exe') { if ($null -ne $value) { Set-ItemProperty -Path $key -Name RemoteImages -Value $value -Type DWord } }
    else { Set-IniRemoteImages $value }
}

# ---- viewer control
function Start-Viewer($hostKind, $doc) {
    $savedTemp = $env:TEMP; $savedTmp = $env:TMP; $env:TEMP = $wv2Temp; $env:TMP = $wv2Temp
    try {
        $docArgs = if ($doc) { @('--gui', "`"$doc`"") } else { @('--gui') }
        if ($hostKind -eq 'exe') { $p = Start-Process -FilePath $Exe -ArgumentList $docArgs -PassThru }
        else {
            $p = Start-Process -FilePath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -PassThru -WindowStyle Hidden -ArgumentList @(
                '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$hostScript`"", '-Dll', "`"$Dll`"", '-Doc', "`"$doc`"",
                '-Seconds', 120, '-StatusFile', "`"$(Join-Path $work 'status.txt')`"", '-IniFolder', "`"$tcIni`"")
        }
    } finally { $env:TEMP = $savedTemp; $env:TMP = $savedTmp }
    [void]$started.Add($p)
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 300
        if ($hostKind -eq 'exe') { $p.Refresh(); if ($p.MainWindowHandle -ne 0) { return @{ Proc = $p; Win = $AE::FromHandle($p.MainWindowHandle) } } }
        else {
            $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $p.Id)
            $w = @($AE::RootElement.FindAll($TS::Children, $cond) | Where-Object { $_.Current.Name -like 'Lister (mdview)*' }) | Select-Object -First 1
            if ($w) { return @{ Proc = $p; Win = $w } }
        }
    }
    throw "$hostKind viewer window did not appear"
}
function Stop-Viewer($v) { if (-not $v.Proc.HasExited) { $v.Proc | Stop-Process -Force; $v.Proc.WaitForExit(5000) | Out-Null } }

function Find-Element($v, $type, $nameRegex) {
    foreach ($e in $v.Win.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $type)))) {
        if ($e.Current.Name -match $nameRegex) { return $e }
    }
    $null
}
function Wait-Element($v, $type, $nameRegex, $seconds = 15) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        $e = Find-Element $v $type $nameRegex
        if ($e) { return $e }
        Start-Sleep -Milliseconds 300
    }
    $null
}
function Invoke-Element($e) { $e.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke() }

$failures = 0
function Check($name, $ok, $detail) {
    $status = if ($ok) { 'PASS' } else { 'FAIL' }
    if (-not $ok) { $script:failures++ }
    Write-Output ("{0}  {1}  ({2})" -f $status, $name, $detail)
}

    foreach ($hostKind in 'exe', 'plugin') {
        # Local image, #anchor and .md link
        Set-TestRegistry $null
        $v = Start-Viewer $hostKind (Join-Path $work 'main.md')
        $heading = Wait-Element $v $CT::Text '^Main$'
        $img = Wait-Element $v $CT::Image 'local image' 5
        $imgH = if ($img) { $img.Current.BoundingRectangle.Height } else { 0 }
        Check "${hostKind}: local image is shown" ($imgH -gt 50) "image height=$imgH"
        $bottom = Find-Element $v $CT::Text '^Bottom$'
        $yBefore = if ($bottom) { $bottom.Current.BoundingRectangle.Y } else { $null }
        $anchor = Find-Element $v $CT::Hyperlink 'Jump to bottom'
        if ($anchor) { Invoke-Element $anchor; Start-Sleep -Milliseconds 1200 }
        $yAfter = if ($bottom) { $bottom.Current.BoundingRectangle.Y } else { $null }
        Check "${hostKind}: #anchor link scrolls to the heading" ($anchor -and $yBefore -ne $null -and $yAfter -lt $yBefore) "Y before=$yBefore after=$yAfter"
        $link = Find-Element $v $CT::Hyperlink 'Go to second'
        if ($link) { Invoke-Element $link }
        $second = Wait-Element $v $CT::Text '^Second$' 10
        Check "${hostKind}: .md link opens the other document" ([bool]$second) "second document shown=$([bool]$second)"
        Stop-Viewer $v

        # Remote images: blocked with a banner; "Load images" loads them once
        Set-RemoteImagesSetting $hostKind $null
        $before = Get-PixelHits
        $v = Start-Viewer $hostKind (Join-Path $work 'remote.md')
        [void](Wait-Element $v $CT::Text '^Remote$')
        $banner = Wait-Element $v $CT::Text 'Remote images in this document were blocked' 8
        Start-Sleep -Milliseconds 800
        Check "${hostKind}: remote image is not requested before asking" ((Get-PixelHits) -eq $before) "requests=$((Get-PixelHits) - $before)"
        Check "${hostKind}: banner offers to load remote images" ([bool]$banner) "banner shown=$([bool]$banner)"
        $load = Find-Element $v $CT::Button '^Load images$'
        if ($load) { Invoke-Element $load; Start-Sleep -Seconds 3 }
        Check "${hostKind}: 'Load images' loads the remote image" ((Get-PixelHits) -gt $before) "requests=$((Get-PixelHits) - $before)"
        $bannerAfter = Find-Element $v $CT::Text 'Remote images in this document were blocked'
        Check "${hostKind}: banner is gone after loading" (-not $bannerAfter) "banner shown=$([bool]$bannerAfter)"
        Check "${hostKind}: 'Load images' does not change the saved setting" ($null -eq (Get-RemoteImagesSetting $hostKind)) "RemoteImages=$(Get-RemoteImagesSetting $hostKind)"
        Stop-Viewer $v

        # "Always" is saved, and then images load without a banner
        Set-RemoteImagesSetting $hostKind $null
        $v = Start-Viewer $hostKind (Join-Path $work 'remote.md')
        [void](Wait-Element $v $CT::Text '^Remote$')
        $always = Wait-Element $v $CT::Button '^Always load$' 8
        if ($always) { Invoke-Element $always; Start-Sleep -Seconds 2 }
        $saved = Get-RemoteImagesSetting $hostKind
        $other = Get-OtherRemoteImagesSetting $hostKind
        Stop-Viewer $v
        Check "${hostKind}: 'Always load' saves RemoteImages=1 in its own settings" ($saved -eq 1) "RemoteImages=$saved"
        Check "${hostKind}: 'Always load' does not change the other program's settings" ($null -eq $other) "other=$other"
        Set-RemoteImagesSetting $hostKind 1
        $before = Get-PixelHits
        $v = Start-Viewer $hostKind (Join-Path $work 'remote.md')
        [void](Wait-Element $v $CT::Text '^Remote$')
        Start-Sleep -Seconds 2
        $banner = Find-Element $v $CT::Text 'Remote images in this document were blocked'
        Check "${hostKind}: with 'Always', remote images load without a banner" (((Get-PixelHits) -gt $before) -and -not $banner) "requests=$((Get-PixelHits) - $before) banner=$([bool]$banner)"
        Stop-Viewer $v

        # "Never" is saved, and then there is no banner and no request
        Set-RemoteImagesSetting $hostKind $null
        $v = Start-Viewer $hostKind (Join-Path $work 'remote.md')
        [void](Wait-Element $v $CT::Text '^Remote$')
        $never = Wait-Element $v $CT::Button '^Never$' 8
        if ($never) { Invoke-Element $never; Start-Sleep -Seconds 2 }
        $saved = Get-RemoteImagesSetting $hostKind
        $other = Get-OtherRemoteImagesSetting $hostKind
        Stop-Viewer $v
        Check "${hostKind}: 'Never' saves RemoteImages=2 in its own settings" ($saved -eq 2) "RemoteImages=$saved"
        Check "${hostKind}: 'Never' does not change the other program's settings" ($null -eq $other) "other=$other"
        Set-RemoteImagesSetting $hostKind 2
        $before = Get-PixelHits
        $v = Start-Viewer $hostKind (Join-Path $work 'remote.md')
        [void](Wait-Element $v $CT::Text '^Remote$')
        Start-Sleep -Seconds 2
        $banner = Find-Element $v $CT::Text 'Remote images in this document were blocked'
        Check "${hostKind}: with 'Never', no banner and no request" (((Get-PixelHits) -eq $before) -and -not $banner) "requests=$((Get-PixelHits) - $before) banner=$([bool]$banner)"
        Stop-Viewer $v

        # A document can't add a working bar (#40): its own HTML that imitates
        # the bar has no effect when clicked (no saved choice, no request).
        Set-RemoteImagesSetting $hostKind $null
        $before = Get-PixelHits
        $v = Start-Viewer $hostKind (Join-Path $work 'fake-bar.md')
        [void](Wait-Element $v $CT::Text '^Fake bar$')
        $fake = Wait-Element $v $CT::Button '^Always load$' 8
        if ($fake) { Invoke-Element $fake; Start-Sleep -Seconds 2 }
        $saved = Get-RemoteImagesSetting $hostKind
        $requests = (Get-PixelHits) - $before
        Stop-Viewer $v
        Check "${hostKind}: a bar imitated by the document does nothing" ([bool]$fake -and $null -eq $saved -and $requests -eq 0) "button found=$([bool]$fake) RemoteImages=$saved requests=$requests"
    }

    # Each program keeps its own answer (#34): "Always" in the registry (the
    # standalone viewer, or left by the v0.2.6 plugin) doesn't apply to the
    # plugin, and "Always" in the plugin's mdview.ini doesn't apply to the viewer.
    foreach ($case in @(@{ Store = 'exe'; Viewer = 'plugin' }, @{ Store = 'plugin'; Viewer = 'exe' })) {
        Set-RemoteImagesSetting $case.Store 1
        $before = Get-PixelHits
        $v = Start-Viewer $case.Viewer (Join-Path $work 'remote.md')
        [void](Wait-Element $v $CT::Text '^Remote$')
        $banner = Wait-Element $v $CT::Text 'Remote images in this document were blocked' 8
        Start-Sleep -Milliseconds 800
        $requests = (Get-PixelHits) - $before
        Stop-Viewer $v
        $from = if ($case.Store -eq 'exe') { 'the registry (standalone viewer / v0.2.6)' } else { "the plugin's mdview.ini" }
        Check "$($case.Viewer): 'Always' in $from does not apply here; it still asks" ([bool]$banner -and $requests -eq 0) "banner=$([bool]$banner) requests=$requests"
    }

    # Welcome screen (standalone viewer without a file)
    Set-RemoteImagesSetting 'exe' $null
    $v = Start-Viewer 'exe' $null
    $welcome = Wait-Element $v $CT::Text 'Open a Markdown file using' 15
    Check 'exe: welcome screen is shown without a file' ([bool]$welcome) "welcome text shown=$([bool]$welcome)"
    Stop-Viewer $v
}
finally {
    # Only windows this test started (tracked by process), never others.
    foreach ($p in $started) { if (-not $p.HasExited) { $p | Stop-Process -Force } }
    if (Get-Command Restore-Registry -ErrorAction SilentlyContinue) { Restore-Registry }
    if ($tcp) { $tcp.Stop() }
    if ($job) { Stop-Job $job -ErrorAction SilentlyContinue; Remove-Job $job -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 2
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
exit $failures
