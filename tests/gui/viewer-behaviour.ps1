# End-to-end check that normal viewing still works with the security
# hardening, in the standalone viewer and in the Total Commander plugin:
# local images, .md links, #anchor links, the remote-images banner
# (Load images / Always / Never), the registration bar (#37: no dialog, Not now,
# Don't ask again, Esc), the gate before documents downloaded from the internet
# (#24: Open, remove the mark, Don't ask again, Close), Close on the
# remote-images bar, paths from documents (#49: no escape to other drives,
# shares, pipes; absolute paths; the network bar: Load files, Always load,
# Never; links to a share ask first), and the welcome screen.
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
# Mark of the Web (#24): the Zone.Identifier stream browsers write on download.
function Set-Zone($name, $content) {
    Set-Content -LiteralPath (Join-Path $work $name) -Stream Zone.Identifier -Value $content -NoNewline
}
foreach ($z in @(@('zone3.md', 'Zone 3', "[ZoneTransfer]`r`nZoneId=3`r`nHostUrl=https://example.com/zone3.md`r`n"),
                 @('zone4.md', 'Zone 4', "[ZoneTransfer]`r`nZoneId=4`r`n"),
                 @('zone2.md', 'Zone 2', "[ZoneTransfer]`r`nZoneId=2`r`n"),
                 @('zonebad.md', 'Zone garbage', "not a zone identifier"))) {
    Set-Content (Join-Path $work $z[0]) -Value "# $($z[1])`n`ntext`n`n![local image](local.jpg)" -NoNewline
    Set-Zone $z[0] $z[2]
}
Set-Content (Join-Path $work 'zonenone.md') -Value "# Zone none`n`ntext" -NoNewline
Set-Content (Join-Path $work 'zoneremote.md') -NoNewline -Value (
    "# Zone remote`n`n[Go away](zoneaway.md)`n`n![remote pixel](http://127.0.0.1:$port/pixel.png)")
Set-Zone 'zoneremote.md' "[ZoneTransfer]`r`nZoneId=3`r`n"
Set-Content (Join-Path $work 'zoneaway.md') -Value "# Zone away`n`n[Go back](zoneremote.md)" -NoNewline

# Paths from documents (#49). "Another local drive" is a temporary subst
# letter, the share is \\localhost\C$, a network drive is a temporary net use.
# Only drive letters this test defined are removed again.
$freeLetters = @('Q', 'R', 'W', 'Y', 'X', 'V', 'U', 'T', 'S', 'P', 'O', 'N') | Where-Object { -not (Test-Path "$($_):\") }
$substLetter = $freeLetters[0]; $netLetter = $freeLetters[1]
$otherDir = Join-Path $work 'other'
New-Item -ItemType Directory $otherDir | Out-Null
Copy-Item (Join-Path $work 'local.jpg') (Join-Path $otherDir 'big.jpg')
Copy-Item (Join-Path $work 'local.jpg') (Join-Path $work 'big.jpg')
subst "$($substLetter):" $otherDir
$substDefined = Test-Path "$($substLetter):\big.jpg"
if ($work -notmatch '^[Cc]:\\') { throw "the share checks need the test folder on C: ($work)" }
$shareWork = '\\localhost\C$\' + $work.Substring(3)
$shareUrl = 'file://localhost/C$/' + ($work.Substring(3) -replace '\\', '/')
$shareReachable = Test-Path (Join-Path $shareWork 'big.jpg')
$netOutput = net use "$($netLetter):" '\\localhost\C$' /persistent:no 2>&1
$netDefined = ($LASTEXITCODE -eq 0) -and (Test-Path "$($netLetter):\")
$up = '../' * 30
$pipeName = "mdview-behaviour-$PID"
$encodedShare = (($shareWork + '\big.jpg') -replace '\\', '%5C') -replace '\$', '%24'
Set-Content (Join-Path $work 'escape.md') -NoNewline -Value ("# Escape`n`n" +
    "![escape drive](${up}$($substLetter)%3A%5Cbig.jpg)`n`n" +
    "![escape share](${up}$encodedShare)`n`n" +
    "![escape pipe](${up}%5C%5C.%5Cpipe%5C$pipeName)")
Set-Content (Join-Path $work 'absolute.md') -NoNewline -Value ("# Absolute`n`n" +
    "![abs drive]($($substLetter):\big.jpg)`n`n![abs share](<$shareUrl/big.jpg>)")
Set-Content (Join-Path $work 'onshare.md') -Value "# On share`n`n![same share](big.jpg)" -NoNewline
Set-Content (Join-Path $work 'sharelink.md') -Value "# Share link`n`n[to the share](<$shareUrl/second.md>)" -NoNewline
Set-Content (Join-Path $work 'mapped.md') -NoNewline -Value (
    "# Mapped`n`n![abs mapped](<$($netLetter):\$($work.Substring(3))\big.jpg>)")
function Start-PipeProbe($name) {
    Start-ThreadJob -ArgumentList $name -ScriptBlock {
        param($n)
        $s = [System.IO.Pipes.NamedPipeServerStream]::new($n, 'InOut', 4, 'Byte', 'Asynchronous')
        try { if ($s.WaitForConnectionAsync().Wait(12000)) { 'CONNECTED' } else { 'no connection' } } finally { $s.Dispose() }
    }
}

# ---- registry
$key = 'HKCU:\Software\MDView'
$regNames = 'Left', 'Top', 'Right', 'Bottom', 'ShowCmd', 'DontAskRegister', 'ZoomPercent', 'RemoteImages', 'OpenDownloadedWithoutAsking', 'NetworkFiles', 'Zoom'
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
    Remove-ItemProperty -Path $key -Name OpenDownloadedWithoutAsking -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $key -Name NetworkFiles -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $key -Name Zoom -ErrorAction SilentlyContinue
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
function Get-IniValue($name) {
    if (-not (Test-Path $pluginIni)) { return $null }
    $inSection = $false
    foreach ($line in Get-Content $pluginIni) {
        if ($line -match '^\s*\[(.+)\]\s*$') { $inSection = ($Matches[1] -eq 'MDView'); continue }
        if ($inSection -and $line -match "^\s*$name\s*=\s*(\d+)\s*$") { return [int]$Matches[1] }
    }
    $null
}
function Get-IniRemoteImages { Get-IniValue 'RemoteImages' }
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

# Where each host keeps its network-files choice (registry or mdview.ini).
function Get-NetworkFilesSetting($hostKind) {
    if ($hostKind -eq 'exe') { (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).NetworkFiles } else { Get-IniValue 'NetworkFiles' }
}
function Get-OtherNetworkFilesSetting($hostKind) { Get-NetworkFilesSetting $(if ($hostKind -eq 'exe') { 'plugin' } else { 'exe' }) }
function Set-NetworkFilesSetting($hostKind, $value) {
    Set-RemoteImagesSetting $hostKind $null
    if ($null -eq $value) { return }
    if ($hostKind -eq 'exe') { Set-ItemProperty -Path $key -Name NetworkFiles -Value $value -Type DWord }
    else { Set-Content -Path $pluginIni -Value "[MDView]`r`nNetworkFiles=$value`r`n" -NoNewline }
}

# ---- viewer control
function Start-Viewer($hostKind, $doc, $iniFolder = $tcIni) {
    $savedTemp = $env:TEMP; $savedTmp = $env:TMP; $env:TEMP = $wv2Temp; $env:TMP = $wv2Temp
    try {
        $docArgs = if ($doc) { @('--gui', "`"$doc`"") } else { @('--gui') }
        if ($hostKind -eq 'exe') { $p = Start-Process -FilePath $Exe -ArgumentList $docArgs -PassThru }
        else {
            $p = Start-Process -FilePath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -PassThru -WindowStyle Hidden -ArgumentList @(
                '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$hostScript`"", '-Dll', "`"$Dll`"", '-Doc', "`"$doc`"",
                '-Seconds', 120, '-StatusFile', "`"$(Join-Path $work 'status.txt')`"", '-IniFolder', "`"$iniFolder`"")
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

Add-Type -Namespace Win32 -Name Behaviour -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
public delegate bool EnumProc(IntPtr h, IntPtr l);
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, System.Text.StringBuilder s, int n);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
[DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
[DllImport("user32.dll")] public static extern void mouse_event(uint flags, int dx, int dy, int data, UIntPtr extra);
'@
# Windows only lets a process take the foreground under some conditions; a
# synthetic F24 tap satisfies them and means nothing to MDView (see focus.ps1).
function Set-Foreground($hwnd) {
    if ([Win32.Behaviour]::GetForegroundWindow() -eq $hwnd) { return $true }
    [Win32.Behaviour]::keybd_event(0x87, 0, 0, [UIntPtr]::Zero)
    [Win32.Behaviour]::keybd_event(0x87, 0, 2, [UIntPtr]::Zero)
    [Win32.Behaviour]::SetForegroundWindow($hwnd) | Out-Null
    Start-Sleep -Milliseconds 800
    [Win32.Behaviour]::GetForegroundWindow() -eq $hwnd
}
function Send-Escape {
    [Win32.Behaviour]::keybd_event(0x1B, 0x01, 0, [UIntPtr]::Zero)
    [Win32.Behaviour]::keybd_event(0x1B, 0x01, 2, [UIntPtr]::Zero)
}
# A visible message box (dialog class #32770) of the viewer process, if one is
# open. Found with EnumWindows: UI Automation files an owned dialog under its
# owner window, not under the desktop.
function Get-Dialog($v) {
    $script:dialogFound = $null
    $targetPid = $v.Proc.Id
    $cb = [Win32.Behaviour+EnumProc]{
        param($h, $l)
        $procId = 0
        [void][Win32.Behaviour]::GetWindowThreadProcessId($h, [ref]$procId)
        if ($procId -eq $targetPid -and [Win32.Behaviour]::IsWindowVisible($h)) {
            $c = New-Object System.Text.StringBuilder 64
            [void][Win32.Behaviour]::GetClassName($h, $c, 64)
            if ($c.ToString() -eq '#32770') { $script:dialogFound = $AE::FromHandle($h); return $false }
        }
        $true
    }
    [void][Win32.Behaviour]::EnumWindows($cb, [IntPtr]::Zero)
    $script:dialogFound
}
# Answers "No" in the old registration dialog, so the remaining checks can run.
function Close-Dialog($v) {
    $d = Get-Dialog $v
    if (-not $d) { return }
    $no = $d.FindFirst($TS::Descendants, (New-Object System.Windows.Automation.AndCondition(
        (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Button)),
        (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, 'No')))))
    if ($no) { Invoke-Element $no; Start-Sleep -Milliseconds 500 }
}

$failures = 0
function Check($name, $ok, $detail) {
    $status = if ($ok) { 'PASS' } else { 'FAIL' }
    if (-not $ok) { $script:failures++ }
    Write-Output ("{0}  {1}  ({2})" -f $status, $name, $detail)
}
# A check that cannot run here is reported, never counted as passed.
function Skip($name, $reason) { Write-Output ("SKIP  {0}  ({1})" -f $name, $reason) }
function Get-ImageHeight($v, $name) {
    $e = Find-Element $v $CT::Image $name
    if ($e) { $e.Current.BoundingRectangle.Height } else { 0 }
}
function Wait-ImageHeight($v, $name, $seconds = 10) {
    $deadline = (Get-Date).AddSeconds($seconds)
    do { $h = Get-ImageHeight $v $name; if ($h -gt 50) { return $h }; Start-Sleep -Milliseconds 300 } while ((Get-Date) -lt $deadline)
    $h
}
function Wait-Dialog($v, $seconds = 8) {
    $deadline = (Get-Date).AddSeconds($seconds)
    do { $d = Get-Dialog $v; if ($d) { return $d }; Start-Sleep -Milliseconds 300 } while ((Get-Date) -lt $deadline)
    $null
}
function Get-DialogText($d) { ($d.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition) | ForEach-Object { $_.Current.Name }) -join ' | ' }
# A message box's buttons have no UIA Invoke pattern (they show as panes of
# class Button), and BM_CLICK only works while the box is in front: bring it
# to the front, click, and report whether the box closed.
function Press-DialogButton($d, $name) {
    $b = $d.FindFirst($TS::Descendants, (New-Object System.Windows.Automation.AndCondition(
        (New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'Button')),
        (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $name)))))
    if (-not $b) { return $false }
    $h = [IntPtr]$d.Current.NativeWindowHandle
    for ($try = 0; $try -lt 4; $try++) {
        if (Set-Foreground $h) {
            [void][Win32.Behaviour]::SendMessage([IntPtr]$b.Current.NativeWindowHandle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)
            Start-Sleep -Milliseconds 800
            try { if ($d.Current.IsOffscreen -or -not [Win32.Behaviour]::IsWindowVisible($h)) { return $true } } catch { return $true }
        }
    }
    $false
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

    # Registration question (#37): a bar in the document instead of a dialog.
    # The test build is never the registered exe, so only DontAskRegister decides.
    function Set-AskRegister { Set-TestRegistry $null; Remove-ItemProperty -Path $key -Name DontAskRegister -ErrorAction SilentlyContinue }
    function Get-DontAskRegister { (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).DontAskRegister }
    $registerText = '^Do you want to make MDView the default viewer for Markdown \(\.md\) files\?$'

    Set-AskRegister
    $v = Start-Viewer 'exe' (Join-Path $work 'main.md')
    [void](Wait-Element $v $CT::Text '^Main$')
    Start-Sleep -Seconds 3
    $dialog = Get-Dialog $v
    Check 'exe: no registration dialog' (-not $dialog) "dialog shown=$([bool]$dialog)"
    $fg = Set-Foreground $v.Proc.MainWindowHandle
    $focused = $AE::FocusedElement
    $focusDetail = "foreground=$fg dialog=$([bool]$dialog) focused=$($focused.Current.ControlType.ProgrammaticName) '$($focused.Current.Name)'"
    # The document itself has focus (it lives in WebView2's own process).
    Check 'exe: the registration question does not take keyboard focus' ($fg -and -not $dialog -and $focused.Current.ControlType -eq $CT::Document) $focusDetail
    Close-Dialog $v
    $bar = Wait-Element $v $CT::Text $registerText 5
    $found = foreach ($b in 'Register', 'Not now', "Don't ask again") { [bool](Find-Element $v $CT::Button "^$([regex]::Escape($b))$") }
    Check "exe: bar offers Register / Not now / Don't ask again" ([bool]$bar -and -not ($found -contains $false)) "bar=$([bool]$bar) buttons=$($found -join ',')"
    $notNow = Find-Element $v $CT::Button '^Not now$'
    if ($notNow) { Invoke-Element $notNow; Start-Sleep -Milliseconds 800 }
    $barAfter = Find-Element $v $CT::Text $registerText
    $link = Find-Element $v $CT::Hyperlink 'Go to second'
    if ($link) { Invoke-Element $link }
    [void](Wait-Element $v $CT::Text '^Second$' 10)
    Start-Sleep -Milliseconds 800
    $barOnSecond = Find-Element $v $CT::Text $registerText
    $dont = Get-DontAskRegister
    Stop-Viewer $v
    Check "exe: 'Not now' closes the bar and saves nothing" ([bool]$notNow -and -not $barAfter -and $null -eq $dont) "clicked=$([bool]$notNow) bar after=$([bool]$barAfter) DontAskRegister=$dont"
    Check "exe: after 'Not now' the next document does not ask again" ([bool]$notNow -and -not $barOnSecond) "clicked=$([bool]$notNow) bar on second=$([bool]$barOnSecond)"

    Set-AskRegister
    $v = Start-Viewer 'exe' (Join-Path $work 'main.md')
    [void](Wait-Element $v $CT::Text '^Main$')
    Start-Sleep -Seconds 2
    Close-Dialog $v
    $never = Wait-Element $v $CT::Button "^Don't ask again$" 5
    if ($never) { Invoke-Element $never; Start-Sleep -Seconds 1 }
    $dont = Get-DontAskRegister
    Stop-Viewer $v
    Check "exe: 'Don't ask again' saves DontAskRegister=1" ($dont -eq 1) "DontAskRegister=$dont"
    $v = Start-Viewer 'exe' (Join-Path $work 'main.md')
    [void](Wait-Element $v $CT::Text '^Main$')
    Start-Sleep -Seconds 3
    $dialog = Get-Dialog $v
    $bar = Find-Element $v $CT::Text $registerText
    Stop-Viewer $v
    Check "exe: after 'Don't ask again' no bar and no dialog" ($dont -eq 1 -and -not $dialog -and -not $bar) "DontAskRegister=$dont dialog=$([bool]$dialog) bar=$([bool]$bar)"

    # Esc closes the window as always; it must not mean "never ask again".
    Set-AskRegister
    $v = Start-Viewer 'exe' (Join-Path $work 'main.md')
    [void](Wait-Element $v $CT::Text '^Main$')
    Start-Sleep -Seconds 3
    $fg = Set-Foreground $v.Proc.MainWindowHandle
    Send-Escape
    $closed = $v.Proc.WaitForExit(5000)
    $dont = Get-DontAskRegister
    Stop-Viewer $v
    Check "exe: Esc closes the window and saves no 'never ask'" ($closed -and $null -eq $dont) "foreground=$fg closed=$closed DontAskRegister=$dont"

    Set-AskRegister
    $v = Start-Viewer 'exe' $null
    [void](Wait-Element $v $CT::Text 'Open a Markdown file using' 15)
    Start-Sleep -Seconds 3
    $dialog = Get-Dialog $v
    $bar = Find-Element $v $CT::Text $registerText
    Stop-Viewer $v
    Check 'exe: the welcome screen asks in the bar' ([bool]$bar -and -not $dialog) "bar=$([bool]$bar) dialog=$([bool]$dialog)"

    Set-AskRegister
    $v = Start-Viewer 'plugin' (Join-Path $work 'main.md')
    [void](Wait-Element $v $CT::Text '^Main$')
    Start-Sleep -Seconds 2
    $bar = Find-Element $v $CT::Text $registerText
    Stop-Viewer $v
    Check 'plugin: never asks to register' (-not $bar) "bar=$([bool]$bar)"

    # Mark of the Web (#24): a downloaded document is not shown until the
    # user chooses Open, Open and remove the mark, or Don't ask again.
    $gateText = '^This document was downloaded from the internet\.$'
    $gateButtons = 'Open', 'Open and remove the mark', "Don't ask again", 'Close'
    $remoteText = 'Remote images in this document were blocked'
    function Get-OpenWithoutAsking($hostKind) {
        if ($hostKind -eq 'exe') { (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).OpenDownloadedWithoutAsking } else { Get-IniValue 'OpenDownloadedWithoutAsking' }
    }
    function Get-OtherOpenWithoutAsking($hostKind) { Get-OpenWithoutAsking $(if ($hostKind -eq 'exe') { 'plugin' } else { 'exe' }) }
    function Test-Marked($name) { [bool](Get-Item -LiteralPath (Join-Path $work $name) -Stream Zone.Identifier -ErrorAction SilentlyContinue) }
    function Find-GateButton($v, $name) { Find-Element $v $CT::Button "^$([regex]::Escape($name))$" }

    foreach ($hostKind in 'exe', 'plugin') {
        foreach ($z in @(@('zone3.md', '^Zone 3$', 'internet (ZoneId=3)'), @('zone4.md', '^Zone 4$', 'restricted (ZoneId=4)'), @('zonebad.md', '^Zone garbage$', 'unreadable Zone.Identifier'))) {
            Set-RemoteImagesSetting $hostKind $null
            $v = Start-Viewer $hostKind (Join-Path $work $z[0])
            $gate = Wait-Element $v $CT::Text $gateText 15
            $found = foreach ($b in $gateButtons) { [bool](Find-GateButton $v $b) }
            $pathShown = [bool](Find-Element $v $CT::Text ([regex]::Escape((Join-Path $work $z[0]).Split('\')[-1]) + '$'))
            Start-Sleep -Seconds 1
            $heading = Find-Element $v $CT::Text $z[1]
            $img = Find-Element $v $CT::Image 'local image'
            Stop-Viewer $v
            Check "${hostKind}: $($z[2]): the gate is shown with Open / remove the mark / Don't ask again / Close" ([bool]$gate -and -not ($found -contains $false) -and $pathShown) "gate=$([bool]$gate) buttons=$($found -join ',') path shown=$pathShown"
            Check "${hostKind}: $($z[2]): nothing from the document is shown" ([bool]$gate -and -not $heading -and -not $img) "heading=$([bool]$heading) image=$([bool]$img)"
        }

        # Where it came from: as text, never as a link.
        $v = Start-Viewer $hostKind (Join-Path $work 'zone3.md')
        [void](Wait-Element $v $CT::Text $gateText 15)
        $from = Find-Element $v $CT::Text 'https://example\.com/zone3\.md'
        $fromLink = Find-Element $v $CT::Hyperlink 'example\.com'
        Stop-Viewer $v
        Check "${hostKind}: the gate shows where the file came from, as text" ([bool]$from -and -not $fromLink) "text=$([bool]$from) link=$([bool]$fromLink)"

        foreach ($z in @(@('zone2.md', '^Zone 2$', 'trusted sites (ZoneId=2)'), @('zonenone.md', '^Zone none$', 'no Zone.Identifier'))) {
            Set-RemoteImagesSetting $hostKind $null
            $v = Start-Viewer $hostKind (Join-Path $work $z[0])
            $heading = Wait-Element $v $CT::Text $z[1]
            $gate = Find-Element $v $CT::Text $gateText
            Stop-Viewer $v
            Check "${hostKind}: $($z[2]): shown directly" ([bool]$heading -and -not $gate) "document=$([bool]$heading) gate=$([bool]$gate)"
        }

        # Open: shown while the document is open, also after a reload ("Load
        # images"); another document and back asks again.
        Set-RemoteImagesSetting $hostKind $null
        $before = Get-PixelHits
        $v = Start-Viewer $hostKind (Join-Path $work 'zoneremote.md')
        [void](Wait-Element $v $CT::Text $gateText 15)
        $open = Find-GateButton $v 'Open'
        if ($open) { Invoke-Element $open }
        $shown = Wait-Element $v $CT::Text '^Zone remote$' 10
        $load = Wait-Element $v $CT::Button '^Load images$' 8
        if ($load) { Invoke-Element $load; Start-Sleep -Seconds 3 }
        $reloaded = (Get-PixelHits) -gt $before
        $afterReload = Find-Element $v $CT::Text '^Zone remote$'
        $gateAfterReload = Find-Element $v $CT::Text $gateText
        $away = Find-Element $v $CT::Hyperlink 'Go away'
        if ($away) { Invoke-Element $away }
        [void](Wait-Element $v $CT::Text '^Zone away$' 10)
        $back = Wait-Element $v $CT::Hyperlink 'Go back' 5
        if ($back) { Invoke-Element $back }
        $again = Wait-Element $v $CT::Text $gateText 10
        $stillMarked = Test-Marked 'zoneremote.md'
        Stop-Viewer $v
        Check "${hostKind}: Open shows the document" ([bool]$open -and [bool]$shown) "clicked=$([bool]$open) shown=$([bool]$shown)"
        Check "${hostKind}: after Open, a reload shows the document, not the gate" ($reloaded -and [bool]$afterReload -and -not $gateAfterReload) "reloaded=$reloaded document=$([bool]$afterReload) gate=$([bool]$gateAfterReload)"
        Check "${hostKind}: another document and back asks again; Open keeps the mark" ([bool]$again -and $stillMarked) "gate again=$([bool]$again) marked=$stillMarked"

        # [x] on the remote-images bar: closed while the document is open.
        Set-RemoteImagesSetting $hostKind $null
        Set-Content (Join-Path $work 'remote2.md') -NoNewline -Value "# Remote two`n`n[Go away](zoneaway.md)`n`n![remote pixel](http://127.0.0.1:$port/pixel.png)"
        Set-Content (Join-Path $work 'zoneaway.md') -NoNewline -Value "# Zone away`n`n[Go back](zoneremote.md)`n`n[Back to remote two](remote2.md)"
        $v = Start-Viewer $hostKind (Join-Path $work 'remote2.md')
        [void](Wait-Element $v $CT::Text '^Remote two$')
        $remoteBar = Wait-Element $v $CT::Text $remoteText 8
        $close = $null
        if ($remoteBar) {
            $bar = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($remoteBar)
            $close = $bar.FindFirst($TS::Descendants, (New-Object System.Windows.Automation.AndCondition(
                (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Button)),
                (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, 'Close')))))
        }
        if ($close) { Invoke-Element $close; Start-Sleep -Milliseconds 800 }
        $afterClose = Find-Element $v $CT::Text $remoteText
        $away = Find-Element $v $CT::Hyperlink 'Go away'
        if ($away) { Invoke-Element $away }
        [void](Wait-Element $v $CT::Text '^Zone away$' 10)
        $back = Wait-Element $v $CT::Hyperlink 'Back to remote two' 5
        if ($back) { Invoke-Element $back }
        [void](Wait-Element $v $CT::Text '^Remote two$' 10)
        $againBar = Wait-Element $v $CT::Text $remoteText 8
        $saved = Get-RemoteImagesSetting $hostKind
        Stop-Viewer $v
        Check "${hostKind}: Close hides the remote-images bar and saves nothing" ([bool]$close -and -not $afterClose -and $null -eq $saved) "close=$([bool]$close) bar after=$([bool]$afterClose) RemoteImages=$saved"
        Check "${hostKind}: the remote-images bar is back after another document and back" ([bool]$close -and [bool]$againBar) "bar again=$([bool]$againBar)"

        # Open and remove the mark: fails clearly while the mark is in use...
        Set-Content (Join-Path $work 'unblock.md') -Value "# Unblock me`n`ntext" -NoNewline
        Set-Zone 'unblock.md' "[ZoneTransfer]`r`nZoneId=3`r`n"
        $held = [IO.File]::Open((Join-Path $work 'unblock.md:Zone.Identifier'), 'Open', 'Read', 'None')
        try {
            $v = Start-Viewer $hostKind (Join-Path $work 'unblock.md')
            [void](Wait-Element $v $CT::Text $gateText 15)
            $unblock = Find-GateButton $v 'Open and remove the mark'
            if ($unblock) { Invoke-Element $unblock; Start-Sleep -Seconds 2 }
            $dialog = Get-Dialog $v
            $dialogText = if ($dialog) { ($dialog.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition) | ForEach-Object { $_.Current.Name }) -join ' | ' } else { '' }
            # Close it as the user would. (Invoking OK sends BM_CLICK, which
            # does nothing while the dialog is not the active window.)
            if ($dialog) {
                $dialog.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close()
                Start-Sleep -Milliseconds 800
            }
        } finally { $held.Close() }
        $gateStays = Find-Element $v $CT::Text $gateText
        $docShown = Find-Element $v $CT::Text '^Unblock me$'
        Check "${hostKind}: removing the mark fails with a clear message, and the gate stays" ([bool]$dialog -and $dialogText -match 'Could not remove the mark' -and $dialogText -match 'unblock\.md' -and [bool]$gateStays -and -not $docShown) "dialog=$([bool]$dialog) text='$dialogText' gate=$([bool]$gateStays) document=$([bool]$docShown)"
        # ...and works once it is no longer in use.
        $dialogOpen = [bool](Get-Dialog $v)
        $unblock = Find-GateButton $v 'Open and remove the mark'
        if ($unblock) { Invoke-Element $unblock }
        $docShown = Wait-Element $v $CT::Text '^Unblock me$' 10
        $marked = Test-Marked 'unblock.md'
        $dialog2 = Get-Dialog $v
        $dialog2Text = if ($dialog2) { ($dialog2.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition) | ForEach-Object { $_.Current.Name }) -join ' | ' } else { '' }
        Stop-Viewer $v
        Check "${hostKind}: 'Open and remove the mark' removes the stream and shows the document" ([bool]$docShown -and -not $marked) "document=$([bool]$docShown) still marked=$marked dialog before=$dialogOpen dialog after='$dialog2Text'"
        $v = Start-Viewer $hostKind (Join-Path $work 'unblock.md')
        $docShown = Wait-Element $v $CT::Text '^Unblock me$' 15
        $gate = Find-Element $v $CT::Text $gateText
        Stop-Viewer $v
        Check "${hostKind}: after removing the mark, no gate" ([bool]$docShown -and -not $gate) "document=$([bool]$docShown) gate=$([bool]$gate)"

        # Close: same as Esc; nothing is saved.
        Set-RemoteImagesSetting $hostKind $null
        $v = Start-Viewer $hostKind (Join-Path $work 'zone3.md')
        [void](Wait-Element $v $CT::Text $gateText 15)
        $close = Find-GateButton $v 'Close'
        if ($close) { Invoke-Element $close }
        $closed = $v.Proc.WaitForExit(8000)
        $saved = Get-OpenWithoutAsking $hostKind
        Stop-Viewer $v
        Check "${hostKind}: Close closes the window and saves nothing" ([bool]$close -and $closed -and $null -eq $saved -and (Test-Marked 'zone3.md')) "clicked=$([bool]$close) closed=$closed saved=$saved"

        # Don't ask again: saved in this program's own settings only.
        Set-RemoteImagesSetting $hostKind $null
        $v = Start-Viewer $hostKind (Join-Path $work 'zone3.md')
        [void](Wait-Element $v $CT::Text $gateText 15)
        $never = Find-GateButton $v "Don't ask again"
        if ($never) { Invoke-Element $never }
        $docShown = Wait-Element $v $CT::Text '^Zone 3$' 10
        $saved = Get-OpenWithoutAsking $hostKind
        $other = Get-OtherOpenWithoutAsking $hostKind
        Stop-Viewer $v
        Check "${hostKind}: 'Don't ask again' shows the document and saves OpenDownloadedWithoutAsking=1 in its own settings only" ([bool]$docShown -and $saved -eq 1 -and $null -eq $other) "document=$([bool]$docShown) saved=$saved other=$other"
        $v = Start-Viewer $hostKind (Join-Path $work 'zone4.md')
        $docShown = Wait-Element $v $CT::Text '^Zone 4$' 15
        $gate = Find-Element $v $CT::Text $gateText
        Stop-Viewer $v
        Check "${hostKind}: after 'Don't ask again' downloaded documents open directly" ([bool]$docShown -and -not $gate -and (Test-Marked 'zone4.md')) "document=$([bool]$docShown) gate=$([bool]$gate)"
        Set-Content (Join-Path $work 'zoneaway.md') -Value "# Zone away`n`n[Go back](zoneremote.md)" -NoNewline

        # ---- Paths from documents (#49)
        $netText = '^This document loads files from a network location:'
        $shareName = '^\\\\localhost\\C\$$'
        # Encoded separators, drive letters and pipes never reach anything,
        # even with network files allowed.
        Set-NetworkFilesSetting $hostKind 1
        $probe = Start-PipeProbe $pipeName
        $v = Start-Viewer $hostKind (Join-Path $work 'escape.md')
        $shown = Wait-Element $v $CT::Text '^Escape$' 15
        Start-Sleep -Seconds 3
        $hDrive = Get-ImageHeight $v 'escape drive'; $hShare = Get-ImageHeight $v 'escape share'
        $pipe = Receive-Job $probe -Wait; Remove-Job $probe
        Stop-Viewer $v
        Check "${hostKind}: encoded paths reach no other drive, share or pipe" ([bool]$shown -and $substDefined -and $shareReachable -and $hDrive -lt 50 -and $hShare -lt 50 -and $pipe -ne 'CONNECTED') "document=$([bool]$shown) drive=$hDrive share=$hShare pipe=$pipe (subst=$substDefined share reachable=$shareReachable)"

        # Absolute paths: another local drive is shown; a share waits for the bar.
        Set-NetworkFilesSetting $hostKind $null
        $v = Start-Viewer $hostKind (Join-Path $work 'absolute.md')
        [void](Wait-Element $v $CT::Text '^Absolute$' 15)
        $bar = Wait-Element $v $CT::Text $netText 10
        $location = Find-Element $v $CT::Text $shareName
        $hDrive = Wait-ImageHeight $v 'abs drive' 5
        $hShare = Get-ImageHeight $v 'abs share'
        Check "${hostKind}: an image with an absolute path on another local drive is shown" ($hDrive -gt 50) "height=$hDrive"
        Check "${hostKind}: an image on a share is blocked and the bar names the share" ($hShare -lt 50 -and [bool]$bar -and [bool]$location) "height=$hShare bar=$([bool]$bar) location shown=$([bool]$location)"
        $load = Find-Element $v $CT::Button '^Load files$'
        if ($load) { Invoke-Element $load }
        $hShare = Wait-ImageHeight $v 'abs share' 10
        $barAfter = Find-Element $v $CT::Text $netText
        Stop-Viewer $v
        $saved = Get-NetworkFilesSetting $hostKind
        Check "${hostKind}: 'Load files' shows the image from the share, bar gone, nothing saved" ([bool]$load -and $hShare -gt 50 -and -not $barAfter -and $null -eq $saved) "clicked=$([bool]$load) height=$hShare bar=$([bool]$barAfter) saved=$saved"
        $v = Start-Viewer $hostKind (Join-Path $work 'absolute.md')
        [void](Wait-Element $v $CT::Text '^Absolute$' 15)
        $bar = Wait-Element $v $CT::Text $netText 10
        $hShare = Get-ImageHeight $v 'abs share'
        Check "${hostKind}: opened again, the network bar asks again" ([bool]$bar -and $hShare -lt 50) "bar=$([bool]$bar) height=$hShare"
        $always = Find-Element $v $CT::Button '^Always load$'
        if ($always) { Invoke-Element $always }
        $hShare = Wait-ImageHeight $v 'abs share' 10
        $saved = Get-NetworkFilesSetting $hostKind; $other = Get-OtherNetworkFilesSetting $hostKind
        Stop-Viewer $v
        Check "${hostKind}: 'Always load' shows it and saves NetworkFiles=1 in its own settings only" ([bool]$always -and $hShare -gt 50 -and $saved -eq 1 -and $null -eq $other) "clicked=$([bool]$always) height=$hShare saved=$saved other=$other"
        $v = Start-Viewer $hostKind (Join-Path $work 'absolute.md')
        [void](Wait-Element $v $CT::Text '^Absolute$' 15)
        $hShare = Wait-ImageHeight $v 'abs share' 10
        $bar = Find-Element $v $CT::Text $netText
        Stop-Viewer $v
        Check "${hostKind}: after 'Always load' files from shares are shown without a bar" ($hShare -gt 50 -and -not $bar) "height=$hShare bar=$([bool]$bar)"
        Set-NetworkFilesSetting $hostKind $null
        $v = Start-Viewer $hostKind (Join-Path $work 'absolute.md')
        [void](Wait-Element $v $CT::Text $netText 15)
        $never = Find-Element $v $CT::Button '^Never$'
        if ($never) { Invoke-Element $never; Start-Sleep -Seconds 1 }
        $bar = Find-Element $v $CT::Text $netText
        $hShare = Get-ImageHeight $v 'abs share'
        $saved = Get-NetworkFilesSetting $hostKind
        Stop-Viewer $v
        Check "${hostKind}: 'Never' closes the bar, saves NetworkFiles=2, the image stays blocked" ([bool]$never -and -not $bar -and $saved -eq 2 -and $hShare -lt 50) "clicked=$([bool]$never) bar=$([bool]$bar) saved=$saved height=$hShare"

        # "Always load" for remote images does not cover network files.
        Set-RemoteImagesSetting $hostKind 1
        $v = Start-Viewer $hostKind (Join-Path $work 'absolute.md')
        [void](Wait-Element $v $CT::Text '^Absolute$' 15)
        $bar = Wait-Element $v $CT::Text $netText 10
        $hShare = Get-ImageHeight $v 'abs share'
        Stop-Viewer $v
        Check "${hostKind}: remote images 'Always' still asks for network files" ([bool]$bar -and $hShare -lt 50) "bar=$([bool]$bar) height=$hShare"

        # A document on a share shows its own images without asking.
        Set-NetworkFilesSetting $hostKind $null
        $v = Start-Viewer $hostKind (Join-Path $shareWork 'onshare.md')
        $shown = Wait-Element $v $CT::Text '^On share$' 15
        $h = Wait-ImageHeight $v 'same share' 10
        $bar = Find-Element $v $CT::Text $netText
        Stop-Viewer $v
        Check "${hostKind}: a document on a share shows images from that share without a bar" ([bool]$shown -and $h -gt 50 -and -not $bar) "document=$([bool]$shown) height=$h bar=$([bool]$bar)"

        # A link to a share asks first.
        $v = Start-Viewer $hostKind (Join-Path $work 'sharelink.md')
        $link = Wait-Element $v $CT::Hyperlink 'to the share' 15
        if ($link) { Invoke-Element $link }
        $d = Wait-Dialog $v
        $text = if ($d) { Get-DialogText $d } else { '' }
        $no = if ($d) { Press-DialogButton $d 'No' } else { $false }
        $stayed = Find-Element $v $CT::Text '^Share link$'
        if ($link) { Invoke-Element $link }
        $d = Wait-Dialog $v
        $yes = if ($d) { Press-DialogButton $d 'Yes' } else { $false }
        $second = Wait-Element $v $CT::Text '^Second$' 10
        Stop-Viewer $v
        Check "${hostKind}: a link to a share asks first; No stays, Yes opens it" ($text -match 'network location' -and $text -match 'localhost' -and $no -and [bool]$stayed -and $yes -and [bool]$second) "dialog='$text' no=$no stayed=$([bool]$stayed) yes=$yes opened=$([bool]$second)"

        # A network drive counts as a network location.
        if ($netDefined) {
            $v = Start-Viewer $hostKind (Join-Path $work 'mapped.md')
            [void](Wait-Element $v $CT::Text '^Mapped$' 15)
            $bar = Wait-Element $v $CT::Text $netText 10
            $location = Find-Element $v $CT::Text "^$($netLetter):$"
            $h = Get-ImageHeight $v 'abs mapped'
            Stop-Viewer $v
            Check "${hostKind}: an image on a network drive waits for the bar, which names the drive" ([bool]$bar -and [bool]$location -and $h -lt 50) "bar=$([bool]$bar) location=$([bool]$location) height=$h"
        } else {
            Skip "${hostKind}: an image on a network drive waits for the bar" "net use $($netLetter): \\localhost\C$ failed: $netOutput"
        }
    }

    # Nothing is read before the gate (#24): a downloaded document that another
    # program has locked still gets the gate (its Zone.Identifier stream can be
    # read); only after Open does MDView say clearly that it cannot be read.
    Set-RemoteImagesSetting 'exe' $null
    Set-Content (Join-Path $work 'locked.md') -Value "# Locked`n`ntext" -NoNewline
    Set-Zone 'locked.md' "[ZoneTransfer]`r`nZoneId=3`r`n"
    Set-Content (Join-Path $work 'locked-plain.md') -Value "# Locked plain`n`ntext" -NoNewline
    $lock = [IO.File]::Open((Join-Path $work 'locked.md'), 'Open', 'Read', 'None')
    $lockPlain = [IO.File]::Open((Join-Path $work 'locked-plain.md'), 'Open', 'Read', 'None')
    try {
        $v = $null
        try { $v = Start-Viewer 'exe' (Join-Path $work 'locked.md') } catch { }
        $gate = if ($v) { Wait-Element $v $CT::Text $gateText 15 } else { $null }
        $open = if ($gate) { Find-GateButton $v 'Open' } else { $null }
        if ($open) { Invoke-Element $open }
        $errPage = if ($open) { Wait-Element $v $CT::Text '^Could not read$' 10 } else { $null }
        $errorPath = if ($open) { Find-Element $v $CT::Text 'locked\.md$' } else { $null }
        if ($v) { Stop-Viewer $v }
        Check 'exe: a locked downloaded document still gets the gate (not read before)' ([bool]$gate) "window=$([bool]$v) gate=$([bool]$gate)"
        Check "exe: after Open, a document that can't be read says so" ([bool]$errPage -and [bool]$errorPath) "error page=$([bool]$errPage) path=$([bool]$errorPath)"

        $v = $null
        try { $v = Start-Viewer 'exe' (Join-Path $work 'locked-plain.md') } catch { }
        $errPage = if ($v) { Wait-Element $v $CT::Text '^Could not read$' 15 } else { $null }
        if ($v) { Stop-Viewer $v }
        Check "exe: a document that can't be read says so, no empty window" ([bool]$errPage) "window=$([bool]$v) error page=$([bool]$errPage)"

        # The plugin declines a file it cannot open, so TC shows its own message.
        Remove-Item (Join-Path $work 'status.txt') -ErrorAction SilentlyContinue
        $v = $null
        try { $v = Start-Viewer 'plugin' (Join-Path $work 'locked-plain.md') } catch { }
        Start-Sleep -Seconds 2
        $status = Get-Content (Join-Path $work 'status.txt') -ErrorAction SilentlyContinue
        if ($v) { Stop-Viewer $v }
        Check "plugin: a file it can't open is declined (TC shows its own message)" (-not $v -and ($status -contains 'load-failed')) "window=$([bool]$v) status=$($status -join ',')"
    } finally { $lock.Close(); $lockPlain.Close() }
    # After Open, the usual bars: registration on top, then remote images.
    Set-AskRegister
    $v = Start-Viewer 'exe' (Join-Path $work 'zoneremote.md')
    [void](Wait-Element $v $CT::Text $gateText 15)
    $gateRegister = Find-Element $v $CT::Text $registerText
    $open = Find-GateButton $v 'Open'
    if ($open) { Invoke-Element $open }
    [void](Wait-Element $v $CT::Text '^Zone remote$' 10)
    $remote = Wait-Element $v $CT::Text $remoteText 8
    $register = Find-Element $v $CT::Text $registerText
    $ys = foreach ($e in $register, $remote) { if ($e) { $e.Current.BoundingRectangle.Y } else { $null } }
    Stop-Viewer $v
    Check 'exe: the gate shows only the gate' (-not $gateRegister) "registration bar on the gate=$([bool]$gateRegister)"
    Check 'exe: after Open, bars are stacked Register, Remote images' ($null -notin $ys -and $ys[0] -lt $ys[1]) "Y=$($ys -join ', ')"
    # Welcome screen (standalone viewer without a file)
    Set-RemoteImagesSetting 'exe' $null
    $v = Start-Viewer 'exe' $null
    $welcome = Wait-Element $v $CT::Text 'Open a Markdown file using' 15
    Check 'exe: welcome screen is shown without a file' ([bool]$welcome) "welcome text shown=$([bool]$welcome)"
    Stop-Viewer $v

    # ---- Zoom (#16): Ctrl+Plus / Ctrl+wheel / Ctrl+0, remembered per program.
    # Measured as the height of the first heading, through UI Automation.
    function Get-ZoomSetting($hostKind) {
        if ($hostKind -eq 'exe') { (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).Zoom } else { Get-IniValue 'Zoom' }
    }
    function Get-HeadingHeight($v, $name) {
        $e = Find-Element $v $CT::Text "^$name$"
        if ($e) { $e.Current.BoundingRectangle.Height } else { 0 }
    }
    # Keys only reach the document when it has keyboard focus.
    function Focus-Document($v) {
        $hwnd = [IntPtr]$v.Win.Current.NativeWindowHandle
        if (-not (Set-Foreground $hwnd)) { return $false }
        $doc = Find-Element $v $CT::Document '.*'
        if (-not $doc) { return $false }
        $doc.SetFocus(); Start-Sleep -Milliseconds 300
        $true
    }
    function Send-CtrlKey($vk) {
        [Win32.Behaviour]::keybd_event(0x11, 0x1D, 0, [UIntPtr]::Zero)
        [Win32.Behaviour]::keybd_event($vk, 0, 0, [UIntPtr]::Zero)
        [Win32.Behaviour]::keybd_event($vk, 0, 2, [UIntPtr]::Zero)
        [Win32.Behaviour]::keybd_event(0x11, 0x1D, 2, [UIntPtr]::Zero)
    }
    # One notch of the mouse wheel (120 = up) with Ctrl held, over the document.
    function Send-CtrlWheel($v, $delta) {
        $r = $v.Win.Current.BoundingRectangle
        [void][Win32.Behaviour]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2))
        Start-Sleep -Milliseconds 200
        [Win32.Behaviour]::keybd_event(0x11, 0x1D, 0, [UIntPtr]::Zero)
        [Win32.Behaviour]::mouse_event(0x0800, 0, 0, $delta, [UIntPtr]::Zero)
        Start-Sleep -Milliseconds 100
        [Win32.Behaviour]::keybd_event(0x11, 0x1D, 2, [UIntPtr]::Zero)
    }
    # Waits until the heading is $ratio times $base high; returns the ratio seen.
    function Wait-Ratio($v, $name, $base, $ratio, $seconds = 5) {
        $deadline = (Get-Date).AddSeconds($seconds)
        do {
            $seen = (Get-HeadingHeight $v $name) / $base
            if ([Math]::Abs($seen - $ratio) -lt 0.05) { break }
            Start-Sleep -Milliseconds 250
        } while ((Get-Date) -lt $deadline)
        [Math]::Round($seen, 2)
    }
    # Heights are whole pixels, so a ratio is only close to the step.
    function Near($seen, $expected) { [Math]::Abs($seen - $expected) -lt 0.05 }
    $VK_OEM_PLUS = 0xBB; $VK_0 = 0x30

    foreach ($hostKind in 'exe', 'plugin') {
        $otherKind = if ($hostKind -eq 'exe') { 'plugin' } else { 'exe' }
        Set-RemoteImagesSetting $hostKind $null
        $v = Start-Viewer $hostKind (Join-Path $work 'main.md')
        [void](Wait-Element $v $CT::Text '^Main$')
        Start-Sleep -Milliseconds 500
        $base = Get-HeadingHeight $v 'Main'
        $focused = Focus-Document $v
        Send-CtrlKey $VK_OEM_PLUS
        $plus = Wait-Ratio $v 'Main' $base 1.1
        Start-Sleep -Milliseconds 300
        $saved = Get-ZoomSetting $hostKind; $other = Get-ZoomSetting $otherKind
        Check "${hostKind}: Ctrl+Plus zooms to 110% and saves Zoom=110 in its own settings" ($focused -and (Near $plus 1.1) -and $saved -eq 110 -and $null -eq $other) "focused=$focused ratio=$plus Zoom=$saved other=$other"

        Send-CtrlWheel $v 120
        $up = Wait-Ratio $v 'Main' $base 1.25
        Send-CtrlWheel $v -120
        $down = Wait-Ratio $v 'Main' $base 1.1
        Check "${hostKind}: one Ctrl+wheel notch is one step (110, 125, back to 110)" ((Near $up 1.25) -and (Near $down 1.1)) "up=$up down=$down Zoom=$(Get-ZoomSetting $hostKind)"

        $link = Wait-Element $v $CT::Hyperlink 'Go to second' 5
        if ($link) { Invoke-Element $link }
        [void](Wait-Element $v $CT::Text '^Second$' 10)
        $second = Wait-Ratio $v 'Second' $base 1.1
        Check "${hostKind}: the zoom stays after following a link" (Near $second 1.1) "ratio=$second"
        Stop-Viewer $v

        $v = Start-Viewer $hostKind (Join-Path $work 'main.md')
        [void](Wait-Element $v $CT::Text '^Main$')
        $restored = Wait-Ratio $v 'Main' $base 1.1
        Check "${hostKind}: the zoom is restored after a restart" (Near $restored 1.1) "ratio=$restored"

        $focused = Focus-Document $v
        Send-CtrlKey $VK_0
        $reset = Wait-Ratio $v 'Main' $base 1.0
        Start-Sleep -Milliseconds 300
        $saved = Get-ZoomSetting $hostKind
        Check "${hostKind}: Ctrl+0 resets to 100% and saves it" ($focused -and (Near $reset 1.0) -and $saved -eq 100) "focused=$focused ratio=$reset Zoom=$saved"
        Stop-Viewer $v
    }

    # The plugin with settings that can't be saved: one message, the zoom still works.
    Set-RemoteImagesSetting 'plugin' $null
    $badIni = Join-Path $work 'tc-ini-bad'
    New-Item -ItemType Directory -Force (Join-Path $badIni 'mdview.ini') | Out-Null
    $v = Start-Viewer 'plugin' (Join-Path $work 'main.md') $badIni
    [void](Wait-Element $v $CT::Text '^Main$')
    Start-Sleep -Milliseconds 500
    $base = Get-HeadingHeight $v 'Main'
    $focused = Focus-Document $v
    Send-CtrlKey $VK_OEM_PLUS
    $dialog = Wait-Dialog $v 8
    $dialogText = if ($dialog) { Get-DialogText $dialog } else { '' }
    $closedOk = if ($dialog) { Press-DialogButton $dialog 'OK' } else { $false }
    [void](Focus-Document $v)
    Send-CtrlKey $VK_OEM_PLUS
    $ratio = Wait-Ratio $v 'Main' $base 1.25
    $again = Wait-Dialog $v 3
    if ($again) { [void](Press-DialogButton $again 'OK') }
    Stop-Viewer $v
    Check "plugin: a zoom that can't be saved is reported once, and zooming still works" ($focused -and $dialogText -match 'Could not save the zoom level' -and $closedOk -and -not $again -and (Near $ratio 1.25)) "focused=$focused text='$dialogText' closed=$closedOk second dialog=$([bool]$again) ratio=$ratio"
}
finally {
    # Only windows this test started (tracked by process), never others.
    foreach ($p in $started) { if (-not $p.HasExited) { $p | Stop-Process -Force } }
    if (Get-Command Restore-Registry -ErrorAction SilentlyContinue) { Restore-Registry }
    if ($substDefined) { subst "$($substLetter):" /d }
    if ($netDefined) { net use "$($netLetter):" /delete /y | Out-Null }
    if ($tcp) { $tcp.Stop() }
    if ($job) { Stop-Job $job -ErrorAction SilentlyContinue; Remove-Job $job -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 2
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
exit $failures
