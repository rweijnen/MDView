# End-to-end test: does the document get keyboard focus, so Ctrl+F, scrolling
# keys and Esc work without clicking first?
#
# Asks Windows which window has keyboard focus (GetGUIThreadInfo) instead of
# sending keystrokes: focus must be in WebView2's content window (class
# Chrome_*), not in MDView's own host window. Brings the MDView window to the
# foreground for a few seconds. Needs an interactive desktop. Backs up and
# restores HKCU\Software\MDView, and sets DontAskRegister while it runs.
#
# Usage:
#   .\tests\gui\focus.ps1 -Exe .\target\debug\mdview.exe -Doc .\assets\manual-test\second.md
# Exit code = number of failed checks.
param(
    [Parameter(Mandatory)] [string]$Exe,
    [Parameter(Mandatory)] [string]$Doc
)

$ErrorActionPreference = 'Stop'

Add-Type -Namespace Win32 -Name Focus -MemberDefinition @'
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
[StructLayout(LayoutKind.Sequential)] public struct GUITHREADINFO {
    public int cbSize; public int flags; public IntPtr hwndActive; public IntPtr hwndFocus;
    public IntPtr hwndCapture; public IntPtr hwndMenuOwner; public IntPtr hwndMoveSize;
    public IntPtr hwndCaret; public RECT rcCaret; }
[DllImport("user32.dll")] public static extern bool GetGUIThreadInfo(uint tid, ref GUITHREADINFO info);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern IntPtr GetShellWindow();
[DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, System.Text.StringBuilder s, int n);
[DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
'@

$key = 'HKCU:\Software\MDView'
$names = 'Left', 'Top', 'Right', 'Bottom', 'ShowCmd', 'DontAskRegister', 'ZoomPercent'

function Get-Saved {
    $p = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
    $h = @{}
    foreach ($n in $names) { if ($p -and $null -ne $p.$n) { $h[$n] = $p.$n } }
    $h
}

function Set-Saved($values) {
    if (-not (Test-Path $key)) { New-Item -Path $key | Out-Null }
    foreach ($n in $names) {
        if ($values.ContainsKey($n)) {
            Set-ItemProperty -Path $key -Name $n -Value $values[$n] -Type DWord
        } else {
            Remove-ItemProperty -Path $key -Name $n -ErrorAction SilentlyContinue
        }
    }
}

# Windows only lets a process take the foreground under some conditions;
# a synthetic key tap satisfies them. F24 is used because it means nothing to
# MDView; Alt would put MDView's menu bar into keyboard mode, which moves focus.
function Set-Foreground($hwnd) {
    if ([Win32.Focus]::GetForegroundWindow() -eq $hwnd) { return $true }
    [Win32.Focus]::keybd_event(0x87, 0, 0, [UIntPtr]::Zero)
    [Win32.Focus]::keybd_event(0x87, 0, 2, [UIntPtr]::Zero)
    [Win32.Focus]::SetForegroundWindow($hwnd) | Out-Null
    Start-Sleep -Milliseconds 800
    [Win32.Focus]::GetForegroundWindow() -eq $hwnd
}

function Get-FocusClass {
    $info = New-Object Win32.Focus+GUITHREADINFO
    $info.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf($info)
    if (-not [Win32.Focus]::GetGUIThreadInfo(0, [ref]$info)) { throw "GetGUIThreadInfo failed" }
    if ($info.hwndFocus -eq [IntPtr]::Zero) { return '<none>' }
    $sb = New-Object System.Text.StringBuilder 256
    [Win32.Focus]::GetClassName($info.hwndFocus, $sb, 256) | Out-Null
    $sb.ToString()
}

$failures = 0
function Check($name, $ok, $detail) {
    $status = if ($ok) { 'PASS' } else { 'FAIL' }
    if (-not $ok) { $script:failures++ }
    Write-Output ("{0}  {1}  ({2})" -f $status, $name, $detail)
}

$backup = Get-Saved
$proc = $null
try {
    Set-Saved (@{ Left = 200; Top = 150; Right = 1200; Bottom = 850; ShowCmd = 1; DontAskRegister = 1 })

    $proc = Start-Process -FilePath $Exe -ArgumentList "--gui", "`"$((Resolve-Path $Doc).Path)`"" -PassThru
    $deadline = (Get-Date).AddSeconds(30)
    while ($proc.MainWindowHandle -eq [IntPtr]::Zero) {
        if ((Get-Date) -gt $deadline) { throw "mdview.exe did not show a window within 30 seconds" }
        Start-Sleep -Milliseconds 200
        $proc.Refresh()
    }
    Start-Sleep -Milliseconds 1500
    $hwnd = $proc.MainWindowHandle

    # 1. Right after opening.
    if (-not (Set-Foreground $hwnd)) { throw "could not bring MDView to the foreground" }
    $cls = Get-FocusClass
    Check 'document has keyboard focus after opening' ($cls -like 'Chrome_*') "focus window class=$cls"

    # 2. After switching to another window and back.
    if (-not (Set-Foreground ([Win32.Focus]::GetShellWindow()))) { throw "could not switch away from MDView" }
    if (-not (Set-Foreground $hwnd)) { throw "could not bring MDView back to the foreground" }
    $cls = Get-FocusClass
    Check 'document has keyboard focus after switching back' ($cls -like 'Chrome_*') "focus window class=$cls"
}
finally {
    if ($proc -and -not $proc.HasExited) {
        [Win32.Focus]::PostMessage($proc.MainWindowHandle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
        if (-not $proc.WaitForExit(10000)) { $proc | Stop-Process -Force }
    }
    Set-Saved $backup
}

exit $failures
