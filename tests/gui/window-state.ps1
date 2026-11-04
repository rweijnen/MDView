# End-to-end test for the standalone viewer's window state (issue #6).
# Launches mdview.exe --gui with a "show normal" hint (as Explorer and Total
# Commander do), drives the window with Win32 calls, and checks what Windows
# reports. Needs an interactive desktop. Backs up and restores
# HKCU\Software\MDView, and sets DontAskRegister while it runs so the
# registration prompt cannot block it.
#
# Usage:
#   .\tests\gui\window-state.ps1 -Exe .\target\debug\mdview.exe -Doc .\assets\manual-test\second.md
# Exit code = number of failed checks.
param(
    [Parameter(Mandatory)] [string]$Exe,
    [Parameter(Mandatory)] [string]$Doc
)

$ErrorActionPreference = 'Stop'

Add-Type -Namespace Win32 -Name User32 -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr w, IntPtr l);
'@

$SW_MAXIMIZE = 3; $SW_MINIMIZE = 6; $WM_CLOSE = 0x0010
$key = 'HKCU:\Software\MDView'
$names = 'Left', 'Top', 'Right', 'Bottom', 'ShowCmd', 'DontAskRegister'

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

function Start-Viewer {
    $proc = Start-Process -FilePath $Exe -ArgumentList "--gui", "`"$((Resolve-Path $Doc).Path)`"" -PassThru
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline) {
        $proc.Refresh()
        if ($proc.MainWindowHandle -ne [IntPtr]::Zero) {
            Start-Sleep -Milliseconds 1500   # let startup (placement, show) finish
            return $proc
        }
        Start-Sleep -Milliseconds 200
    }
    $proc | Stop-Process -Force
    throw "mdview.exe did not show a window within 30 seconds"
}

function Close-Viewer($proc) {
    [Win32.User32]::PostMessage($proc.MainWindowHandle, $WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
    if (-not $proc.WaitForExit(10000)) {
        $proc | Stop-Process -Force
        throw "mdview.exe did not exit after WM_CLOSE"
    }
}

$failures = 0
function Check($name, $ok, $detail) {
    $status = if ($ok) { 'PASS' } else { 'FAIL' }
    if (-not $ok) { $script:failures++ }
    Write-Output ("{0}  {1}  ({2})" -f $status, $name, $detail)
}

$backup = Get-Saved
try {
    $normal = @{ Left = 200; Top = 150; Right = 1200; Bottom = 850; DontAskRegister = 1 }

    # 1. Saving: maximise, close, the saved state must be "maximised" (3).
    Set-Saved ($normal + @{ ShowCmd = 1 })
    $p = Start-Viewer
    [Win32.User32]::ShowWindow($p.MainWindowHandle, $SW_MAXIMIZE) | Out-Null
    Start-Sleep -Milliseconds 500
    Close-Viewer $p
    $saved = (Get-Saved).ShowCmd
    Check 'closing maximised saves ShowCmd=3' ($saved -eq 3) "ShowCmd=$saved"

    # 2. Restoring: with "maximised" saved, the window must open maximised.
    Set-Saved ($normal + @{ ShowCmd = 3 })
    $p = Start-Viewer
    $zoomed = [Win32.User32]::IsZoomed($p.MainWindowHandle)
    Close-Viewer $p
    Check 'opens maximised when ShowCmd=3 is saved' $zoomed "IsZoomed=$zoomed"

    # 3. Normal state still restores as normal.
    Set-Saved ($normal + @{ ShowCmd = 1 })
    $p = Start-Viewer
    $zoomed = [Win32.User32]::IsZoomed($p.MainWindowHandle)
    $iconic = [Win32.User32]::IsIconic($p.MainWindowHandle)
    Close-Viewer $p
    Check 'opens normal when ShowCmd=1 is saved' ((-not $zoomed) -and (-not $iconic)) "IsZoomed=$zoomed IsIconic=$iconic"

    # 4. Closing while minimised must not make the next start minimised.
    Set-Saved ($normal + @{ ShowCmd = 1 })
    $p = Start-Viewer
    [Win32.User32]::ShowWindow($p.MainWindowHandle, $SW_MINIMIZE) | Out-Null
    Start-Sleep -Milliseconds 500
    Close-Viewer $p
    $saved = (Get-Saved).ShowCmd
    $p = Start-Viewer
    $iconic = [Win32.User32]::IsIconic($p.MainWindowHandle)
    Close-Viewer $p
    Check 'does not reopen minimised after closing minimised' (-not $iconic) "saved ShowCmd=$saved IsIconic=$iconic"

    # 5. Maximised, then minimised, then closed: reopen maximised, as Windows
    #    itself would restore it.
    Set-Saved ($normal + @{ ShowCmd = 1 })
    $p = Start-Viewer
    [Win32.User32]::ShowWindow($p.MainWindowHandle, $SW_MAXIMIZE) | Out-Null
    Start-Sleep -Milliseconds 500
    [Win32.User32]::ShowWindow($p.MainWindowHandle, $SW_MINIMIZE) | Out-Null
    Start-Sleep -Milliseconds 500
    Close-Viewer $p
    $saved = (Get-Saved).ShowCmd
    $p = Start-Viewer
    $zoomed = [Win32.User32]::IsZoomed($p.MainWindowHandle)
    $iconic = [Win32.User32]::IsIconic($p.MainWindowHandle)
    Close-Viewer $p
    Check 'reopens maximised after maximise -> minimise -> close' ($zoomed -and -not $iconic) "saved ShowCmd=$saved IsZoomed=$zoomed IsIconic=$iconic"

    # 6. A minimised state already saved by an older version must not start minimised.
    Set-Saved ($normal + @{ ShowCmd = 2 })
    $p = Start-Viewer
    $iconic = [Win32.User32]::IsIconic($p.MainWindowHandle)
    Close-Viewer $p
    Check 'does not start minimised when ShowCmd=2 was saved earlier' (-not $iconic) "IsIconic=$iconic"
}
finally {
    Set-Saved $backup
    Write-Output ("registry restored: " + (((Get-Saved).GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ' '))
}

exit $failures
