# Minimal Total Commander Lister host for tests: loads mdview_wlx.dll, shows a
# parent window like TC's Lister, calls ListLoadW and pumps messages.
#
# Must run in Windows PowerShell 5.1 (powershell.exe, x64): WebView2 needs an
# STA thread, which 5.1 uses by default. Writes a status line to -StatusFile:
#   loaded | load-failed | closed-by-plugin | timeout
#
# Like TC, it calls ListSetDefaultParams right after loading the plugin (if the
# plugin exports it), with DefaultIniName = <IniFolder>\lsplugin.ini. TC puts
# that file next to wincmd.ini.
#
# Usage:
#   powershell.exe -NoProfile -File lister-host.ps1 -Dll mdview_wlx.dll -Doc x.md -Seconds 6 -StatusFile s.txt [-Search text] [-IniFolder dir] [-CommandFile c.txt -LogFile l.txt]
#
# With -CommandFile and -LogFile it acts like TC on request (#59): each new line of
# the command file is run once, and the result appended to the log:
#   search <flags> <text>  ListSearchTextW (lcs_* flags)   -> result <n> <return value>
#   send <command>         ListSendCommand (lc_* command) -> result <n> <return value>
#   print                  ListPrintW                     -> result <n> <return value | missing>
# Keys the plugin passes on to this window (as to TC's Lister) are logged as
#   key <virtual key code>
param(
    [Parameter(Mandatory)] [string]$Dll,
    [Parameter(Mandatory)] [string]$Doc,
    [int]$Seconds = 6,
    [Parameter(Mandatory)] [string]$StatusFile,
    [string]$Search,
    # Use this folder as TEMP (an elevated host does not inherit the caller's environment)
    [string]$Temp,
    # Folder TC would report as its ini folder (where wincmd.ini is); default: next to the status file
    [string]$IniFolder,
    [string]$CommandFile,
    [string]$LogFile
)

$ErrorActionPreference = 'Stop'
if ($Temp) { $env:TEMP = $Temp; $env:TMP = $Temp }
Add-Type -AssemblyName System.Windows.Forms

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class ListerHost
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern IntPtr LoadLibraryW(string path);
    [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
    static extern IntPtr GetProcAddress(IntPtr module, string name);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr window, int cmd);

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    delegate IntPtr ListLoadWFn(IntPtr parent, [MarshalAs(UnmanagedType.LPWStr)] string file, int flags);
    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    delegate void ListCloseWindowFn(IntPtr window);
    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    delegate int ListSearchTextWFn(IntPtr window, [MarshalAs(UnmanagedType.LPWStr)] string text, int flags);
    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    delegate int ListSendCommandFn(IntPtr window, int command, int parameter);
    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    delegate int ListPrintWFn(IntPtr window, [MarshalAs(UnmanagedType.LPWStr)] string file, [MarshalAs(UnmanagedType.LPWStr)] string printer, int flags, IntPtr margins);

    // WLX SDK: ListDefaultParamStruct (DefaultIniName is ANSI, MAX_PATH)
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    public struct ListDefaultParamStruct
    {
        public int size;
        public uint PluginInterfaceVersionLow;
        public uint PluginInterfaceVersionHi;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string DefaultIniName;
    }
    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    delegate void ListSetDefaultParamsFn(ref ListDefaultParamStruct dps);

    static ListLoadWFn load;
    static ListCloseWindowFn close;
    static ListSearchTextWFn search;
    static ListSetDefaultParamsFn setDefaultParams;
    static ListSendCommandFn sendCommand;
    static ListPrintWFn print;

    public static void Init(string dll)
    {
        IntPtr module = LoadLibraryW(dll);
        if (module == IntPtr.Zero)
            throw new Exception("LoadLibrary failed for " + dll + ", error " + Marshal.GetLastWin32Error());
        load = Get<ListLoadWFn>(module, "ListLoadW");
        close = Get<ListCloseWindowFn>(module, "ListCloseWindow");
        search = Get<ListSearchTextWFn>(module, "ListSearchTextW");
        // Optional export, as in TC: older plugins don't have it.
        IntPtr p = GetProcAddress(module, "ListSetDefaultParams");
        if (p != IntPtr.Zero)
            setDefaultParams = Marshal.GetDelegateForFunctionPointer(p, typeof(ListSetDefaultParamsFn)) as ListSetDefaultParamsFn;
        p = GetProcAddress(module, "ListSendCommand");
        if (p != IntPtr.Zero) sendCommand = Marshal.GetDelegateForFunctionPointer(p, typeof(ListSendCommandFn)) as ListSendCommandFn;
        p = GetProcAddress(module, "ListPrintW");
        if (p != IntPtr.Zero) print = Marshal.GetDelegateForFunctionPointer(p, typeof(ListPrintWFn)) as ListPrintWFn;
    }

    static T Get<T>(IntPtr module, string name) where T : class
    {
        IntPtr p = GetProcAddress(module, name);
        if (p == IntPtr.Zero) throw new Exception("export not found: " + name);
        return Marshal.GetDelegateForFunctionPointer(p, typeof(T)) as T;
    }

    /// Like TC right after loading the plugin; returns false if the plugin has no ListSetDefaultParams.
    public static bool SetDefaultParams(string iniName)
    {
        if (setDefaultParams == null) return false;
        var dps = new ListDefaultParamStruct();
        dps.size = Marshal.SizeOf(typeof(ListDefaultParamStruct));
        dps.PluginInterfaceVersionLow = 12;   // interface 2.12, as in current TC versions
        dps.PluginInterfaceVersionHi = 2;
        dps.DefaultIniName = iniName;
        setDefaultParams(ref dps);
        return true;
    }

    public static IntPtr Load(IntPtr parent, string file) { return load(parent, file, 0); }
    public static void Close(IntPtr window) { close(window); }
    public static int Search(IntPtr window, string text) { return search(window, text, 0); }
    public static int SearchWith(IntPtr window, string text, int flags) { return search(window, text, flags); }
    /// null when the plugin does not export it.
    public static int? Send(IntPtr window, int command) { if (sendCommand == null) return null; return sendCommand(window, command, 0); }
    public static int? Print(IntPtr window, string file) { if (print == null) return null; return print(window, file, null, 0, IntPtr.Zero); }
}
'@

function Write-Status($s) { Set-Content -Path $StatusFile -Value $s -NoNewline }

[ListerHost]::Init((Resolve-Path $Dll).Path)
if (-not $IniFolder) { $IniFolder = Split-Path -Parent ([IO.Path]::GetFullPath($StatusFile)) }
New-Item -ItemType Directory -Force $IniFolder | Out-Null
[void][ListerHost]::SetDefaultParams((Join-Path $IniFolder 'lsplugin.ini'))

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Lister (mdview) - test host'
$form.Width = 980
$form.Height = 720
# Fits a 1024x768 screen (GitHub's runners).
$form.StartPosition = 'Manual'
$form.Left = 20
$form.Top = 20
$closedByPlugin = $false
# Keys the plugin passes on to TC's Lister window arrive here.
if ($LogFile) {
    $form.KeyPreview = $true
    $form.Add_KeyDown({ param($sender, $e) Add-Content -Path $LogFile -Value ("key {0}" -f [int]$e.KeyCode) })
}
$form.Add_FormClosing({ $script:closedByPlugin = $true })
$form.Show()
# When started with -WindowStyle Hidden, Windows applies "hidden" to the first
# window this process shows; show it again explicitly (SW_SHOWNORMAL).
[void][ListerHost]::ShowWindow($form.Handle, 1)
[System.Windows.Forms.Application]::DoEvents()

$child = [ListerHost]::Load($form.Handle, (Resolve-Path $Doc).ProviderPath)
if ($child -eq [IntPtr]::Zero) { Write-Status 'load-failed'; $form.Dispose(); exit 2 }
Write-Status 'loaded'

$loadedAt = Get-Date
$deadline = $loadedAt.AddSeconds($Seconds)
$searched = $false
$commandsDone = 0
while ((Get-Date) -lt $deadline -and -not $closedByPlugin) {
    [System.Windows.Forms.Application]::DoEvents()
    # Search once the document has had time to load, as a user pressing F7 would.
    if ($Search -and -not $searched -and (Get-Date) -gt $loadedAt.AddSeconds(6)) {
        [void][ListerHost]::Search($child, $Search)
        $searched = $true
    }
    if ($CommandFile -and $LogFile -and (Test-Path $CommandFile)) {
        $lines = @(Get-Content -Path $CommandFile)
        while ($commandsDone -lt $lines.Count) {
            $line = $lines[$commandsDone]; $commandsDone++
            $result = switch -Regex ($line) {
                '^search (\d+) (.+)$' { [ListerHost]::SearchWith($child, $Matches[2], [int]$Matches[1]) }
                '^send (\d+)$' { $r = [ListerHost]::Send($child, [int]$Matches[1]); if ($null -eq $r) { 'missing' } else { $r } }
                '^print$' { $r = [ListerHost]::Print($child, (Resolve-Path $Doc).ProviderPath); if ($null -eq $r) { 'missing' } else { $r } }
                default { 'unknown' }
            }
            Add-Content -Path $LogFile -Value ("result {0} {1}" -f $commandsDone, $result)
        }
    }
    Start-Sleep -Milliseconds 30
}

if ($closedByPlugin) {
    Write-Status 'closed-by-plugin'
} else {
    Write-Status 'timeout'
    [ListerHost]::Close($child)
    $form.Close()
}
$form.Dispose()
