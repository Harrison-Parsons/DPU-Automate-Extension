Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class Win32Functions
{
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

$outlookProcess = Get-Process -Name outlook | Select-Object -First 1
if ($outlookProcess) {
    $handle = $outlookProcess.MainWindowHandle
    # SW_RESTORE (9) restores a minimized window and activates it.
    # SW_SHOW (5) activates a window and displays it in its current size and position.
    [Win32Functions]::ShowWindow($handle, 9)
} else {
    Start-Process "outlook.exe"
}