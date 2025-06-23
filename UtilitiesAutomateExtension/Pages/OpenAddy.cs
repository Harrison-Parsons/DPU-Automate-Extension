using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.CommandPalette.Extensions;
using Microsoft.CommandPalette.Extensions.Toolkit;

namespace UtilitiesAutomateExtension.Pages
{
    internal partial class OpenAddy : InvokableCommand // Marked as partial to fix CsWinRT1028
    {
        internal string shortcut;

        public CommandResult Result { get; set; } = CommandResult.KeepOpen();

        public OpenAddy(string textShortcut)
        {
            Icon = IconHelpers.FromRelativePath("\uE715");
            Name = "OpenAddy";
            shortcut = textShortcut;
        }

        // Win32 API imports for window management
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private const int SW_RESTORE = 9;
        private const int SW_MAXIMIZE = 3;

        [STAThread]
        public override ICommandResult Invoke()
        {
            // Find all running Outlook processes and get the main window handle if available
            Process[] outlookProcs = Process.GetProcessesByName("OUTLOOK");
            IntPtr outlookHwnd = IntPtr.Zero;

            foreach (var proc in outlookProcs)
            {
                if (proc.MainWindowHandle != IntPtr.Zero)
                {
                    outlookHwnd = proc.MainWindowHandle;
                    break;
                }
            }

            if (outlookHwnd != IntPtr.Zero)
            {
                // If Outlook is already running, maximize and bring it to the foreground
                ShowWindow(outlookHwnd, SW_MAXIMIZE);
                SetForegroundWindow(outlookHwnd);
            }
            else
            {
                // Start Outlook if not running
                Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
                // Wait for Outlook to start and create a window (up to 20 seconds)
                for (int i = 0; i < 40; i++)
                {
                    Thread.Sleep(500);
                    outlookProcs = Process.GetProcessesByName("OUTLOOK");
                    foreach (var proc in outlookProcs)
                    {
                        if (proc.MainWindowHandle != IntPtr.Zero)
                        {
                            outlookHwnd = proc.MainWindowHandle;
                            break;
                        }
                    }
                    if (outlookHwnd != IntPtr.Zero)
                        break;
                }
                if (outlookHwnd != IntPtr.Zero)
                {
                    // Maximize and bring the new Outlook window to the foreground
                    ShowWindow(outlookHwnd, SW_MAXIMIZE);
                    SetForegroundWindow(outlookHwnd);
                }
            }

            // Wait until Outlook is the foreground window and add a small extra delay for UI readiness
            for (int i = 0; i < 20; i++)
            {
                if (GetForegroundWindow() == outlookHwnd)
                    break;
                Thread.Sleep(200);
                SetForegroundWindow(outlookHwnd);
            }
            Thread.Sleep(1500); // Extra delay to ensure Outlook is ready for input

            // Send the keyboard shortcut (e.g., Ctrl+Shift+B) to Outlook
            SendKeys.SendWait(shortcut);

            return Result;
        }
    }
}
