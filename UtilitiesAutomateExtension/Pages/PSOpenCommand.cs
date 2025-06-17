using Microsoft.CommandPalette.Extensions;
using Microsoft.CommandPalette.Extensions.Toolkit;
using System.Diagnostics;

namespace UtilitiesAutomateExtension.Pages
{
    /// <summary>
    /// Command to open PowerShell and execute a specified command.
    /// Uses the system PATH to resolve the PowerShell executable, ensuring compatibility across Windows environments.
    /// </summary>
    internal class PSOpenCommand : InvokableCommand
    {
        /// <summary>
        /// The PowerShell executable name. This allows Windows to resolve the correct path via the system PATH environment variable.
        /// </summary>
        public string psExecutable = "powershell.exe";

        /// <summary>
        /// The command to execute in PowerShell.
        /// </summary>
        public string command;

        /// <summary>
        /// Gets or sets the command result.
        /// </summary>
        public CommandResult Result { get; set; } = CommandResult.KeepOpen();

        /// <summary>
        /// Initializes a new instance of the <see cref="PSOpenCommand"/> class.
        /// </summary>
        /// <param name="type">The command to execute in PowerShell.</param>
        public PSOpenCommand(string type)
        {
            command = type;
            Name = "PS";
            Icon = new IconInfo(""); // Unicode for PowerShell, or keep your icon
        }

        /// <summary>
        /// Invokes the PowerShell command using the system-resolved executable.
        /// </summary>
        /// <returns>The result of the command invocation.</returns>
        public override ICommandResult Invoke()
        {
            // Use ProcessStartInfo for robust launching and compatibility with the system PATH.
            var psi = new ProcessStartInfo
            {
                FileName = psExecutable,
                Arguments = $"-NoExit -Command \"{command}\"",
                UseShellExecute = true
            };
            Process.Start(psi);
            return Result;
        }
    }
}
