using Microsoft.CommandPalette.Extensions;
using Microsoft.CommandPalette.Extensions.Toolkit;


namespace UtilitiesAutomateExtension.Pages
{
    internal class PSOpenCommand : InvokableCommand
    {
        public string psPath = @"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe";
        public string scriptPath = @"C:\Users\Par149\OneDrive - County of Henrico VA\Desktop\POwershell Scripting";
        public string command;

        public CommandResult Result { get; set; } = CommandResult.KeepOpen();

        public PSOpenCommand(string type)
        {
            command = type;
            Name = "PS";
            Icon = new IconInfo(@"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe");
        }

        public override ICommandResult Invoke()
        {
            ShellHelpers.OpenInShell(psPath, command, scriptPath);
            return Result;
        }
    }
}
