using Microsoft.CommandPalette.Extensions;
using Microsoft.CommandPalette.Extensions.Toolkit;

namespace UtilitiesAutomateExtension.Pages
{
    internal class EmailOpenCommand : InvokableCommand
    {
        public string email;
        public string emailPath = @"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe";

        public CommandResult Result { get; set; } = CommandResult.KeepOpen();

        public EmailOpenCommand(string input)
        {
            email = input;
            Name = "Open Email Draft";
            Icon = new IconInfo(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
        }

        public override ICommandResult Invoke()
        {
            ShellHelpers.OpenInShell(emailPath, "Start-Process mailto:" + email);
            return Result;
        }
    }
}