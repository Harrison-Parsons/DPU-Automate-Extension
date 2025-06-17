using Microsoft.CommandPalette.Extensions.Toolkit;

namespace UtilitiesAutomateExtension.Pages
{
    internal class WebSearchCommand : InvokableCommand
    {

        public string ffPath = @"C:\Users\Par149\AppData\Local\Mozilla Firefox\firefox.exe";

        public CommandResult Result { get; set; } = CommandResult.KeepOpen();

        public string openingLink;

        public WebSearchCommand(string link)
        {
            openingLink = link;
            Name = $"Open {link}";
            Icon = new IconInfo("\ue8a7");
        }

        public override CommandResult Invoke()
        {
            ShellHelpers.OpenInShell(ffPath, openingLink);
            return Result;
        }
    }
}
