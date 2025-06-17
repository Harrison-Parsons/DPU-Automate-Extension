using Microsoft.CommandPalette.Extensions.Toolkit;
using System.Diagnostics;

namespace UtilitiesAutomateExtension.Pages
{
    /// <summary>
    /// Command to open a web link in the user's default browser.
    /// Uses Process.Start with UseShellExecute to ensure the default browser is used, not a hardcoded path.
    /// </summary>
    internal class WebSearchCommand : InvokableCommand
    {
        /// <summary>
        /// Gets or sets the command result.
        /// </summary>
        public CommandResult Result { get; set; } = CommandResult.KeepOpen();

        /// <summary>
        /// The URL to open in the default browser.
        /// </summary>
        public string openingLink;

        /// <summary>
        /// Initializes a new instance of the <see cref="WebSearchCommand"/> class.
        /// </summary>
        /// <param name="link">The URL to open.</param>
        public WebSearchCommand(string link)
        {
            openingLink = link;
            Name = $"Open {link}";
            Icon = new IconInfo("\ue8a7");
        }

        /// <summary>
        /// Invokes the command to open the URL in the user's default browser.
        /// </summary>
        /// <returns>The result of the command invocation.</returns>
        public override CommandResult Invoke()
        {
            // Open the URL in the default browser using ProcessStartInfo and UseShellExecute.
            var psi = new ProcessStartInfo
            {
                FileName = openingLink,
                UseShellExecute = true
            };
            Process.Start(psi);
            return Result;
        }
    }
}
