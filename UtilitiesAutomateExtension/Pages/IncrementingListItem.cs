using Microsoft.CommandPalette.Extensions.Toolkit;

namespace UtilitiesAutomateExtension.Pages
{
    internal sealed partial class IncrementingListItem : ListItem
    {
        public IncrementingListItem(UtilitiesAutomateExtensionPage page) :
            base(new NoOpCommand())
        {
            _page = page;
            Command = new AnonymousCommand(action: _page.Increment) { Result = CommandResult.KeepOpen() };
            Title = "Increment";
        }

        private UtilitiesAutomateExtensionPage _page;
    }
}
