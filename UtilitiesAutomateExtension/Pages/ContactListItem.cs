using Microsoft.CommandPalette.Extensions.Toolkit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UtilitiesAutomateExtension.Pages
{
    internal class ContactListItem : ListItem
    {
        public ContactListItem(ContactSearch page,string name, string email, string number) : base(new NoOpCommand())
        {
            
            _page = page;

            Title = $"{name}";
            Subtitle = $"{email} | {number}";
            Icon = new IconInfo("\ue716");

            EmailOpenCommand emailCommand = new EmailOpenCommand("par149@henrico.gov");

            Command = new AnonymousCommand(action: () => 
            {
                // This is where you would define what happens when the contact is clicked
                // For example, you could open an email client with the person's email address
                emailCommand.Invoke();
            });
        }

        private readonly ContactSearch _page;
    }
}
