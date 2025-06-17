using Microsoft.CommandPalette.Extensions;
using Microsoft.CommandPalette.Extensions.Toolkit;
using System.Collections.Generic;

namespace UtilitiesAutomateExtension.Pages
{
    internal class ContactSearch : ListPage
    {
        List<Person> _people = new();
        List<ListItem> items = new();

        public ContactSearch()
        {
            Icon = new IconInfo("\ue716");
            Title = "Contact Search";
            Name = "Contact Search";

            items.Add(new ListItem(new AnonymousCommand(action: updateItems) 
            {  
                Result = CommandResult.KeepOpen(),
                Icon = new IconInfo("\ue72c") // Use a suitable icon for update action
            }));

            _people = ListGen.LoadPeopleFromContactBook(@"C:\Users\Par149\Documents\Logs\ContactBook.txt");

            //AnonymousCommand _updateCommand = new AnonymousCommand(action: updateItems)
            //{
            //    Result = CommandResult.KeepOpen(),
            //    Icon = new IconInfo("\ue72c") // Use a suitable icon for update action
            //};

            //_updateCommand.Invoke(); // Initial load of items
        }

        public override IListItem[] GetItems()
        {
            //var items = new List<IListItem>();

            //this.IsLoading = true;

            //foreach (var person in _people)
            //{
            //    items.Add(new ListItem(new EmailOpenCommand(person.getEmail()))
            //    {
            //        Title = person.getName(),
            //        Subtitle = $"{person.getPhoneNumber()} {person.getEmail()}",
            //        //Icon = new IconInfo(@"C:\Path\To\Icon.png") // Replace with actual icon path
            //    });
            //    RaiseItemsChanged();
            //} 

            //this.IsLoading = false;

            //RaiseItemsChanged();

            return items.ToArray();
        }

        internal void updateItems()
        {
            this.IsLoading = true;
            //items.Clear();
            foreach (var person in _people)
            {
                items.Add(new ContactListItem(this, person.getName(), person.getEmail().Substring(1), person.getPhoneNumber()));
            }
            RaiseItemsChanged();
            this.IsLoading = false;
        }
    }
}
