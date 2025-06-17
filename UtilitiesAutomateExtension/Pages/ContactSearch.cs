using Microsoft.CommandPalette.Extensions;
using Microsoft.CommandPalette.Extensions.Toolkit;
using System.Collections.Generic;

namespace UtilitiesAutomateExtension.Pages
{
    /// <summary>
    /// Represents a page for searching and displaying contacts.
    /// </summary>
    internal class ContactSearch : ListPage
    {
        /// <summary>
        /// The list of people loaded from the contact book.
        /// </summary>
        List<Person> _people = new();

        /// <summary>
        /// The list of items displayed on the page.
        /// </summary>
        List<ListItem> items = new();

        /// <summary>
        /// Initializes a new instance of the <see cref="ContactSearch"/> class.
        /// Loads contacts and sets up the update command.
        /// </summary>
        public ContactSearch()
        {
            Icon = new IconInfo("\ue716");
            Title = "Contact Search";
            Name = "Contact Search";

            items.Add(new ListItem(new AnonymousCommand(action: updateItems)
            {
                Result = CommandResult.KeepOpen(),
                Icon = new IconInfo("\ue72c")
            }));

            _people = ListGen.LoadPeopleFromContactBook(@"C:\Users\Par149\Documents\Logs\ContactBook.txt");
        }

        /// <summary>
        /// Gets the items to be displayed on the page.
        /// </summary>
        /// <returns>An array of <see cref="IListItem"/> representing the items.</returns>
        public override IListItem[] GetItems()
        {
            return items.ToArray();
        }

        /// <summary>
        /// Updates the items list with the current people from the contact book.
        /// </summary>
        internal void updateItems()
        {
            this.IsLoading = true;
            foreach (var person in _people)
            {
                items.Add(new ContactListItem(this, person.getName(), person.getEmail().Substring(1), person.getPhoneNumber()));
            }
            RaiseItemsChanged();
            this.IsLoading = false;
        }
    }
}
