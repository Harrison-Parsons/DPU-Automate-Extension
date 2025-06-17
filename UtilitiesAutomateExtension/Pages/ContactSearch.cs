using Microsoft.CommandPalette.Extensions;
using Microsoft.CommandPalette.Extensions.Toolkit;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

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

            EnvDeclarations.EnsureDirectories();

            /// <summary>
            /// Loads people from the contact book file using the path from <see cref="EnvDeclarations.contactBookFilePath"/>.
            /// </summary>
            _people = ListGen.LoadPeopleFromContactBook(EnvDeclarations.contactBookFilePath);
            //MessageBox(0,$"{EnvDeclarations.contactBookFilePath}", "Contact Book Path", 0x00000000);
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

        /// <summary>
        /// Displays a message box using the Win32 API.
        /// </summary>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);
    }
}
