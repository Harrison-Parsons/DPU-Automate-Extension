using Microsoft.CommandPalette.Extensions.Toolkit;
using Microsoft.CommandPalette.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Storage.AccessCache;

namespace UtilitiesAutomateExtension.Pages
{
    /// <summary>
    /// Command to update the contact list items on a <see cref="ContactSearch"/> page.
    /// </summary>
    internal class ContactUpdate : InvokableCommand
    {
        /// <summary>
        /// The associated <see cref="ContactSearch"/> page.
        /// </summary>
        ContactSearch _page = new ContactSearch();

        /// <summary>
        /// The list of people to update.
        /// </summary>
        List<Person> _people = new List<Person>();

        /// <summary>
        /// The list of items to update.
        /// </summary>
        List<ListItem> _items = new List<ListItem>();

        /// <summary>
        /// Gets or sets the command result.
        /// </summary>
        public CommandResult Result { get; set; } = CommandResult.KeepOpen();

        /// <summary>
        /// Initializes a new instance of the <see cref="ContactUpdate"/> class.
        /// </summary>
        /// <param name="page">The <see cref="ContactSearch"/> page to update.</param>
        /// <param name="people">The list of people to update.</param>
        /// <param name="items">A reference to the list of items to update.</param>
        public ContactUpdate(ContactSearch page, List<Person> people, ref List<ListItem> items)
        {
            _page = page;
            _people = people;
            _items = items;
        }

        /// <summary>
        /// Invokes the update operation, adding contact items to the list.
        /// </summary>
        /// <returns>The result of the command invocation.</returns>
        public override ICommandResult Invoke()
        {
            _page.IsLoading = true;
            // Optionally clear items here if needed.
            foreach (var person in _people)
            {
                _items.Add(new ContactListItem(_page, person.getName(), person.getEmail(), person.getPhoneNumber()));
            }
            // Optionally raise items changed event here if needed.
            _page.IsLoading = false;

            return Result;
        }
    }
}
