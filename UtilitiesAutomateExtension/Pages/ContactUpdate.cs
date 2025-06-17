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
    internal class ContactUpdate : InvokableCommand
    {
        ContactSearch _page = new ContactSearch();
        List<Person> _people = new List<Person>();
        List<ListItem> _items = new List<ListItem>();

        public CommandResult Result { get; set; } = CommandResult.KeepOpen();

        public ContactUpdate(ContactSearch page, List<Person> people, ref List<ListItem> items) { 
            _page = page;
            _people = people;
            _items = items;
        }

        public override ICommandResult Invoke()
        {
            _page.IsLoading = true;
            //items.Clear();
            foreach (var person in _people)
            {
                _items.Add(new ContactListItem(_page, person.getName(), person.getEmail(), person.getPhoneNumber()));
            }
            //_page.RaiseItemsChanged();
            _page.IsLoading = false;

            return Result;
        }
    }
}
