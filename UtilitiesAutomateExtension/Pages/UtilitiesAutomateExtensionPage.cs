// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using Microsoft.CommandPalette.Extensions;
using Microsoft.CommandPalette.Extensions.Toolkit;
using System.Collections.Generic;
using UtilitiesAutomateExtension.Pages;

namespace UtilitiesAutomateExtension;

internal sealed partial class UtilitiesAutomateExtensionPage : ListPage
{
    public UtilitiesAutomateExtensionPage()
    {
        Icon = IconHelpers.FromRelativePath("Assets\\StoreLogo.png");
        Title = "DPU Automation";
        Name = "Open";

        _items = [new IncrementingListItem(this) { Subtitle = $"Item 0" }];

    }

    public override IListItem[] GetItems()
    {
        var ps = new PSOpenCommand(@".\ContactSearchUTF8.ps1");
        var email = new EmailOpenCommand("par149@henrico.gov");
        var regen = new RegenListArray();
        var parallelRegen = new RegenListArrayParallel();
        var ContactSearch = new ContactSearch();

        return [
            new ListItem(ContactSearch){
                Title = "Contact Search",
                Icon = new IconInfo(@"C:\Users\par149\OneDrive - County of Henrico VA\Desktop\POwershell Scripting\contact.png")
            },
            new ListItem(new AnonymousCommand(action: ()=> {
                parallelRegen.Invoke();
            })),
            //new ListItem(ps){
            //    Title = "Contact Search"
            //},
            //new ListItem(new UsefulLinksPage()){
            //    Title = "Useful Links"
            //},
            //new ListItem(email){
            //    Title = "Email Me",
            //    Icon = new IconInfo(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"),
            //},
            //new ListItem(regen){
            //    Title = "Regen List",
            //},
            //new ContactListItem(){
            //    Title = "Test Contact",
            //},
            
        ];
    }

    internal void Increment()
    {
        _items.Add(new IncrementingListItem(this) { Subtitle = $"Item {_items.Count}" });
        RaiseItemsChanged();
    }

    private List<ListItem> _items;
}
