// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using Microsoft.CommandPalette.Extensions;
using Microsoft.CommandPalette.Extensions.Toolkit;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using UtilitiesAutomateExtension.Pages;

namespace UtilitiesAutomateExtension;

internal sealed partial class UtilitiesAutomateExtensionPage : ListPage
{
    public UtilitiesAutomateExtensionPage()
    {
        Icon = IconHelpers.FromRelativePath("Assets\\StoreLogo.png");
        Title = "DPU Automation";
        Name = "Open";

    }

    public override IListItem[] GetItems()
    {
        //var ps = new PSOpenCommand(@".\ContactSearchUTF8.ps1");
        //var email = new EmailOpenCommand("par149@henrico.gov");
        //var parallelRegen = new RegenListArrayParallel();
        var ContactSearch = new ContactSearch();
        var openAddy = new OpenAddy("^(+(b))");
        var openSearch = new OpenAddy("^(e)");

        return [
            new ListItem(ContactSearch){
                Title = "Contact Search",
                Icon = new IconInfo("\uEA8C")
            },
            new ListItem(new AnonymousCommand(action: ()=> {
                UpdateItems();
                Icon = new IconInfo("\uE777");
            })),
            new ListItem(openAddy){
                Title = "Open Addy",
                Icon = new IconInfo("\uE715")
            },
            new ListItem(openSearch){
                Title = "Open Search",
                Icon = new IconInfo("\uE721")
            },
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

    internal void UpdateItems()
    {

        var parallelRegen = new RegenListArrayParallel();

        this.IsLoading = true;
        Stopwatch stopwatch = Stopwatch.StartNew();

        parallelRegen.Invoke();

        stopwatch.Stop();
        this.IsLoading = false;

        var elapsedTime = stopwatch.ElapsedMilliseconds;
        MessageBox(IntPtr.Zero, $"Regen completed in {elapsedTime/1000} s", "Regen List", 0x00000000 | 0x00001000 | 0x00000040);
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);

    private List<ListItem> _items;
}
