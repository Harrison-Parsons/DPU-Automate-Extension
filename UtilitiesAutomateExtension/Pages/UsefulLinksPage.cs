using Microsoft.CommandPalette.Extensions;
using Microsoft.CommandPalette.Extensions.Toolkit;
using System;
using System.Runtime.InteropServices;


namespace UtilitiesAutomateExtension.Pages
{
    internal sealed partial class UsefulLinksPage : ListPage
    {
        public UsefulLinksPage()
        {
            Icon = new("\ue8a7");
            Title = "Useful Links";
            Name = "Open";
        }

        public override IListItem[] GetItems()
        {
            
            MessageBox(0,"Get Item called","Get Items",0x00001000);

            var hrms = new WebSearchCommand("https://ebiz-int.henrico.gov/OA_HTML/OA.jsp?OAFunc=OASIMPLEHOMEPAGE");
            var sharePoint = new WebSearchCommand("https://henricova.sharepoint.com/_layouts/15/sharepoint.aspx");
            var tasks = new WebSearchCommand("https://to-do.office.com/tasks/inbox");
            var flowchart = new WebSearchCommand("http://draw.io/");

            return [
                new ListItem(hrms){
                    Title = "Time Card",
                    Icon = new IconInfo(@"C:\Users\Par149\OneDrive - County of Henrico VA\Desktop\POwershell Scripting\faviconV2.png")
                },
                new ListItem(sharePoint){
                    Title = "SharePoint",
                    Icon = new IconInfo(@"C:\Users\Par149\OneDrive - County of Henrico VA\Desktop\POwershell Scripting\6296669_microsoft_office_office365_sharepoint_icon.png")
                },
                new ListItem(tasks){
                    Title = "ToDo Tasks",
                    Icon = new IconInfo(@"C:\Users\Par149\OneDrive - County of Henrico VA\Desktop\POwershell Scripting\logo.png")
                },
                new ListItem(flowchart){
                    Title = "Draw.io",
                    Icon = new IconInfo(@"C:\Users\Par149\OneDrive - County of Henrico VA\Desktop\POwershell Scripting\drawio.png")
                },
            ];
        }
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);
    }
}
