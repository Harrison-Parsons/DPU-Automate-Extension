using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UtilitiesAutomateExtension.Pages
{
    internal class EnvDeclarations
    {

        static public string oabFilePath;
        static public string contactBookFilePath;
        static public string logsFilePath;
        static public string user = Environment.UserName;

        public EnvDeclarations()
        {
            oabFilePath = $@"C:\Users\{user}\AppData\Local\Microsoft\Outlook\Offline Address Books\9f97379f-d3ed-41a0-a0c0-9f50ecc3f3e8\udetails.oab";
            contactBookFilePath = $@"C:\Users\{user}\Documents\Logs\ContactBook.txt";
            logsFilePath = $@"C:\Users\{user}\Documents\Logs\ContactSearchLog.txt";
        }
    }
}
