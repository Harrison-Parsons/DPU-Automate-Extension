using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace UtilitiesAutomateExtension.Pages
{
    /// <summary>
    /// Provides environment-specific file paths and user information for the application.
    /// </summary>
    internal class EnvDeclarations
    {
        /// <summary>
        /// The current user's username.
        /// </summary>
        public static string user = Environment.UserName;

        /// <summary>
        /// The path to the user's local AppData folder.
        /// </summary>
        public static string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

        /// <summary>
        /// The path to the user's Documents folder.
        /// </summary>
        public static string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        /// <summary>
        /// The file path to the Outlook Address Book (OAB) file.
        /// </summary>
        public static string oabFilePath = Path.Combine(
            localAppData,
            @"Microsoft\Outlook\Offline Address Books\9f97379f-d3ed-41a0-a0c0-9f50ecc3f3e8\udetails.oab"
        );

        /// <summary>
        /// The file path to the application's contact book file.
        /// </summary>
        public static string contactBookFilePath = Path.Combine(
            documents,
            @"Logs\ContactBook.txt"
        );

        /// <summary>
        /// The file path to the application's log file.
        /// </summary>
        public static string logsFilePath = Path.Combine(
            documents,
            @"Logs\ContactSearchLog.txt"
        );

        /// <summary>
        /// Ensures that the directories for the contact book and log files exist.
        /// </summary>
        public static void EnsureDirectories()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(contactBookFilePath));
            Directory.CreateDirectory(Path.GetDirectoryName(logsFilePath));
        }
    }
}
