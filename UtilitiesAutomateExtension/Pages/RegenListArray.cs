using Microsoft.CommandPalette.Extensions;
using Microsoft.CommandPalette.Extensions.Toolkit;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace UtilitiesAutomateExtension.Pages
{
    /// <summary>
    /// Command to regenerate the contact list from the Outlook address book using a single-threaded approach.
    /// </summary>
    public class RegenListArray : InvokableCommand
    {
        static String filePath = @"C:\Users\Par149\AppData\Local\Microsoft\Outlook\Offline Address Books\9f97379f-d3ed-41a0-a0c0-9f50ecc3f3e8\udetails.oab";
        static int startLine = 11;
        static string[] strings = System.IO.File.ReadAllLines(filePath);
        public static String contents = String.Join(Environment.NewLine, strings, startLine, strings.Length - startLine);
        String pattern = "Microsoft Private MDB";
        MatchCollection matches;
        DateTime dateTime = DateTime.Now;

        TimeSpan elapsed;

        /// <summary>
        /// Gets or sets the command result.
        /// </summary>
        public CommandResult Result { get; set; } = CommandResult.KeepOpen();

        /// <summary>
        /// Initializes a new instance of the <see cref="RegenListArray"/> class.
        /// </summary>
        public RegenListArray()
        {
            Name = "Regen List Array";
            Icon = new IconInfo(@"C:\Users\Par149\OneDrive - County of Henrico VA\Desktop\POwershell Scripting\faviconV2.png");
        }

        /// <summary>
        /// Invokes the regeneration process, parsing the address book and writing contacts to file.
        /// </summary>
        /// <returns>The result of the command invocation.</returns>
        public override ICommandResult Invoke()
        {
            File.WriteAllText(@"C:\Users\Par149\Documents\Logs\logs.txt", string.Empty); // Clear log file
            File.WriteAllText(@"C:\Users\Par149\Documents\Logs\ContactBook.txt", string.Empty); // Clear contact book file
            try
            {
                using (var logStream = new FileStream(@"C:\Users\Par149\Documents\Logs\logs.txt", FileMode.Create, FileAccess.Write, FileShare.Read, 4096, FileOptions.WriteThrough))
                using (var logWriter = new StreamWriter(logStream))
                using (var contactStream = new FileStream(@"C:\Users\Par149\Documents\Logs\ContactBook.txt", FileMode.Create, FileAccess.Write, FileShare.Read, 4096, FileOptions.WriteThrough))
                using (var contactWriter = new StreamWriter(contactStream))
                {
                    logWriter.AutoFlush = false;
                    contactWriter.AutoFlush = false;

                    Stopwatch stopwatch = Stopwatch.StartNew();

                    if (string.IsNullOrEmpty(contents) || string.IsNullOrEmpty(pattern))
                        throw new InvalidOperationException("Input data is missing.");

                    var regex = new Regex(pattern, RegexOptions.Compiled);
                    matches = regex.Matches(contents);

                    int count = 0;
                    DateTime dateTime;
                    int contentsLength = contents.Length;

                    foreach (Match match in matches)
                    {
                        try
                        {
                            dateTime = DateTime.Now;
                            logWriter.Write(count);
                            logWriter.Write(dateTime.ToString(" [MM/dd/yyyy HH:mm:ss.fff] "));

                            int index = match.Index;
                            int exIndex = backIndex("ExchangeLabs", index);
                            int atIndex = backIndex("@", exIndex);

                            if (exIndex < 0 || atIndex < 0 || atIndex >= contentsLength || exIndex > contentsLength || exIndex <= atIndex)
                            {
                                logWriter.WriteLine("Invalid indexes for tag extraction.");
                                continue;
                            }

                            logWriter.Write($"Index: {index} ExIndex: {exIndex} AtIndex: {atIndex} | ");

                            string tagSub = contents.Substring(atIndex, exIndex - atIndex);
                            int slashIdx = tagSub.IndexOf("/");
                            if (slashIdx <= 1)
                            {
                                logWriter.WriteLine("Invalid tag format.");
                                continue;
                            }
                            string emailTag = tagSub.Substring(1, slashIdx - 1)
                                .Replace("\0", string.Empty).Trim().Replace("/", string.Empty);

                            if (emailTag.Length > 100)
                            {
                                logWriter.WriteLine($"Skipped due to email tag length > 100 characters: {emailTag}\n");
                                continue;
                            }

                            string email = $"{emailTag}@henrico.gov";
                            logWriter.Write($"Email: {email} | ");

                            int nextTagIndex = backIndex(emailTag, atIndex);
                            int nameIndex = nextTagIndex + emailTag.Length;

                            if (nameIndex < 0 || nameIndex >= contentsLength || atIndex <= nameIndex)
                            {
                                logWriter.WriteLine("Invalid name indexes.");
                                continue;
                            }

                            logWriter.Write($"Next Tag Index: {nextTagIndex} Name Index: {nameIndex} EmailTag.Length: {emailTag.Length} | ");

                            string nameChunk = contents.Substring(nameIndex, atIndex - nameIndex)
                                .Replace("\0", string.Empty).Trim().Replace("/", string.Empty);

                            int delimIdx = nameChunk.IndexOf("�");
                            if (delimIdx < 0)
                            {
                                logWriter.WriteLine("Name delimiter not found.");
                                continue;
                            }
                            string nameRaw = nameChunk.Substring(0, delimIdx)
                                .Replace("\0", string.Empty).Trim().Replace("/", string.Empty);
                            logWriter.Write($"Name Raw: {nameRaw} | ");

                            if (nameRaw.Length > 150)
                            {
                                logWriter.WriteLine($"Skipped due to name length > 150 characters: {nameRaw}\n");
                                continue;
                            }

                            string fullName = string.Empty;
                            int commaIdx = nameRaw.IndexOf(",");
                            if (commaIdx == -1)
                            {
                                fullName = nameRaw.Trim();
                                logWriter.Write($"Full Name: {fullName} | ");
                            }
                            else
                            {
                                string lastName = nameRaw.Substring(0, commaIdx);
                                string firstName = nameRaw.Substring(commaIdx + 1).TrimStart(' ');
                                fullName = $"{firstName} {lastName}";
                                logWriter.Write($"First Name: {firstName} Last Name: {lastName} Full Name: {fullName} | ");
                            }

                            string combString = contents.Substring(index);
                            int endIdx = combString.IndexOf("\u0006");
                            if (endIdx < 0)
                            {
                                logWriter.WriteLine("End delimiter not found.");
                                continue;
                            }
                            string narrowedString = combString.Substring(0, endIdx);

                            var phoneRegex = new Regex(@"\d{3}-\d{3}-\d{4}|\(\d{3}\) \d{3}-\d{4}|\d{10}|\d{3}-\d{4}|\d{7}", RegexOptions.Compiled);
                            Match phoneMatch = phoneRegex.Match(combString);

                            string phoneNumber = phoneMatch.Success ? phoneMatch.Value : "No Number Found";

                            contactWriter.WriteLine($"{fullName}|{phoneNumber} @{email}");

                            count++;
                            logWriter.Write($"Count: {count} | ");
                            logWriter.WriteLine($"Person Added: {fullName}");
                        }
                        catch (Exception ex)
                        {
                            logWriter.WriteLine($"Error processing match: {ex.Message}");
                        }
                    }
                    stopwatch.Stop();
                    logWriter.Flush();
                    contactWriter.Flush();
                    elapsed = stopwatch.Elapsed;
                }
                // Sanitize ContactBook.txt to remove lines > 80 chars
                string contactBookPath = @"C:\Users\Par149\Documents\Logs\ContactBook.txt";
                var allLines = File.ReadAllLines(contactBookPath);
                var sanitizedLines = new List<string>();
                foreach (var line in allLines)
                {
                    if (line.Length <= 80)
                        sanitizedLines.Add(line);
                }
                File.WriteAllLines(contactBookPath, sanitizedLines);

                MessageBox(0, $"Elapsed Time: {elapsed.TotalMilliseconds / 1000} s", "People Length", 0x00001000);

                return Result;
            }
            catch (Exception ex)
            {
                MessageBox(0, $"Fatal error: {ex.Message}", "Error", 0x00001000);
                throw;
            }
        }

        /// <summary>
        /// Finds the last index of a substring before a given index in the contents.
        /// </summary>
        /// <param name="toFind">The substring to find.</param>
        /// <param name="currIndex">The index to search backwards from.</param>
        /// <returns>The last index of the substring, or -1 if not found.</returns>
        private static int backIndex(String toFind, int currIndex)
        {
            int index = -1;
            String contentsToSearch = contents.Substring(0, currIndex);
            index = contentsToSearch.LastIndexOf(toFind);
            return index;
        }

        /// <summary>
        /// Displays a message box using the Win32 API.
        /// </summary>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);
    }
}
