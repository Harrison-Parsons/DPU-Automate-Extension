using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace UtilitiesAutomateExtension.Pages
{
    /// <summary>
    /// Command to regenerate the contact list from the Outlook address book using a parallelized approach.
    /// </summary>
    public class RegenListArrayParallel
    {
        /// <summary>
        /// The list of people loaded after regeneration.
        /// </summary>
        List<Person> people = new List<Person>();

        /// <summary>
        /// Provides environment-specific file paths and user information.
        /// </summary>
        EnvDeclarations env = new EnvDeclarations();

        /// <summary>
        /// The file path to the Outlook Address Book (OAB) file, initialized from <see cref="EnvDeclarations.oabFilePath"/>.
        /// </summary>
        static string filePath = EnvDeclarations.oabFilePath;

        static int startLine = 11;

        /// <summary>
        /// The lines read from the OAB file, using the path from <see cref="EnvDeclarations.oabFilePath"/>.
        /// </summary>
        static string[] strings = System.IO.File.ReadAllLines(EnvDeclarations.oabFilePath);

        /// <summary>
        /// The contents of the OAB file, joined as a single string.
        /// </summary>
        public static string contents = String.Join(Environment.NewLine, strings, startLine, strings.Length - startLine);

        string pattern = "Microsoft Private MDB";
        MatchCollection matches;
        DateTime dateTime = DateTime.Now;
        TimeSpan elapsed;

        /// <summary>
        /// Gets or sets the name of the command.
        /// </summary>
        String Name { get; set; }

        /// <summary>
        /// Gets or sets the icon path for the command.
        /// </summary>
        String Icon { get; set; }

        // Log and contact queues and background writer tasks
        private BlockingCollection<string> logQueue = new(10000);
        private BlockingCollection<string> contactQueue = new(10000);
        private Task logWriterTask;
        private Task contactWriterTask;

        /// <summary>
        /// Initializes a new instance of the <see cref="RegenListArrayParallel"/> class.
        /// </summary>
        public RegenListArrayParallel()
        {
            Name = "Regen List Array Parallel";
            //Icon = @"C:\Users\Par149\OneDrive - County of Henrico VA\Desktop\POwershell Scripting\faviconV2.png";
        }

        /// <summary>
        /// Invokes the parallel regeneration process, parsing the address book and writing contacts to file.
        /// Uses <see cref="EnvDeclarations.logsFilePath"/> and <see cref="EnvDeclarations.contactBookFilePath"/> for output.
        /// </summary>
        public void Invoke()
        {
            //MessageBox(0, "Regen List Array Parallel Invoked", "Regen List Array Parallel", 0x00001000);

            File.WriteAllText(EnvDeclarations.logsFilePath.Replace("ContactSearchLog.txt", "logs.csv"), string.Empty); // Clear log file
            File.WriteAllText(EnvDeclarations.contactBookFilePath, string.Empty); // Clear contact book file

            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();

                if (string.IsNullOrEmpty(contents) || string.IsNullOrEmpty(pattern))
                    throw new InvalidOperationException("Input data is missing.");

                var regex = new Regex(pattern, RegexOptions.Compiled);
                var phoneRegex = new Regex(@"\d{3}-\d{3}-\d{4}|\(\d{3}\) \d{3}-\d{4}|\d{10}|\d{3}-\d{4}|\d{7}", RegexOptions.Compiled);
                matches = regex.Matches(contents);

                int contentsLength = contents.Length;

                // Start background log writer
                logWriterTask = Task.Run(() =>
                {
                    using var writer = new StreamWriter(EnvDeclarations.logsFilePath.Replace("ContactSearchLog.txt", "logs.csv"));
                    writer.WriteLine("Index,Timestamp,Status,Email,FullName,PhoneNumber,Message");
                    foreach (var log in logQueue.GetConsumingEnumerable())
                        writer.WriteLine(log);
                });

                // Start background contact writer
                contactWriterTask = Task.Run(() =>
                {
                    using var writer = new StreamWriter(EnvDeclarations.contactBookFilePath);
                    foreach (var contact in contactQueue.GetConsumingEnumerable())
                        writer.WriteLine(contact);
                });

                Parallel.ForEach(
                    Partitioner.Create(0, matches.Count),
                    () => 0, // dummy local state
                    (range, state, localState) =>
                    {
                        for (int i = range.Item1; i < range.Item2; i++)
                        {
                            try
                            {
                                var match = matches[i];
                                DateTime dateTime = DateTime.Now;

                                int index = match.Index;
                                int exIndex = backIndex("ExchangeLabs", index);
                                int atIndex = backIndex("@", exIndex);

                                if (exIndex < 0 || atIndex < 0 || atIndex >= contentsLength || exIndex > contentsLength || exIndex <= atIndex)
                                {
                                    logQueue.Add($"{i},{dateTime:MM/dd/yyyy HH:mm:ss.fff},InvalidIndexes,,,,Invalid indexes for tag extraction.");
                                    continue;
                                }

                                string tagSub = contents.Substring(atIndex, exIndex - atIndex);
                                int slashIdx = tagSub.IndexOf("/");
                                if (slashIdx <= 1)
                                {
                                    logQueue.Add($"{i},{dateTime:MM/dd/yyyy HH:mm:ss.fff},InvalidTagFormat,,,,Invalid tag format.");
                                    continue;
                                }
                                string emailTag = tagSub.Substring(1, slashIdx - 1)
                                    .Replace("\0", string.Empty).Trim().Replace("/", string.Empty);

                                if (emailTag.Length > 100)
                                {
                                    logQueue.Add($"{i},{dateTime:MM/dd/yyyy HH:mm:ss.fff},EmailTagTooLong,,,{emailTag},Skipped due to email tag length > 100 characters");
                                    continue;
                                }

                                string email = $"{emailTag}@henrico.gov";

                                int nextTagIndex = backIndex(emailTag, atIndex);
                                int nameIndex = nextTagIndex + emailTag.Length;

                                if (nameIndex < 0 || nameIndex >= contentsLength || atIndex <= nameIndex)
                                {
                                    logQueue.Add($"{i},{dateTime:MM/dd/yyyy HH:mm:ss.fff},InvalidNameIndexes,{email},,,Invalid name indexes.");
                                    continue;
                                }

                                string nameChunk = contents.Substring(nameIndex, atIndex - nameIndex)
                                    .Replace("\0", string.Empty).Trim().Replace("/", string.Empty);

                                int delimIdx = nameChunk.IndexOf("�");
                                if (delimIdx < 0)
                                {
                                    logQueue.Add($"{i},{dateTime:MM/dd/yyyy HH:mm:ss.fff},NameDelimiterNotFound,{email},,,Name delimiter not found.");
                                    continue;
                                }
                                string nameRaw = nameChunk.Substring(0, delimIdx)
                                    .Replace("\0", string.Empty).Trim().Replace("/", string.Empty);

                                if (nameRaw.Length > 150)
                                {
                                    logQueue.Add($"{i},{dateTime:MM/dd/yyyy HH:mm:ss.fff},NameTooLong,{email},{nameRaw},,Skipped due to name length > 150 characters");
                                    continue;
                                }

                                string fullName = string.Empty;
                                int commaIdx = nameRaw.IndexOf(",");
                                if (commaIdx == -1)
                                {
                                    fullName = nameRaw.Trim();
                                }
                                else
                                {
                                    string lastName = nameRaw.Substring(0, commaIdx);
                                    string firstName = nameRaw.Substring(commaIdx + 1).TrimStart(' ');
                                    fullName = $"{firstName} {lastName}";
                                }

                                int endIdx = contents.IndexOf('\u0006', index);

                                // Log the endIdx for debugging
                                logQueue.Add($"{i},{dateTime:MM/dd/yyyy HH:mm:ss.fff},EndIdx,,,{fullName},endIdx={endIdx}");

                                if (endIdx < 0)
                                {
                                    logQueue.Add($"{i},{dateTime:MM/dd/yyyy HH:mm:ss.fff},EndDelimiterNotFound,{email},{fullName},,End delimiter not found.");
                                    continue;
                                }
                                string narrowedString = contents.Substring(index, endIdx - index);

                                // Log the narrowed string for debugging
                                logQueue.Add($"{i},{dateTime:MM/dd/yyyy HH:mm:ss.fff},NarrowedString,,,{fullName},narrowedString=\"{narrowedString.Replace("\"", "\"\"")}\"");

                                Match phoneMatch = phoneRegex.Match(narrowedString);

                                string phoneNumber = phoneMatch.Success ? phoneMatch.Value : "No Number Found";

                                // Write contact line to queue
                                contactQueue.Add($"{fullName}|{phoneNumber} @{email}");

                                logQueue.Add($"{i},{dateTime:MM/dd/yyyy HH:mm:ss.fff},Success,{email},{fullName},{phoneNumber},Person Added");
                            }
                            catch (Exception ex)
                            {
                                logQueue.Add($",,,Error,,,{ex.Message}");
                            }
                        }
                        return localState;
                    },
                    localState => { }
                );

                // Signal completion and wait for the writers to finish
                logQueue.CompleteAdding();
                contactQueue.CompleteAdding();
                logWriterTask.Wait();
                contactWriterTask.Wait();

                // Sanitize ContactBook.txt to remove lines > 80 chars
                string contactBookPath = EnvDeclarations.contactBookFilePath;
                var sanitizedLines = File.ReadLines(contactBookPath)
                    .Where(line => line.Length <= 80)
                    .ToArray();
                File.WriteAllLines(contactBookPath, sanitizedLines);

                stopwatch.Stop();
                elapsed = stopwatch.Elapsed;
                MessageBox(0, $"Elapsed Time: {elapsed.TotalMilliseconds / 1000} s", "People Length", 0x00001000);
            }
            catch (Exception ex)
            {
                MessageBox(0, $",,,Fatal error,,,{ex.Message}", "Error", 0x00001000);
                throw;
            }
            /// <summary>
            /// Loads people from the contact book file using the path from <see cref="EnvDeclarations.contactBookFilePath"/>.
            /// </summary>
            people = ListGen.LoadPeopleFromContactBook(EnvDeclarations.contactBookFilePath);

            Random random = new Random();

            int randInt = random.Next(0, people.Count);
            MessageBox(0, $"Count: {people[randInt]}", "Regen List Array Parallel", 0x00001000);
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