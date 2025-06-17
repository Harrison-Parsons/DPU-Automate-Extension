using System;
using System.Collections.Generic;
using System.IO;

namespace UtilitiesAutomateExtension.Pages
{
    /// <summary>
    /// Utility class for loading people from the contact book file.
    /// </summary>
    internal class ListGen
    {
        /// <summary>
        /// Loads people from the specified contact book file.
        /// </summary>
        /// <param name="contactBookPath">The path to the contact book file.</param>
        /// <returns>A list of <see cref="Person"/> objects.</returns>
        public static List<Person> LoadPeopleFromContactBook(string contactBookPath)
        {
            var people = new List<Person>();

            if (!File.Exists(contactBookPath))
                return people;

            foreach (var line in File.ReadLines(contactBookPath))
            {
                // Expected format: FullName|PhoneNumber @Email
                var parts = line.Split('|', 2);
                if (parts.Length != 2)
                    continue;

                var name = parts[0].Trim();
                if (string.IsNullOrWhiteSpace(name))
                    continue;

                var phoneAndEmail = parts[1].Split('@', 2);
                if (phoneAndEmail.Length != 2)
                    continue;

                var phone = phoneAndEmail[0].Trim();
                var email = phoneAndEmail[1].Trim();
                if (string.IsNullOrWhiteSpace(phone) || string.IsNullOrWhiteSpace(email))
                    continue;

                // Re-add the '@' to the email
                people.Add(new Person(name, "@" + email, phone));
            }

            return people;
        }
    }
}
