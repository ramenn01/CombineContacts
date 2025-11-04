using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace CombineContacts
{
    public class ContactObject
    {
        public ContactObject(ContactItem contact)
        {
            FirstName = contact.FirstName;
            MiddleName = contact.MiddleName;
            LastName = contact.LastName;
            Suffix = contact.Suffix;
            JobTitle = contact.JobTitle;
            CompanyName = contact.CompanyName;
            BusinessHomePage = contact.BusinessHomePage;
            BusinessTelephoneNumber = contact.BusinessTelephoneNumber;
            Business2TelephoneNumber = contact.Business2TelephoneNumber;
            Email1Address = contact.Email1Address;
            Email2Address = contact.Email2Address;
            Email3Address = contact.Email3Address;
            MobileTelephoneNumber = contact.MobileTelephoneNumber;
            Categories = contact.Categories;
            Body = contact.Body;
            ContactItem = contact;
        }

        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public string Suffix { get; set; }
        public string JobTitle { get; set; }
        public string CompanyName { get; set; }
        public string BusinessHomePage { get; set; }
        public string Categories { get; set; }
        public string Body { get; set; }
        public string BusinessTelephoneNumber { get; set; }
        public string Business2TelephoneNumber { get; set; }
        public string Email1Address { get; set; }
        public string Email2Address { get; set; }
        public string Email3Address { get; set; }
        public string MobileTelephoneNumber { get; set; }
        public ContactItem ContactItem { get; set; }

        public static bool HasCategories(ContactItem contact)
        {
            return string.IsNullOrWhiteSpace(contact.Categories);
        }

        public static int NumberOfFieldSet(ContactItem contact)
        {
            int count = 0;

            if (!string.IsNullOrWhiteSpace(contact.FirstName)) count++;
            if (!string.IsNullOrWhiteSpace(contact.MiddleName)) count++;
            if (!string.IsNullOrWhiteSpace(contact.LastName)) count++;
            if (!string.IsNullOrWhiteSpace(contact.Suffix)) count++;
            if (!string.IsNullOrWhiteSpace(contact.JobTitle)) count++;
            if (!string.IsNullOrWhiteSpace(contact.CompanyName)) count++;
            if (!string.IsNullOrWhiteSpace(contact.BusinessHomePage)) count++;
            if (!string.IsNullOrWhiteSpace(contact.Categories)) count++;
            if (!string.IsNullOrWhiteSpace(contact.Body)) count++;
            if (!string.IsNullOrWhiteSpace(contact.BusinessTelephoneNumber)) count++;
            if (!string.IsNullOrWhiteSpace(contact.Business2TelephoneNumber)) count++;
            if (!string.IsNullOrWhiteSpace(contact.Email1Address)) count++;
            if (!string.IsNullOrWhiteSpace(contact.Email2Address)) count++;
            if (!string.IsNullOrWhiteSpace(contact.Email3Address)) count++;
            if (!string.IsNullOrWhiteSpace(contact.MobileTelephoneNumber)) count++;

            return count;
        }

        public List<string> GenerateListOfCaterogies()
        {
            // Outlook stores categories as a comma-separated string
            if (string.IsNullOrWhiteSpace(Categories))
                return new List<string>();

            return Categories
                          .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                          .Select(c => c.Trim())
                          .Where(c => !string.IsNullOrEmpty(c))
                          .Distinct(StringComparer.OrdinalIgnoreCase)
                          .ToList();
        }

        public string CombineCategories(ContactItem target)
        {
            if (!string.IsNullOrWhiteSpace(Categories) == true)
            {
                return null;
            }

            // Get existing categories from both contacts
            string targetCategories = target.Categories ?? string.Empty;
            string sourceCategories = Categories ?? string.Empty;

            // Split by comma, trim spaces, and combine uniquely
            var combined = targetCategories.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                           .Select(c => c.Trim())
                                           .Union(sourceCategories.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                                  .Select(c => c.Trim()), StringComparer.OrdinalIgnoreCase)
                                           .ToList();

            // Join back into a single string
            string newCategories = string.Join(", ", combined);
                
            return newCategories;
        }

        public string GenerateBody(ContactItem currentContact)
        {
            var sb = new StringBuilder();

            sb.AppendLine("Merge Content Begin");

            if (FirstName != currentContact.FirstName && FirstName != null)
                sb.AppendLine($"FirstName: {FirstName}");
            if (MiddleName != currentContact.MiddleName && MiddleName != null)
                sb.AppendLine($"MiddleName: {MiddleName}");
            if (LastName != currentContact.LastName && LastName != null)
                sb.AppendLine($"LastName: {LastName}");
            if (Suffix != currentContact.Suffix && Suffix != null)
                sb.AppendLine($"Suffix: {Suffix}");
            if (JobTitle != currentContact.JobTitle && JobTitle != null)
                sb.AppendLine($"JobTitle: {JobTitle}");
            if (CompanyName != currentContact.CompanyName && CompanyName != null)
                sb.AppendLine($"CompanyName: {CompanyName}");
            if (BusinessHomePage != currentContact.BusinessHomePage && BusinessHomePage != null)
                sb.AppendLine($"BusinessHomePage: {BusinessHomePage}");
            // if (Categories != other.Categories)
            //    sb.AppendLine($"Categories: {Categories}");
            if (Body != currentContact.Body && Body != null)
                sb.AppendLine($"Body: {Body}");
            if (BusinessTelephoneNumber != currentContact.BusinessTelephoneNumber && BusinessTelephoneNumber != null)
                sb.AppendLine($"BusinessTelephoneNumber: {BusinessTelephoneNumber}");
            if (Business2TelephoneNumber != currentContact.Business2TelephoneNumber && Business2TelephoneNumber != null)
                sb.AppendLine($"Business2TelephoneNumber: {Business2TelephoneNumber}");
            if (Email1Address != currentContact.Email1Address && Email1Address != null)
                sb.AppendLine($"Email1Address: {Email1Address}");
            if (Email2Address != currentContact.Email2Address && Email2Address != null)
                sb.AppendLine($"Email2Address: {Email2Address}");
            if (Email3Address != currentContact.Email3Address && Email3Address != null)
                sb.AppendLine($"Email3Address: {Email3Address}");
            if (MobileTelephoneNumber != currentContact.MobileTelephoneNumber && MobileTelephoneNumber != null)
                sb.AppendLine($"MobileTelephoneNumber: {MobileTelephoneNumber}");

            sb.AppendLine("Merge Content End");

            return sb.ToString();
        }
    }

    public static class Logger
    {
        private static readonly string logFilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "app.log"
        );

        public static void Log(string message)
        {
            try
            {
                string logEntry = $"{message}";
                File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
            }
            catch (System.Exception ex)
            {
                // Optional: Write to console if logging fails
                Console.Error.WriteLine($"Logging failed: {ex.Message}");
            }
        }
    }

    internal class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            MAPIFolder contactsFolder = null;
            NameSpace ns = null;
            Application outlookApp = null;

            try
            {
                int count = 0;
                ContactItem currentItem = null;
                List<ContactObject> contactsToCombine = new List<ContactObject>();

                // Start Outlook application
                outlookApp = new Application();
                ns = outlookApp.GetNamespace("MAPI");
                ns.Logon("", "", Missing.Value, Missing.Value);
                contactsFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

                Items contactItems = contactsFolder.Items;
                contactItems.Sort("[FullName]", false);

                foreach (ContactItem contact in contactItems)
                {
                    count++;
                    Console.WriteLine(count);

                    Marshal.ReleaseComObject(currentItem);

                    /*
                    if(currentItem == null)
                    {
                        currentItem = contact;
                        continue;
                    }

                    if (currentItem.FullName == contact.FullName)
                    {
                        if (currentItem.FirstName == contact.FirstName &&
                            currentItem.MiddleName == contact.MiddleName &&
                            currentItem.LastName == contact.LastName &&
                            currentItem.Suffix == contact.Suffix &&
                            currentItem.JobTitle == contact.JobTitle &&
                            currentItem.CompanyName == contact.CompanyName &&
                            currentItem.BusinessHomePage == contact.BusinessHomePage &&
                            // currentItem.Categories == contact.Categories &&
                            // currentItem.Body == contact.Body &&
                            currentItem.BusinessTelephoneNumber == contact.BusinessTelephoneNumber &&
                            currentItem.Business2TelephoneNumber == contact.Business2TelephoneNumber &&
                            currentItem.Email1Address == contact.Email1Address &&
                            currentItem.Email2Address == contact.Email2Address &&
                            currentItem.Email3Address == contact.Email3Address &&
                            currentItem.MobileTelephoneNumber == contact.MobileTelephoneNumber)
                        {
                            bool bSave = false;

                            if (contact.Body != null)
                            {
                                string newBody = currentItem.Body;
                                newBody = contact.Body + $"\n{contact.Body}";
                                currentItem.Body = newBody;
                                bSave = true;
                            }

                            if (contact.Categories != null)
                            {
                                string existingCategoriesNew = contact.Categories ?? string.Empty;
                                string existingCategoriesOld = currentItem.Categories ?? string.Empty;

                                // Split into a list
                                var categoriesNew = existingCategoriesNew
                                    .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                    .Select(c => c.Trim())
                                    .ToList();

                                var categoriesOld = existingCategoriesOld
                                    .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                    .Select(c => c.Trim())
                                    .ToList();


                                foreach (string c in categoriesNew)
                                {
                                    if (!categoriesOld.Contains(c, StringComparer.OrdinalIgnoreCase))
                                    {
                                        categoriesOld.Add(c);
                                    }
                                }

                                // Join back into a semicolon-delimited string
                                currentItem.Categories = string.Join(";", categoriesOld);

                                bSave = true;
                            }

                            // Save the contact
                            if (bSave)
                            {
                                currentItem.Save();
                            }

                            string fullName = contact.FullName;
                            contact.Delete();
                            Console.WriteLine($"Delete {fullName}");

                            Marshal.ReleaseComObject(contact);
                            continue;
                        }
                        else
                        {
                            if (ContactObject.NumberOfFieldSet(currentItem) >= ContactObject.NumberOfFieldSet(contact))
                            {
                                ContactObject co = new ContactObject(contact);
                                contactsToCombine.Add(co);
                                // Marshal.ReleaseComObject(contact);
                            }
                            else
                            {
                                ContactObject co = new ContactObject(currentItem);
                                contactsToCombine.Add(co);
                                // Marshal.ReleaseComObject(currentItem);
                                currentItem = contact;
                            }
                            continue;
                        }
                    }
                    else
                    {
                        bool bSave = false;

                        if (contactsToCombine.Count > 0)
                        {
                            StringBuilder newBody = new StringBuilder();
                            newBody.AppendLine(currentItem.Body);
                            Dictionary<string, string> categoryMap = new Dictionary<string, string>();

                            foreach (ContactObject co in contactsToCombine)
                            {
                                string generateBoday = co.GenerateBody(currentItem);
                                if (string.IsNullOrWhiteSpace(generateBoday) == false)
                                {
                                    newBody.AppendLine(generateBoday);
                                }

                                List<string> newCategories = co.GenerateListOfCaterogies();

                                foreach (string category in newCategories)
                                {
                                    if (categoryMap.ContainsKey(category) == false)
                                    {
                                        categoryMap.Add(category, category);
                                    }
                                }
                            }

                            string newBodyStr = newBody.ToString();

                            if (string.IsNullOrWhiteSpace(newBodyStr) == false)
                            {
                                currentItem.Body = newBody.ToString();
                                bSave = true;

                            }

                            if (categoryMap.Count > 0)
                            {
                                if (currentItem.Categories != null)
                                {
                                    List<string> currentCategories = currentItem.Categories
                                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(c => c.Trim())
                                        .Where(c => !string.IsNullOrEmpty(c))
                                        .Distinct(StringComparer.OrdinalIgnoreCase)
                                        .ToList();

                                    foreach (string category in currentCategories)
                                    {
                                        if (categoryMap.ContainsKey(category) == false)
                                        {
                                            categoryMap.Add(category, category);
                                        }
                                    }
                                }

                                List<string> cl = categoryMap.Keys.ToList();
                                string newCategories = string.Join(", ", cl);
                                currentItem.Categories = newCategories;
                                bSave = true;
                            }

                            if (bSave)
                            {
                                currentItem.Save();
                            }

                            foreach (ContactObject co in contactsToCombine)
                            {
                                co.ContactItem.Delete();
                                Marshal.ReleaseComObject(co.ContactItem);
                            }

                            Logger.Log($"{currentItem.FullName}");
                        }

                        contactsToCombine.Clear();
                        Marshal.ReleaseComObject(currentItem);
                        currentItem = contact;
                        continue;
                    }*/
                } 
            
                ns.Logoff();    
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(contactsFolder);
                Marshal.ReleaseComObject(ns);
                Marshal.ReleaseComObject(outlookApp);
            }
        }
    }
}

