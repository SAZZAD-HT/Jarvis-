using System;
using System.Collections.Generic;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMailReader
{
    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Connecting to Outlook...");
                // Initialize Outlook Application
                var outlookApp = new Outlook.Application();
                var outlookNamespace = outlookApp.GetNamespace("MAPI");

                // Try to log on to Outlook (this may throw a security exception)
                outlookNamespace.Logon("", "", Missing.Value, Missing.Value);

                Console.WriteLine("Successfully connected to Outlook.");

                // Step 2: List all folders under Inbox
                Outlook.MAPIFolder inbox = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Console.WriteLine("\nAvailable folders under Inbox:");

                bool folderFound = false;
                foreach (Outlook.MAPIFolder folder in inbox.Folders)
                {
                    Console.WriteLine($"- {folder.Name}");

                    // Check if the folder exists
                    if (folder.Name.Equals("Robo", StringComparison.OrdinalIgnoreCase))
                    {
                        folderFound = true;
                    }
                }

                if (!folderFound)
                {
                    Console.WriteLine("Robo folder not found!");
                    return;
                }

                Console.WriteLine("\nEnter the name of the folder you want to read emails from (case-sensitive):");
                string targetFolderName = Console.ReadLine();

                Outlook.MAPIFolder targetFolder = null;
                foreach (Outlook.MAPIFolder folder in inbox.Folders)
                {
                    if (folder.Name.Equals(targetFolderName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetFolder = folder;
                        break;
                    }
                }

                if (targetFolder == null)
                {
                    Console.WriteLine($"Folder '{targetFolderName}' not found inside Inbox!");
                    return;
                }

                // Step 3: Get unread mails
                Console.WriteLine($"Looking for unread emails in '{targetFolderName}'...");

                Outlook.Items unreadItems = targetFolder.Items.Restrict("[Unread]=true");

                List<MailRecord> mails = new List<MailRecord>();

                foreach (object obj in unreadItems)
                {
                    if (obj is Outlook.MailItem mail)
                    {
                        if (mail.Subject != null && mail.Subject.StartsWith("Order:", StringComparison.OrdinalIgnoreCase))
                        {
                            mails.Add(new MailRecord
                            {
                                Subject = mail.Subject,
                                SenderEmail = mail.SenderEmailAddress,
                                ReceivedTime = mail.ReceivedTime,
                                Body = mail.Body
                            });

                            // Optional: Mark the mail as read
                            mail.UnRead = false;
                            mail.Save();
                        }
                    }
                }

                // Step 4: Print Mails in Console
                Console.WriteLine($"\nFound {mails.Count} mail(s) matching criteria.\n");

                foreach (var mail in mails)
                {
                    Console.WriteLine("-----------------------------------");
                    Console.WriteLine($"Subject      : {mail.Subject}");
                    Console.WriteLine($"Sender Email : {mail.SenderEmail}");
                    Console.WriteLine($"ReceivedTime : {mail.ReceivedTime}");
                    Console.WriteLine($"Body         : {(string.IsNullOrEmpty(mail.Body) ? "(Empty)" : mail.Body.Substring(0, Math.Min(mail.Body.Length, 100)))}...");
                    Console.WriteLine("-----------------------------------\n");
                }

                Console.WriteLine("Completed reading mails. Press any key to exit...");
                Console.ReadKey();
            }
            catch (System.UnauthorizedAccessException ex)
            {
                Console.WriteLine("Permission Error: The application does not have the required permissions to access Outlook.");
                Console.WriteLine($"Error Details: {ex.Message}");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine("COM Error: Failed to interact with Outlook.");
                Console.WriteLine($"Error Details: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred.");
                Console.WriteLine($"Error Details: {ex.Message}");
            }
        }
    }

    public class MailRecord
    {
        public string Subject { get; set; }
        public string SenderEmail { get; set; }
        public DateTime ReceivedTime { get; set; }
        public string Body { get; set; }
    }
}
