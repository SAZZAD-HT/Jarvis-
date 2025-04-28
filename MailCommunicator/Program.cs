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
            Console.WriteLine("Connecting to Outlook...");
            Console.Out.Flush();

            // Step 1: Initialize Outlook Application
            var outlookApp = new Outlook.Application();
            var outlookNamespace = outlookApp.GetNamespace("MAPI");
            outlookNamespace.Logon("", "", Missing.Value, Missing.Value);

            // Step 2: List all folders under Inbox
            Outlook.MAPIFolder inbox = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Console.WriteLine("\nAvailable folders under Inbox:");
            foreach (Outlook.MAPIFolder folder in inbox.Folders)
            {
                Console.WriteLine($"- {folder.Name}");
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

            // Step 4: Print Mails in Console (or you can pass to another method)
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
    }

    public class MailRecord
    {
        public string Subject { get; set; }
        public string SenderEmail { get; set; }
        public DateTime ReceivedTime { get; set; }
        public string Body { get; set; }
    }
}
