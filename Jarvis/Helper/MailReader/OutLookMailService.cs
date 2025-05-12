using System;
using System.Collections.Generic;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Jarvis.Helper.MailReader
{
    public class OutlookDesktopReader
    {
        private Outlook.Application _outlookApp;
        private Outlook.NameSpace _outlookNamespace;

        public OutlookDesktopReader()
        {
            _outlookApp = new Outlook.Application();
            _outlookNamespace = _outlookApp.GetNamespace("MAPI");
            _outlookNamespace.Logon("", "", Missing.Value, Missing.Value);
        }

        public List<OutlookMailItem> GetUnreadOrderMailsFromRoboFolder()
        {
            var result = new List<OutlookMailItem>();

            Outlook.MAPIFolder inboxFolder = _outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.MAPIFolder roboFolder = null;

            // Look for subfolder "Robo"
            foreach (Outlook.MAPIFolder folder in inboxFolder.Folders)
            {
                if (folder.Name.Equals("Robo", StringComparison.OrdinalIgnoreCase))
                {
                    roboFolder = folder;
                    break;
                }
            }
            if (roboFolder == null)
                throw new Exception("Robo folder not found.");

            Outlook.Items mailItems = roboFolder.Items;
            mailItems = mailItems.Restrict("[Unread]=true");

            foreach (Outlook.MailItem item in mailItems)
            {
                if (item.Subject != null && item.Subject.StartsWith("Order:", StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(new OutlookMailItem
                    {
                        Subject = item.Subject,
                        Body = item.Body,
                        ReceivedTime = item.ReceivedTime,
                        SenderEmail = item.SenderEmailAddress
                    });

                    // Mark as read (optional)
                    item.UnRead = false;
                    item.Save();
                }
            }
            return result;
        }
    }

    public class OutlookMailItem
    {
        public string Subject { get; set; }
        public string Body { get; set; }
        public DateTime ReceivedTime { get; set; }
        public string SenderEmail { get; set; }
    }
}
