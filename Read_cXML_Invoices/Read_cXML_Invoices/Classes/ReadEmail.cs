using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.IO;
using System.Text;

namespace Read_cXML_Invoices.Classes
{
    public class ReadEmail
    {
        private ExchangeService service;
        private FolderId searchFolderId;
        private int numOfItemPull;
        private string dumpFolder;
        private List<string> files;
        private List<ItemId> emailIds;
        public List<EmailMessage> emailMsgs;

        public ReadEmail(string url, string userId, string password, string domain, string folderName, int maxItemNo, string fileFolder)
        {
            service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            service.Url = new Uri(url);
            service.Credentials = new NetworkCredential(userId, password, domain);
            searchFolderId = GetFolderId(service, folderName);
            numOfItemPull = maxItemNo;
            dumpFolder = fileFolder;

            files = new List<string>();
            emailIds = new List<ItemId>();
            GetEmails();
        }
        private void GetEmails()
        {
            ExtendedPropertyDefinition X_STATE = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings, "X-STATE", MapiPropertyType.String);
            SearchFilter newItemCheck = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
            FindItemsResults<Item> findResults = service.FindItems(searchFolderId, newItemCheck, new ItemView(numOfItemPull));
            emailMsgs = new List<EmailMessage>();

            for (int i = 0; i < findResults.Items.Count; i++)
            {
                EmailMessage message = EmailMessage.Bind(service, findResults.Items[i].Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments));
                message.Load();
                emailMsgs.Add(message);
            }
        }
        private FolderId GetFolderId(ExchangeService s, string folderName)
        {
            SearchFilter searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, folderName);
            FindFoldersResults findFolderResults = s.FindFolders(WellKnownFolderName.Inbox, searchFilter, new FolderView(10));
            return findFolderResults.Folders[0].Id;
        }
        public int EmailCount { get { return files.Count; } }
        public string GetFileName(int pos) { return files[pos]; }
        public void UpdateEmailStatus(int pos)
        {
            EmailMessage message = EmailMessage.Bind(service, emailIds[pos], new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.IsRead));
            message.IsRead = true;
            message.Update(ConflictResolutionMode.AlwaysOverwrite);
        }
        private string FileName(string folder, string emailFileName, int i)
        {
            string name = "", ext = "";

            ext = emailFileName.Substring(emailFileName.LastIndexOf('.'));
            name = emailFileName.Replace(ext, "") + "_" + i + ext;
            if (File.Exists(folder + @"\" + name))
            {
                i += 1;
                name = FileName(folder, emailFileName, i);
            }

            return name;
        }
    }
}
