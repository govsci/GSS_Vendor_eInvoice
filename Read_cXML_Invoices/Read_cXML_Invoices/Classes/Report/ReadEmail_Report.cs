using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;
using System.Net;

namespace Read_cXML_Invoices.Classes.Report
{
    public class ReadEmail_Report
    {
        private ExchangeService service;
        private FolderId searchFolderId;
        private int numOfItemPull;
        private List<string> files;
        private List<ItemId> emailIds;
        public List<EmailMessage> emailMsgs;

        public ReadEmail_Report(string url, string userId, string password, string domain, string folderName, int maxItemNo, DateTime date)
        {
            try
            {
                service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                service.Url = new Uri(url);
                service.Credentials = new NetworkCredential(userId, password, domain);
                searchFolderId = GetFolderId(service, folderName);
                numOfItemPull = maxItemNo;

                files = new List<string>();
                emailIds = new List<ItemId>();
                GetEmails(date);
            }
            catch (Exception ex) 
            {
                Constants.ERRORS.Add(new Objects.Error(ex, "ReadEmail", "ReadEmail"));
            }
        }
        private void GetEmails(DateTime date)
        {
            ExtendedPropertyDefinition X_STATE = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings, "X-STATE", MapiPropertyType.String);
            SearchFilter filter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsGreaterThanOrEqualTo(EmailMessageSchema.DateTimeReceived, date.Date));
            FindItemsResults<Item> findResults = service.FindItems(searchFolderId, filter, new ItemView(numOfItemPull));
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

        private void CheckDirectory(string dir)
        {
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
        }
    }
}
