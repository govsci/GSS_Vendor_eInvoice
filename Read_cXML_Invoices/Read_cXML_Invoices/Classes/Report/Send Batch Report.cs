using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Read_cXML_Invoices.Classes.Report
{
    public class Send_Batch_Report
    {
        private string folder = "";
        private List<DuplicateOrderID> duplicateOrderIDs;
        public Send_Batch_Report(string invoiceDropFolder)
        {
            try
            {
                folder = invoiceDropFolder;
                duplicateOrderIDs = new List<DuplicateOrderID>();
                CheckFolders();
                CheckList();
            }
            catch(Exception ex)
            {
                Constants.ERRORS.Add(new Objects.Error(ex, "Send_Batch_Report", "Send_Batch_Report"));
            }
        }

        private void CheckMasterFolder() //TEST ONLY
        {
            string path = folder + "\\MASTER_COPY_ONLY\\";
            string[] files = Directory.GetFiles(path);
            for (int i = 0; i < files.Length; i++)
            {
                if (!files[i].EndsWith("Thumbs.db"))
                {
                    string[] fileNameS = files[i].Split('\\');
                    string fileName = fileNameS[fileNameS.Length - 1];
                    string orderID = "";

                    try
                    {
                        string[] pdfNameS = fileName.Split('.');

                        if (pdfNameS.Length == 4)
                            orderID = pdfNameS[1];
                        else
                            orderID = pdfNameS[1] + "." + pdfNameS[2];
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(files[i] + "\n\n" + ex.ToString());
                        string r = Console.ReadLine();
                    }

                    if (CheckFolder(path, files[i], orderID))
                    {
                        string s = "MASTER_COPY_ONLY";
                        DuplicateOrderID dup = duplicateOrderIDs.Find(d => d.Ship == s && d.OrderID == orderID);

                        if (dup == null)
                            duplicateOrderIDs.Add(new DuplicateOrderID(
                                s
                                , ""
                                , ""
                                , orderID
                                ));
                    }
                }
            }
        }

        private void CheckFolders()
        {
            List<string> ships = Directory.GetDirectories(folder).ToList();
            foreach(string ship in ships)
            {
                List<string> batches = Directory.GetDirectories(ship).ToList();
                foreach(string batch in batches)
                {
                    List<string> vendors = Directory.GetDirectories(batch).ToList();
                    foreach(string vendor in vendors)
                    {
                        string[] files = Directory.GetFiles(vendor);
                        for(int i = 0; i < files.Length; i++)
                        {
                            if (!files[i].EndsWith("Thumbs.db"))
                            {
                                string[] fileNameS = files[i].Split('\\');
                                string fileName = fileNameS[fileNameS.Length - 1];
                                string orderID = "";

                                try
                                {
                                    string[] pdfNameS = fileName.Split('.');

                                    if (pdfNameS.Length == 4)
                                        orderID = pdfNameS[1];
                                    else
                                        orderID = pdfNameS[1] + "." + pdfNameS[2];
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(files[i] + "\n\n" + ex.ToString());
                                    string r = Console.ReadLine();
                                }

                                if (CheckFolder(vendor, files[i], orderID))
                                {
                                    string s = ship.Split('\\')[ship.Split('\\').Length - 1];
                                    string b = batch.Split('\\')[batch.Split('\\').Length - 1];
                                    string v = vendor.Split('\\')[vendor.Split('\\').Length - 1];

                                    DuplicateOrderID dup = duplicateOrderIDs.Find(d => d.Ship == s && d.Batch == b && d.Vendor == v && d.OrderID == orderID);

                                    if (dup == null)
                                        duplicateOrderIDs.Add(new DuplicateOrderID(
                                            s
                                            , b
                                            , v
                                            , orderID
                                            ));
                                }
                            }
                        }
                    }
                }
            }
        }

        private bool CheckFolder(string path, string fileName, string orderID)
        {
            if (!fileName.EndsWith(".xls") && !fileName.EndsWith(".xlsx"))
            {
                string[] files = Directory.GetFiles(path);
                for (int i = 0; i < files.Length; i++)
                {
                    if (files[i] != fileName && files[i].Contains(orderID))
                        return true;
                }
            }
            return false;
        }

        private void CheckList()
        {
            if (duplicateOrderIDs.Count > 0)
            {
                string table = "<table border='1'><tbody><tr><th>Ship Type</th><th>Batch</th><th>Vendor</th><th>Order ID</th></tr>";
                foreach(DuplicateOrderID order in duplicateOrderIDs)
                    table += "<tr><td>" + order.Ship + "</td><td>" + order.Batch + "</td><td>" + order.Vendor + "</td><td>" + order.OrderID + "</td></tr>";
                table += "</tbody></table>";

                string msg = "The following batches contain multiple invoices with the same respective Order No.<br><br>" + table;
                Email.SendEmail(msg, "Invoices Received - Duplicated Order No.s", "", Constants.EmailRecipients, "", "", "", true);
            }
            else
                Email.SendEmail("Order IDs were not duplicated in the batches", "Invoices Received - No Duplications", "", Constants.EmailRecipients, "", "", "", false);
        }

        private class DuplicateOrderID
        {
            public DuplicateOrderID(string ship, string batch, string vendor, string orderID)
            {
                Ship = ship;
                Batch = batch;
                Vendor = vendor;
                OrderID = orderID;
            }

            public string Ship { get; }
            public string Batch { get; }
            public string Vendor { get; }
            public string OrderID { get; }
        }
    }
}
