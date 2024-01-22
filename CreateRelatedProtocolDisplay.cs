using System;
using System.Collections.Generic;
using Microsoft.SharePoint;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportAttachments 
{
    class CreateRelatedProtocolDisplay
    {
        string listName = "Γενικό Πρωτόκολλο";
        string viewName = "toProcess";



        public void run(string url)
        {
            try
            {
                init(url);
                using (SPSite site = new SPSite(url))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPList list = web.Lists[listName];
                        SPView v = list.Views[viewName];
                        SPListItemCollection col = list.GetItems(v);
                        Console.WriteLine("Documents to are process:" + col.Count + " -Press any key to continue, Ctrl + C to terminate");
                        Console.ReadLine();
                        foreach (SPListItem itm in list.GetItems(v))
                        {
                            processItem(list, itm);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR:" + e.Message);
            }
        }

        private void init(string url)
        {

            Console.WriteLine("Site to process:" + url);

            Console.WriteLine("List to Process:" + listName + "  -Press Enter to keep the name, If to change please write the new name");
            string tmp = Console.ReadLine();
            if (!tmp.Trim().Equals(""))
            {
                listName = tmp;
                Console.WriteLine("List to Process changed to:" + listName);
            }

            Console.WriteLine("View to Process:" + viewName + "  -Press Enter to keep the name, If to change please write the new name");
            tmp = Console.ReadLine();
            if (!tmp.Trim().Equals(""))
            {
                viewName = tmp;
                Console.WriteLine("View to Process changed to:" + viewName);
            }
            Console.WriteLine("Press any key to continue, Ctrl+C to terminate");
            Console.ReadLine();

        }



        void processItem(SPList l, SPListItem itm)
        {
            SPQuery qry = new SPQuery();
            qry.Query = "<Where><Eq><FieldRef Name='ID'/><Value Type='Number'>" + itm.ID + "</Value></Eq></Where>";
            //  qry.ViewFields = "<FieldRef Name='ID'/><FieldRef Name='ID'/><FieldRef Name='ID'/><FieldRef Name='ID'/><FieldRef Name='ID'/><FieldRef Name='ID'/><FieldRef Name='ID'/><FieldRef Name='ID'/><FieldRef Name='ID'/><FieldRef Name='ID'/><FieldRef Name='ID'/>"
            qry.ViewFields = "<FieldRef Name='ID'/><FieldRef Name='RelatedProtocol'/><FieldRef Name='RelatedProtocol_Display'/><FieldRef Name='Sender'/><FieldRef Name='Sender_Display'/><FieldRef Name='Sender2'/><FieldRef Name='Sender2_Display'/><FieldRef Name='Receiver'/><FieldRef Name='Receiver_Display'/><FieldRef Name='Receiver2'/><FieldRef Name='Receiver2_Display'/>";
            SPListItemCollection col = l.GetItems(qry);
            if (col.Count > 0)
            {
                SPListItem fillItem = col[0];
                bool toUpdate = false;
                bool rs = processField(fillItem, "RelatedProtocol");
                toUpdate = toUpdate || rs;
                rs = processField(fillItem, "Sender");
                toUpdate = toUpdate || rs;
                rs = processField(fillItem, "Sender2");
                toUpdate = toUpdate || rs;
                rs = processField(fillItem, "Receiver");
                toUpdate = toUpdate || rs;
                rs = processField(fillItem, "Receiver2");
                toUpdate = toUpdate || rs;
                if (toUpdate) { Console.WriteLine("updating"); fillItem.SystemUpdate(); }
            }
        }

        bool processField(SPListItem itm, string fieldName)
        {
            bool toUpdate = false;
            try
            {
                string tmp = (itm[fieldName] ?? "").ToString();
                Console.WriteLine(itm.ID + " --> " + tmp);
                if (tmp.Contains(";#"))
                {
                    string[] arr = tmp.Split(new string[] { ";#" }, StringSplitOptions.RemoveEmptyEntries);
                    int i = 0;
                    List<string> rsArr = new List<string>();
                    foreach (string term in arr)
                    {
                        if ((i % 2).Equals(1))
                        {
                            rsArr.Add(term);
                        }
                        i++;
                    }
                    string rs = String.Join(", ", rsArr);
                    Console.WriteLine("to " + rs);
                    if (!rs.Trim().Equals(""))
                    {
                        itm[fieldName + "_Display"] = rs;
                        toUpdate = true;
                    }
                }
                else
                {
                    if (!tmp.Trim().Equals(""))
                    {
                        toUpdate = true;
                        itm[fieldName + "_Display"] = tmp;
                    }
                }
                return toUpdate;
            }
            catch (Exception e)
            {
                Console.WriteLine(itm.ID + " --> Error:" + e.Message);
                return toUpdate;
            }
        }

    }
}
