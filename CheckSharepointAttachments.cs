using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using System.IO;




namespace ExportAttachments
{
    class CheckSharepointAttachments
    {
        SPSite site;
        SPWeb web;
        int idStart = 0; 
        public void run()
        {
            using (site = new SPSite("https://easervice.epant.gr"))
            {
                using (web = site.OpenWeb())
                {
                    string no = Console.ReadLine();
                    getQuery(no); 
                    //processDocLib();
                    //processGeneralProtocolDocs();
                    //processCaseDocuments();
                }
            }
        }


        private void getQuery (string protocolNo)
        {
            SPList l1 = web.Lists["Γενική Βιβλιοθήκη Εγγράφων"];
            SPQuery qry1 = new SPQuery();
            qry1.ViewFields = "<FieldRef Name='ID'/><FieldRef Name='ProtocolLookup'/>";
            qry1.Query = "<Where><Eq><FieldRef Name='ProtocolLookup' LookupID='TRUE' /><Value Type='Lookup'> " + protocolNo + "</Value></Eq></Where>";
            qry1.ViewAttributes = "Scope='Recursive'";
            SPListItemCollection col1 = l1.GetItems(qry1);
            Console.WriteLine("Γενική Βιβλιοθήκη Εγγράφων" + col1.Count);

            SPList l2 = web.Lists["Επισυναπτόμενα Έγγραφα Γενικού Πρωτοκόλου"];
            SPQuery qry2 = new SPQuery();
            qry2.ViewFields = "<FieldRef Name='ID'/><FieldRef Name='GeneralProtocolID'/>";
            qry2.Query = "<Where><Eq><FieldRef Name='GeneralProtocolID' LookupID='TRUE'/><Value Type='Lookup' > " + protocolNo + "</Value></Eq></Where>";
            qry2.ViewAttributes = "Scope='Recursive'";
            SPListItemCollection col2 = l2.GetItems(qry2);
            Console.WriteLine("Επισυναπτόμενα Έγγραφα Γενικού Πρωτοκόλου" + col2.Count);



            SPList l3 = web.Lists["Βιβλιοθήκη Υποθέσεων"];
            SPQuery qry3 = new SPQuery();
            qry3.ViewFields = "<FieldRef Name='ID'/><FieldRef Name='ProtocolLookup' />";
            qry3.Query = "<Where><Eq><FieldRef Name='ProtocolLookup' LookupID='TRUE'/><Value Type='Lookup' > " + protocolNo + "</Value></Eq></Where>";
            qry3.ViewAttributes = "Scope='Recursive'";
            SPListItemCollection col3 = l3.GetItems(qry3);
            Console.WriteLine("Βιβλιοθήκη Υποθέσεων" + col3.Count);
        }

        private SPListItemCollection getColFormList(SPList l)
        {
            SPQuery qry = new SPQuery();
            qry.Query = "<Where><Geq><FieldRef Name='ID'/><Value Type='Number'>" + idStart + "</Value></Geq></Where><OrderBy><FieldRef Name='ID'/></OrderBy>";
            qry.ViewFields = "<FieldRef Name='ID'/>";
            qry.ViewAttributes = "Scope='Recursive'";
            SPListItemCollection col = l.GetItems(qry);
            return col; 
        }

        private void processDocLib()
        {
            SPList l = web.Lists["Γενική Βιβλιοθήκη Εγγράφων"];
            SPListItemCollection col = getColFormList(l); 
            Console.WriteLine((l.GetItemById(col[0].ID)["ProtocolLookup"] ?? "-").ToString());
        }

        private void processGeneralProtocolDocs()
        {
            SPList l = web.Lists["Επισυναπτόμενα Έγγραφα Γενικού Πρωτοκόλου"];
            SPListItemCollection col = getColFormList(l);
            Console.WriteLine((l.GetItemById(col[0].ID)["GeneralProtocolID"] ?? "-").ToString());
            
        }

        private void processCaseDocuments()
        {
            SPList l = web.Lists["Βιβλιοθήκη Υποθέσεων"];
            SPListItemCollection col = getColFormList(l);
            foreach (SPListItem itm in col)
            {
                string rs = (l.GetItemById(itm.ID)["ProtocolLookup"] ?? "-").ToString(); 

                Console.WriteLine(rs);

            }
        }
    }
}
