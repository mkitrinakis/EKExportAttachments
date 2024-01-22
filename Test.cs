using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.SharePoint;

namespace ExportAttachments
{
   public class Test
    {
        public void run()
        {
            string rootFolder = "F:\\ExportProtocols\\";
            foreach (string fyear in new string[] { "2017", "2015" })
            {
                foreach (string filePath in Directory.EnumerateFiles(rootFolder + fyear, "*", SearchOption.AllDirectories))
                {

                    string[] parts = filePath.Replace(rootFolder, "").Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length.Equals(3))
                    {
                        string year = parts[0];
                        string protocolNo = parts[1];
                        string fileName = parts[2];
                        string url = ""; 
                        byte[] Content = System.IO.File.ReadAllBytes(filePath);
                        Console.WriteLine("OK");
                    }
                }
            }
        }


        public void run2()
        {
//            $SiteURL = ""
//$ListName = "Πρωτόκολλο Επιτροπής Ανταγωνισμού"
//$FieldName = "Καταθέτης"
//$ItemId = "5500"
            //using (SPSite site = new SPSite("https://easervice.epant.gr"))
            using (SPSite site = new SPSite("https://arxeio.epant.gr/sites/EpantProtocols/All/"))
              
              
            {
                using (SPWeb web = site.OpenWeb())
                {
                    //SPList l = web.Lists["Πρωτόκολλο Επιτροπής Ανταγωνισμού"];
                    SPList l = web.Lists["ΠΡΩΤΟΚΟΛΛΟ ΓΡΑΦΕΙΟΥ ΓΡΑΜΜΑΤΕΩΝ ΕΑ"];
                    //SPListItem itm = l.GetItemById(5500);
                    SPListItem itm = l.GetItemById(675);
                    Console.WriteLine("a:" + (itm["Καταθέτης"] ?? "").ToString());
                    Console.WriteLine(("b:" + itm["Katathetis2"] ?? "").ToString());
                    itm["Καταθέτης"] = "4017;#CRM Παντοπωλεια Κρήτης ΑΕ"; 
                    itm.SystemUpdate();
                    Console.WriteLine("a:" + (itm["Καταθέτης"] ?? "").ToString());
                    Console.WriteLine(("b:" + itm["Katathetis2"] ?? "").ToString());
                }

            }
        }
    }
}
