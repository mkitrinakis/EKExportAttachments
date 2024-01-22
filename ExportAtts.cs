using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using System.IO;

namespace ExportAttachments 
{
    class ExportAtts
    {
        int idStart = 1;
        // string listName = "Documents";
        string listName = "Γενική Βιβλιοθήκη Εγγράφων";
        string filePath = "F:\\ExportProtocol\\";
        string folderFieldName = "Γενικό Πρωτόκολλο";
        string alternateField = "Ημ/νία Αρχείου"; 
        public void run(string url)
        {
            init(url);
            using (SPSite site = new SPSite(url))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPList l = web.Lists[listName];
                    // SPList l = web.Lists["F416"];
                    process(web, l);
                }
            }
        }

        private void process(SPWeb web, SPList l)
        {
            SPQuery qry = new SPQuery();
            qry.Query = "<Where><Geq><FieldRef Name='ID'/><Value Type='Number'>" + idStart + "</Value></Geq></Where><OrderBy><FieldRef Name='ID'/></OrderBy>";

            qry.ViewFields = "<FieldRef Name='ID'/>";
            qry.ViewAttributes = "Scope='Recursive'";
            SPListItemCollection col = l.GetItems(qry);
            foreach (SPListItem itm in col)
            {
                processItem(web, l, itm);

            }
        }


        //private void processItem(SPWeb web, SPListItem itm)
        //{

        //    SPFolder folder = web.GetFolder(web.Url + "/Documents/Attachments/" + itm.ID + "/");
        //    SPFileCollection fileCol = folder.Files;
        //    foreach (SPFile file in fileCol)
        //    {
        //        string fileName = file.Name;
        //        Stream stream = file.OpenBinaryStream();
        //        using (var fileStream = new System.IO.FileStream(fileName, FileMode.Create))
        //        {
        //            stream.CopyTo(fileStream);
        //        }
        //    }
        //}

        private void init(string url)
        {

            Console.WriteLine("Site to process:" + url);

            Console.WriteLine("Document Library to Process:" + listName + "  -Press Enter to keep the name, If to change please write the new name");
            //string tmp = Console.ReadLine();
            //if (!tmp.Trim().Equals(""))
            //{
            //    listName = tmp;
            //    Console.WriteLine("Document Library to Process changed to:" + listName);
            //}

            Console.WriteLine("Folder to extract will be in:" + filePath + "  -Press Enter to keep the name, If to change please write the new name (including the trailing \\)");
            //tmp = Console.ReadLine();
            //if (!tmp.Trim().Equals(""))
            //{
            //    filePath = tmp;
            //    Console.WriteLine("File Path to Extract to changed to:" + filePath);
            //}


            //Console.WriteLine("If you want to extract documents in subfolders based on a field type the name of the field");
            //tmp = Console.ReadLine();
            //if (!tmp.Trim().Equals(""))
            //{
            //    folderFieldName = tmp;
            //    Console.WriteLine("Attachments will be extracted in subfolder named as the field:" + folderFieldName);
            //}


            Console.WriteLine("Starting from ID=1. If previously partially run and want to start from the middle write the ID to start ");
            string tmp = Console.ReadLine();
            if (!tmp.Trim().Equals(""))
            {
                idStart = Convert.ToInt32(tmp.Trim());
                Console.WriteLine("ID changed to :" + idStart);
            }

        }

        private void processItem(SPWeb web, SPList l, SPListItem mitm)
        {
            string fileFinalPath = filePath;
            SPListItem itm = l.GetItemById(mitm.ID);
            try
            {
                //string url = (itm["FileDirRef"] ?? "").ToString();
                string fileUrl = itm.Url;
                string[] urls = fileUrl.Split('/');
                string title = urls[urls.Length - 1];
                //url = itm.Url;
                // string url = itm.Folder.Url;
                // string fileUrl = web.Url + "/Documents/" + itm.Title;
                //string fileUrl = url + "/" + itm.Title;

                if (!folderFieldName.Equals(""))
                {
                    string subFolder = (itm[folderFieldName] ?? "").ToString().Trim();
                    if (!subFolder.Trim().Equals(""))
                    {
                        if (subFolder.Contains(";#"))
                        {
                            subFolder = subFolder.Split(new string[] { ";#" }, StringSplitOptions.RemoveEmptyEntries)[1];
                        }
                    }
                    else
                    {
                      
                        try
                        {
                            subFolder = ((DateTime)itm[alternateField]).Year.ToString(); 
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(itm.ID + " cannot create folder:" + e.Message);
                            subFolder = ""; 
                        }
                    }
                    if (!subFolder.Trim().Equals(""))
                    {
                        fileFinalPath += subFolder + "/";
                        if (!Directory.Exists(fileFinalPath))
                        {
                            Directory.CreateDirectory(fileFinalPath);
                        }
                    }
                }

                SPFile file = web.GetFile(fileUrl);
                Stream stream = file.OpenBinaryStream();
                
                using (var fileStream = new System.IO.FileStream(fileFinalPath + title, FileMode.Create))
                {
                    stream.CopyTo(fileStream);
                }
                stream.Close();
                Console.WriteLine(itm.ID);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error on item with id: " + itm.ID + " [" + (fileFinalPath ?? "").ToString() + (itm.Title ?? "").ToString() + "] --> " + e.Message);
            }
        }

    }
}
