using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using System.IO;
using System.Security;
using System.Net;

namespace ExportAttachments 
{
    class ImportAttsV2
    {
        // https://arxeio.epant.gr/sites/ProtocolsFiles2023/GeneralProtocol_2023/
        string listName = "GeneralProtocol_";
        string rootPath = "F:\\ExportProtocols\\";
        string url = "https://arxeio.epant.gr/sites/ProtocolsFilesXXXX/GeneralProtocol_XXXX/"; 
        public void run(string year)
        {
            Console.WriteLine("starting 'ImportAttachments' ...");
            using (SPSite site = new SPSite(url + "/ProtocolFiles" + year))
            {
                listName += year;
                string path = "C:\\ExportProtocols\\" + year;
                SPWeb web = site.OpenWeb();
                SPList oList = web.Lists[listName];
                foreach (string filePath in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories))
                {
                    if (uploadFileV2(web, oList, filePath, year))
                    {
                        Console.WriteLine(filePath + " --> OK");
                    }
                    else
                    {
                        Console.WriteLine(filePath + " --> NOTOK");
                    }

                }
            }
            }
        


        private bool uploadFileV2(SPWeb web, SPList list, string filePath, string year)
        {
        try
        {
            //    string[] parts = filePath.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            //    if (parts.Length.Equals(5))
            //    {
            //        string protocolNo = parts[3];
            //        string fileName = parts[4];

            //            byte[] Content = System.IO.File.ReadAllBytes(filePath);
            //        Console.WriteLine("OK");
            //        FileCreationInformation fileCreationInformation = new FileCreationInformation
            //        {

            //            Content = System.IO.File.ReadAllBytes(fileName),
            //            Url = System.IO.Path.Combine("GeneralProtocol_" + year + "/" + year + "-" + protocolNo.PadLeft(7, '0') + "/", System.IO.Path.GetFileName(fileName)),
            //            Overwrite = false
            //        };
            //        //List list = clientContext.Web.Lists.GetByTitle("Documents");
            //        Microsoft.SharePoint.Client.File uploadFile = list.RootFolder.Files.Add(fileCreationInformation);
            //        ListItem itm = uploadFile.ListItemAllFields;
            //        itm["Title"] =  fileName;
            //        itm.Update();
            //        clientContext.Load(uploadFile);
            //        clientContext.Load(itm);
            //        clientContext.ExecuteQuery();
            //        return true;
            //    }
            //    else { return false;  }
            return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error on uploading: " + filePath + " --> " + e.Message);
                return false;
            }

        }


       


        //private bool createFolder(ClientContext clientContext, List list, string folderName)
        //{
        //    try
        //    {

        //        var folder = list.RootFolder;
        //        clientContext.Load(folder);
        //        clientContext.ExecuteQuery();
        //        ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
        //        newItemInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
        //        newItemInfo.LeafName = folderName;
        //        ListItem newListItem = list.AddItem(newItemInfo);
        //        newListItem["Title"] = folderName;
        //        newListItem.Update();
        //        clientContext.ExecuteQuery();
        //        return true;
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine("Error on uploading: " + folderName + " --> " + e.Message);
        //        return false;
        //    }

        //}

        private void init(string url)
        {

            //Console.WriteLine("Site to process:" + url);

            //Console.WriteLine("Document Library to Process:" + listName + "  -Press Enter to keep the name, If to change please write the new name");
            //string tmp = Console.ReadLine();
            //if (!tmp.Trim().Equals(""))
            //{
            //    listName = tmp;
            //    Console.WriteLine("Document Library to Process changed to:" + listName);
            //}

            //Console.WriteLine("Folder to import will be in:" + filePath + "  -Press Enter to keep the name, If to change please write the new name (including the trailing \\)");
            //tmp = Console.ReadLine();
            //if (!tmp.Trim().Equals(""))
            //{
            //    filePath = tmp;
            //    Console.WriteLine("File Path to Extract to changed to:" + filePath);
            //}

         
        }




    }
}
