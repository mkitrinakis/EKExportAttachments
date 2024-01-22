using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.IO;
using System.Security;
using System.Net;

namespace ExportAttachments 
{
    class ImportAtts
    {

        string listName = "GeneralProtocol_XXXX";
        string rootPath = "F:\\ExportProtocol\\XXXX";
        string url = "https://arxeio.epant.gr/sites/ProtocolsFilesXXXX/";

        public void run(string year)
        {
            Console.WriteLine("starting 'ImportAttachments' ...");
            using (ClientContext clientContext = new ClientContext(url.Replace("XXXX", year)))
            {
                listName = listName.Replace("XXXX", year);
                rootPath = rootPath.Replace("XXXX", year); 
                //SecureString passWord = new SecureString();
                //foreach (char c in "W0rkInf_portal!63".ToCharArray()) passWord.AppendChar(c);
                clientContext.Credentials = new NetworkCredential("SP_install", "spinstallP@ssw0rd", "epant");
                Web web = clientContext.Web;
                clientContext.Load(web);
                clientContext.ExecuteQuery();
                CamlQuery query = new CamlQuery();
                query.ViewXml = @"";
                List oList = clientContext.Web.Lists.GetByTitle(listName);
                clientContext.Load(oList);
                clientContext.ExecuteQuery();
                foreach (string filePath in Directory.EnumerateFiles(rootPath, "*", SearchOption.AllDirectories))
                {
                    Console.WriteLine("importing: " + filePath); 
                    if (uploadFileV2(clientContext, oList, filePath, year))
                    {
                        Console.WriteLine(filePath + " --> OK");
                    }
                    else
                    {
                        Console.WriteLine(filePath + " --> NOTOK");
                    }

                    //string[] parts = filePath.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                    //if (parts.Length.Equals(5))
                    //{

                    //    string protocolNo = parts[3];
                    //    string fileName = parts[4];

                    //    byte[] Content = System.IO.File.ReadAllBytes(filePath);
                    //    Console.WriteLine("OK");
                    //}
                }

                //foreach (string file in Directory.EnumerateFiles(filePath, "*", SearchOption.TopDirectoryOnly))
                //{
                //    if (uploadFile(clientContext, oList, file))
                //    {
                //        Console.WriteLine(file + " --> OK");
                //    }
                //}
            }

            Console.WriteLine("All files have downloaded!");
            Console.ReadLine();
        }


        private bool uploadFileV2(ClientContext clientContext, List list, string filePath, string year)
        {
            try
            {
                string[] parts = filePath.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length.Equals(5))
                {
                    string protocolNo = parts[3];
                    string fileName = parts[4];
                    
                        byte[] Content = System.IO.File.ReadAllBytes(filePath);
                    Console.WriteLine("OK");
                    FileCreationInformation fileCreationInformation = new FileCreationInformation
                    {

                        Content = System.IO.File.ReadAllBytes(filePath),
                        Url = System.IO.Path.Combine("GeneralProtocol_" + year + "/" + year + "-" + protocolNo.PadLeft(7, '0') + "/", System.IO.Path.GetFileName(fileName)),
                        Overwrite = false
                    };
                    //List list = clientContext.Web.Lists.GetByTitle("Documents");
                    Microsoft.SharePoint.Client.File uploadFile = list.RootFolder.Files.Add(fileCreationInformation);
                    ListItem itm = uploadFile.ListItemAllFields;
                    itm["Title"] =  fileName;
                    itm.Update();
                    clientContext.Load(uploadFile);
                    clientContext.Load(itm);
                    clientContext.ExecuteQuery();
                    Console.WriteLine("imported to:" + fileCreationInformation.Url + "    in root folder:" + list.RootFolder); 
                    return true;
                }
                else { return false;  }
                //else
                //{
                //    Console.WriteLine("Warning not uploaded: " + filePath + " --> Folder ");
                //    return false;
                //}
            }
            catch (Exception e)
            {
                Console.WriteLine("Error on uploading: " + filePath + " --> " + e.Message);
                return false;
            }

        }


        private bool uploadFile(ClientContext clientContext, List list, string fileName)
        {
            try
            {
                string folderName = fileName.Split('.')[0];
                string[] parts = folderName.Split('\\');
                folderName = parts[parts.Length - 1];
                if (createFolder(clientContext, list, folderName))
                {
                    FileCreationInformation fileCreationInformation = new FileCreationInformation
                    {

                        Content = System.IO.File.ReadAllBytes(fileName),
                        Url = System.IO.Path.Combine("Documents/" + folderName + "/", System.IO.Path.GetFileName(fileName)),
                        Overwrite = false
                    };
                    //List list = clientContext.Web.Lists.GetByTitle("Documents");
                    Microsoft.SharePoint.Client.File uploadFile = list.RootFolder.Files.Add(fileCreationInformation);
                    ListItem itm = uploadFile.ListItemAllFields;
                    itm["Title"] = "a-" + fileName;
                    itm.Update();
                    clientContext.Load(uploadFile);
                    clientContext.Load(itm);
                    clientContext.ExecuteQuery();
                    return true;
                }
                else
                {
                    Console.WriteLine("Error on uploading: " + fileName + " --> FolderName not Created");
                    return false;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error on uploading: " + fileName + " --> " + e.Message);
                return false;
            }

        }


        private bool createFolder(ClientContext clientContext, List list, string folderName)
        {
            try
            {

                var folder = list.RootFolder;
                clientContext.Load(folder);
                clientContext.ExecuteQuery();
                ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
                newItemInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                newItemInfo.LeafName = folderName;
                ListItem newListItem = list.AddItem(newItemInfo);
                newListItem["Title"] = folderName;
                newListItem.Update();
                clientContext.ExecuteQuery();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error on uploading: " + folderName + " --> " + e.Message);
                return false;
            }

        }

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
