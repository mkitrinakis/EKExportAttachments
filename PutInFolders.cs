using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO; 

namespace ExportAttachments
{
    class PutInFolders
    {
        public void run()
        {
            foreach (string filePath in Directory.EnumerateFiles(Directory.GetCurrentDirectory(), "*", SearchOption.TopDirectoryOnly))
            {
                FileInfo fileInfo = new FileInfo(filePath);
                string fileName = fileInfo.Name;
                Console.WriteLine(fileName);
                int sno = 0;
                bool isNo = int.TryParse(fileName.Replace(" ", "_").Split('_')[0], out sno);
                if (isNo)
                {
                    process(fileInfo, sno);
                }

            }
        }

        public void process(FileInfo fileInfo, int sno)
        {
            try
            {
                Console.WriteLine("processing: " + fileInfo.Name);
                string folderName = "_" + sno.ToString().PadLeft(7, '0');
                System.IO.Directory.CreateDirectory(folderName);
                fileInfo.MoveTo(folderName + "\\" + fileInfo.Name);
                Console.WriteLine("processed: " + fileInfo.Name);
            }
            catch (Exception e) { Console.WriteLine("Error on " + fileInfo.Name + ": " + e.Message); }
        }
    }
}
