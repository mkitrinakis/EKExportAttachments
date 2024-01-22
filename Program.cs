using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;

namespace ExportAttachments
{
    class Program
    {

        static void Main(string[] args)
        {
            //CheckSharepointAttachments csa = new CheckSharepointAttachments();
            //csa.run();
            //return; 

            //PutInFolders pif = new PutInFolders();
            //pif.run();
            //return; 
            //ImportAtts import = new ImportAtts();
            //import.run("2022");
            //return; 
            ProtocolEA ea = new ProtocolEA();
        //     ea.exportReport();
            ea.importReport(); 
            return; 

            Test t = new Test();
            t.run2();
            return;
            string url = args[0];
            foreach (string arg in args)
            {
                if (arg.ToLower().Contains("createrelatedprotocoldisplay"))
                {
                    CreateRelatedProtocolDisplay crpd = new CreateRelatedProtocolDisplay();
                    crpd.run(url);
                }
                if (arg.ToLower().Contains("exporttocsv"))
                {
                    ExportToCSV csv = new ExportToCSV();
                    csv.run(url);
                }
                if (arg.ToLower().Contains("exportattachments"))
                {
                    ExportAtts atts = new ExportAtts();
                    atts.run(url);
                }
                if (arg.ToLower().Contains("importattachments"))
                {
                    ImportAtts atts = new ImportAtts();
              //      atts.run(url);
                }
            }



        }
    }
}
