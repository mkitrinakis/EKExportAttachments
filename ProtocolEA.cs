using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using System.IO;
using System.Collections; 

namespace ExportAttachments
{
    class ProtocolEA
    {
        public void exportReport()
        {
            using (SPSite site = new SPSite("https://easervice.epant.gr/"))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPList l = web.Lists["Πρωτόκολλο Επιτροπής Ανταγωνισμού"];
                    using (StreamWriter writer = new StreamWriter("log2.txt"))
                    {
                        foreach (SPListItem itm in l.Items)
                        {
                            try
                            {
                                string pno = (itm["ProtocolNumber"] ?? "").ToString();
                                string pto = (itm["Sender"] ?? "").ToString();
                                if (!pto.Trim().Equals("")) { 
                                   pto=pto.Split('#')[1];
                                    writer.WriteLine(pno + "#" + pto);
                                }
                            }
                            catch (Exception e) { Console.WriteLine("ERROR ON:" + itm.ID + " : " + e.Message); }
                        }
                    }
                }
            }
        }


        // https://arxeio.epant.gr/sites/EpantProtocols/All/Lists/Grammateias/
        // ProtocolNumber
        // Katathetis

        // Traders
        //TraderFullName 


        public void importReport()
        {
            using (SPSite site = new SPSite("https://arxeio.epant.gr/sites/EpantProtocols/All/"))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPList lSenders = web.Lists["Συναλλασσόμενοι"];
                    Hashtable protContacts  =  fillProtContact();
                    Console.WriteLine("filled proContacts"); 
                    Hashtable senders = fillSenders(lSenders);
                    Console.WriteLine("filled senders");
                    SPList l = web.Lists["ΠΡΩΤΟΚΟΛΛΟ ΓΡΑΦΕΙΟΥ ΓΡΑΜΜΑΤΕΩΝ ΕΑ"]; 
                    foreach (SPListItem itm in l.Items)
                    {
                        string pno = (itm["ProtocolNumber"] ?? "").ToString(); 
                        if (!pno.Trim().Equals(""))
                        {
                            if (protContacts.ContainsKey(pno))
                            {
                                string contact = (protContacts[pno]).ToString();
                                if (senders.ContainsKey(contact))
                                {
                                    string val = (senders[contact] ?? "").ToString() + ";#" + contact;
                                    itm["Katathetis"] = val;
                                    itm.SystemUpdate(); 
                                    Console.WriteLine(itm.ID + "-->" + val); 

                                }
                                else
                                {
                                    Console.WriteLine("EROR CONTACT NOT FOUND: " + pno + " # " + contact);
                                }
                            }
                            else
                            {
                                Console.WriteLine("ERROR NOT FOUND :" + pno); 
                            }
                        }
                    }
                }
            }
        }


        private Hashtable fillProtContact()
        {
            Hashtable rs = new Hashtable();
            string[] lines = File.ReadAllLines("log2.txt");
            foreach (string line in lines)
            {
                string key = line.Split('#')[0]; //protocolNo 
                string val = line.Split('#')[1]; //Description of sender  
                if (!rs.ContainsKey(key))
                {
                    rs[key] = val; 
                }
            }
            return rs; 
        }

        private Hashtable fillSenders(SPList l)
        {
            Hashtable rs = new Hashtable();
            foreach (SPListItem itm in l.Items)
            {
                string title = (itm["TraderFullName"] ?? "").ToString();
                if (!title.Equals(""))
                {
                    if (!rs.ContainsKey(title))
                        rs[title] = itm.ID;
                }
            }
            return rs; 
        }


    }
}
