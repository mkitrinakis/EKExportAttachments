using System;
using System.Collections.Generic;
using System.Collections;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint;
using System.Text;
using System.IO;

namespace ExportAttachments 
{
    class ExportToCSV
    {
        // string listName = "PSCC";
        // string viewName = "_admin";
        //string listName = "Γενικό Πρωτόκολλο";
        string listName = "xxx";
        string viewName = "toExport";
        char separator = ';';
        SPList list;
        SPView view;

        public void run(string url)
        {
            init(url);
            using (SPSite site = new SPSite(url))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    list = web.Lists[listName];
                    view = list.Views[viewName];
                    exportLargeData();
                }
            }
        }
        public void exportLargeData()
        {


            //HttpContext.Current.Response.Clear();
            //HttpContext.Current.Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", "report.xls"));
            //HttpContext.Current.Response.ContentType = "application/ms-excel";

            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                {

                    string rs = "";
                    foreach (string fname in view.ViewFields)
                    {

                        rs += customConvert(fname) + ";";
                    }
                    rs += Environment.NewLine;

                    SPQuery q = new SPQuery();
                    q.RowLimit = 50;
                    q.Query = view.Query;
                    q.ViewAttributes = "Scope='RecursiveAll'";
                    int found = 1000;
                    int pos = 0;
                    SPListItemCollection col = null;
                    using (StreamWriter sWriter = new StreamWriter("metadata.csv"))
                    {
                        while (found > 0)
                        {
                            if (col != null)
                            {
                                int pid = col[col.Count - 1].ID;
                                SPListItemCollectionPosition position = new SPListItemCollectionPosition("Paged=TRUE&p_ID=" + pid);
                                q.ListItemCollectionPosition = position;
                                Console.WriteLine("processing up to " + pid + "...");
                            }

                            col = list.GetItems(q);
                            found = col.Count;

                            //    if (found > 0)
                            //    {
                            //        pos = col[found - 1].ID;
                            //    }
                            foreach (SPListItem i in col)
                            {

                                foreach (string fname in view.ViewFields)
                                {

                                    string txt = customConvert(i[fname]);
                                    rs += txt + ";";
                                }
                                rs += Environment.NewLine;
                            }
                            sWriter.Write(rs);
                            rs = "";
                        }

                    }


                    //HttpContext.Current.Response.Write(sw.ToString());
                    //HttpContext.Current.Response.End();
                }
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

            Console.WriteLine("csv seperator is ; -Press Enter to keep it or press the separator you want and Enter");
            tmp = Console.ReadLine();
            if (!tmp.Trim().Equals(""))
            {
                separator = tmp.Trim()[0];
                Console.WriteLine("separator changed to :" + separator);
            }

        }

        private string customConvert(object val)
        {
            if (val == null) { return ""; }
            try
            {
                string rs = (val ?? "-").ToString();
                if (rs.Contains(";#")) rs = rs.Split('#')[1];
                if (val.GetType().Equals(typeof(double)) || val.GetType().Equals(typeof(Double)))
                {
                    if (separator.Equals(','))
                    {
                        return rs.Replace('.', ',');
                    }
                    else
                    {
                        return rs.Replace(',', '.');
                    }
                }
                if (separator.Equals(','))
                {
                    rs.Replace(',', ';');
                    rs.Replace(Environment.NewLine, "  ");
                    rs.Replace("\n", " ");
                }
                else
                {
                    rs.Replace(';', ',');
                    rs.Replace(Environment.NewLine, ", ");
                    rs.Replace("\n", ", ");
                }


                //rs = rs.Replace("<rows>", "");
                //rs = rs.Replace("<rows/>", "");
                //rs = rs.Replace("</rows>", "");
                //rs = rs.Replace("<row", "");
                //rs = rs.Replace("/>", "");
                //rs = rs.Replace(">", "");
                //rs = rs.Replace("<", "");
                //rs = rs.Replace("&", "&amp");
                //rs = rs.Replace("'", "&quot");
                return rs;
            }
            catch { return (val ?? "").ToString(); }
        }

    }

}
