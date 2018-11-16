using HtmlAgilityPack;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Parser
{
    class Program
    {
        static void Main(string[] args)
        {
            int pagenum = 1;
            int linknum = 0;
            string pageurl = "https://www.olx.ua/moda-i-stil/odezhda/muzhskaya-odezhda/?search%5Bfilter_enum_state%5D%5B0%5D=new&search%5Bprivate_business%5D=business";
            string pagenumadd = "";
            string[] alllinks = new string[500];

            var url = pageurl;
            var web = new HtmlWeb();
            var doc = web.Load(url);

            while (linknum < 499)
            {
                url = pageurl + pagenumadd;
                web = new HtmlWeb();
                doc = web.Load(url);
                HtmlNodeCollection links = doc.DocumentNode.SelectNodes("//a[@class='thumb vtop inlblk rel tdnone linkWithHash scale4 detailsLink ']");

                foreach (HtmlNode link in links)
                {
                    alllinks[linknum] = (link.GetAttributeValue("href", ""));
                    if (linknum == 499)
                    {
                        linknum++;
                        break;
                    }
                    linknum++;
                }
                pagenum++;
                pagenumadd = "&page=" + pagenum.ToString();
            }

            List<Row> db = new List<Row>();

            System.Drawing.Image image;
            List<string> imageurl;
            int imagenumber;

            db.Add(new Row { Column1 = "Посилання на оголошення:", Column2 = "ID оголошення", Column3 = "Назва оголошення", Column4 = "Місто", Column5 = "Ім'я профілю", Column6 = "К-ть переглядів", Column7 = "Шлях до фото з оголошень" });
            Row r;
            char[] charsToTrim = new char[] { ' ' }; 

            string rootPath;
            string fileName;
            string path = "";

            for (int i = 0; i <= 499; i++)
            {
                url = alllinks[i];
                web = new HtmlWeb();
                doc = web.Load(url);
                r = new Row();
                r.Column1 = url;
                r.Column2 = doc.DocumentNode.SelectSingleNode("//em//small").InnerText.Replace("Номер объявления: ", ""); 
                r.Column3 = doc.DocumentNode.SelectSingleNode("//div[@class='offer-titlebox']//h1").InnerText.Trim();
                r.Column4 = doc.DocumentNode.SelectSingleNode("//a[@class='show-map-link']//strong").InnerText;
                r.Column5 = doc.DocumentNode.SelectSingleNode("//h4//a").InnerText.Trim();
                r.Column6 = doc.DocumentNode.SelectSingleNode("//div[@class='pdingtop10']//strong").InnerText;

                imagenumber = 1;
                imageurl = new List<string>();
                foreach (HtmlNode imgurl in doc.DocumentNode.SelectNodes("//div[@class='photo-glow']//img"))
                    imageurl.Add(imgurl.GetAttributeValue("src", ""));
            
                path = "";
                foreach (string im in imageurl)
                {                    
                    image = DownloadImageFromUrl(im);
                    if (image == null)
                        continue;
                    rootPath = "C:\\images" + "\\" + i;
                    Directory.CreateDirectory(rootPath);
                    fileName = Path.Combine(rootPath, imagenumber + ".jpg");
                    path += fileName + ", ";
                    image.Save(fileName);
                    imagenumber++;
                }
                r.Column7 = path.Remove(path.Length - 2);
                db.Add(r);
                Console.WriteLine(i + 1);
            }
            ExportToExcel(db, "Parser");
        }

        public static void ExportToExcel<T>(IEnumerable<T> data, string worksheetTitle)
        {
            var wb = new XLWorkbook(); 
            var ws = wb.Worksheets.Add(worksheetTitle);             

            if (data != null && data.Count() > 0)
            {
                ws.Cell(1, 1).InsertData(data);
            }
            wb.SaveAs("Parser.xlsx");
        }

        public static System.Drawing.Image DownloadImageFromUrl(string imageUrl)
        {
            System.Drawing.Image image = null;

            try
            {
                System.Net.HttpWebRequest webRequest = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(imageUrl);
                webRequest.AllowWriteStreamBuffering = true;
                webRequest.Timeout = 30000;

                System.Net.WebResponse webResponse = webRequest.GetResponse();

                System.IO.Stream stream = webResponse.GetResponseStream();

                image = System.Drawing.Image.FromStream(stream);

                webResponse.Close();
            }
            catch (Exception ex)
            {
                return null;
            }

            return image;
        }
    }

    public class Row
    {
        public string Column1 { get; set; }
        public string Column2 { get; set; }
        public string Column3 { get; set; }
        public string Column4 { get; set; }
        public string Column5 { get; set; }
        public string Column6 { get; set; }
        public string Column7 { get; set; }
    }

}
