using System.IO.Compression;
using System.Xml;
namespace Code
{
    class Program
    {
        static void Main()
        {
            string filePath = @"D:\odt_word_doc_1.odt";
            string odfType = ".odt";
            string text = ReadText(filePath, odfType);
            Console.WriteLine(text);
            Console.ReadLine();
        }
        public static string ReadText(string filePath, string odfType)
        {
            string textContent = "";

            using (ZipArchive zipArchive = ZipFile.OpenRead(filePath))
            {
                foreach (var entry in zipArchive.Entries)
                {
                    if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        using StreamReader reader = new StreamReader(entry.Open());
                        string xmlContent = reader.ReadToEnd(); // o/p of content.xml in formatted manner
                        textContent += ExtractTextFromXml(xmlContent, odfType);
                    }
                }
            }

            return textContent; // output screenshot
        }

        public static string ExtractTextFromXml(string xmlContent, string odfType)
        {
            string textContent = "";

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xmlContent);
            XmlNamespaceManager nsManager = new XmlNamespaceManager(xmlDoc.NameTable);
            nsManager.AddNamespace("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"); // common for all files

            // add required namespace for different types of documents

            // for word doc file [libreoffice writer]
            if (odfType == ".odt")
            {
                nsManager.AddNamespace("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
                foreach (XmlNode node in xmlDoc.SelectNodes("//text:p | //text:h", nsManager))
                {
                    textContent += node.InnerText + Environment.NewLine;
                }
            }

            // for spreadsheet file [libreoffice calc]
            else if (odfType == ".ods")
            {
                nsManager.AddNamespace("table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0");
                foreach (XmlNode node in xmlDoc.SelectNodes("//table:table-cell/text:p", nsManager))
                {
                    textContent += node.InnerText + "\t"; // Tab-separated for spreadsheet data
                }
                textContent += Environment.NewLine;
            }

            // for presentations/ppt file [libreoffice calc]
            else if (odfType == ".odp")
            {
                nsManager.AddNamespace("presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0");
                foreach (XmlNode node in xmlDoc.SelectNodes("//office:body//text:p", nsManager))
                {
                    textContent += node.InnerText + Environment.NewLine;
                }
            }




            return textContent;
        }

    }
}