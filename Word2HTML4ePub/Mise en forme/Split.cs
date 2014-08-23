using System.IO;
using System.Xml;
namespace Word2HTML4ePub
{
    public partial class WordHTML2ePubHTML
    {
        public static string SplitHTMLFile(string originalFN, string debut, string fin)
        {
            if (!File.Exists(originalFN))
                return null;

            XmlDocument xml = new XmlDocument();
            xml.Load(originalFN);

            int indexst = 0;
            int indexen = xml.OuterXml.Length;

            if (!string.IsNullOrEmpty(debut))
                indexst = xml.OuterXml.IndexOf(debut);

            if (!string.IsNullOrEmpty(fin))
                indexen = xml.OuterXml.IndexOf(fin, indexst);

            if (indexen - indexst < 0)
            {
                indexen = indexst + 10;
            }

            return xml.OuterXml.Substring(indexst, indexen - indexst);
        }

        public static string getHTMLHeader(string originalFN)
        {
            if (!File.Exists(originalFN))
                return null;

            XmlDocument xml = new XmlDocument();
            xml.Load(originalFN);

            string exPath = "/html/body";
            //            XPathNavigator node = lir.SelectSingleNode(lir.Compile(exPath));

            XmlNodeList list = xml.SelectNodes(exPath);

            string bloc = null;
            if (list[0].InnerXml.Length > 50)
                bloc = list[0].InnerXml.Substring(0, 50);
            else
                bloc = list[0].InnerXml;

            int index = xml.InnerXml.IndexOf(bloc);


            return xml.InnerXml.Substring(0, index);
        }


    }
}