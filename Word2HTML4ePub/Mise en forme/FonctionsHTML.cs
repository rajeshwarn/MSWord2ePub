using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Reflection;
using System.IO;
using System.Xml;
using System.Xml.XPath;

namespace Word2HTML4ePub
{
    public partial class WordHTML2ePubHTML
    {
        public static void RemoveIndent(ref XPathNavigator lir)
        {
            RemoveIndent(ref lir, "/html/body//p");
            //Suppression dans les titres...
            for (int i = 1; i < 8; i++)
                RemoveIndent(ref lir, "/html/body//h" + i.ToString());
        }

        private static void RemoveIndent(ref XPathNavigator lir, string XPath)
        {

            System.Xml.XPath.XPathNodeIterator it = lir.Select(lir.Compile(XPath));
            while (it.MoveNext())
            {
                if (!it.Current.MoveToChild(XPathNodeType.Text))
                    continue;
                do
                {
                    StringBuilder sb = new StringBuilder();
                    string[] sentence = it.Current.Value.Split('\r');

                    foreach (string s in sentence)
                    {
                        string s1 = s;

                        if (string.IsNullOrEmpty(s1))
                            continue;

                        if (s1[0].Equals('\n'))
                        {
                            s1 = s1.Replace('\n', ' ');
                            s1 = s1.TrimStart(null);
                        }

                        sb.Append(s1 + " ");
                    }
                    it.Current.SetValue(sb.ToString());
                } while (it.Current.MoveToNext(XPathNodeType.Text));
                it.Current.MoveToParent();
            }
        }
    }
}