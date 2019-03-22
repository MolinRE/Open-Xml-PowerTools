using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace HtmlConverter01
{
    public class Common
    {
        public static XElement ReadHtmlDocument(string file)
        {
            var htmlContent = File.ReadAllText(file);
            var doc = new HtmlDocument();
            doc.OptionWriteEmptyNodes = true;
            doc.LoadHtml(htmlContent);

            var divElement = new XElement("body");
            var nodes = doc.DocumentNode.ChildNodes;
            foreach (HtmlNode node in nodes.Where(p => p.NodeType != HtmlNodeType.Text))
            {
                divElement.Add(XParse(node));
            }

            var html = new XElement("html", divElement);
            return html;
        }

        private static XElement XParse(HtmlNode node)
        {
            try
            {
                // Чистим комментарии, т.к. там бывают неподдерживаемые элементы (if и т.д.)
                foreach (var inner in node.Descendants()
                    .Where(p => p.NodeType == HtmlNodeType.Comment)
                    .ToArray())
                {
                    inner.Remove();
                }

                foreach (var inner in node.Descendants()
                    .ToArray())
                {
                    if (inner.Name.Contains(':'))
                    {
                        if (inner.InnerHtml.Trim().Length == 0)
                        {
                            inner.Remove();
                        }
                        else
                        {
                            inner.Name = inner.Name.Substring(inner.Name.IndexOf(':') + 1);
                        }
                    }

                    var badAttrs = inner.Attributes.Where(p => p.Name.Contains(':'));
                    foreach (var attr in badAttrs)
                    {
                        attr.Name = attr.Name.Substring(attr.Name.IndexOf(':') + 1);
                    }
                }

                string content = node.OuterHtml;
                content = content
                    .Replace("&nbsp;", " ");

                var result = XElement.Parse(content);
                result.DescendantsAndSelf().Attributes().Where(p => p.Name.LocalName.ToLower().StartsWith("mso")).Remove();

                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return new XElement("error", new XAttribute("type", ex.GetType().Name), ex.Message);
            }

        }
    }
}
