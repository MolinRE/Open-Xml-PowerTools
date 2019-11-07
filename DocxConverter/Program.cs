
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using OpenXmlPowerTools.HtmlToWml;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Net;
using System.Text.RegularExpressions;
using CommonPhrasesBuilder;
using Dapper;
using Microsoft.Office.Interop.Word;
using WebArmModels.Document;
using WebArmModels.Upload;

namespace DocxConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Не пихать обычный HTML. Код настроен на формат веб-арма
            string directory = @"C:\Users\k.komarov\source\example\docx\";

            foreach (var file in Directory.GetFiles(directory, "118_69738" + ".xml"))
            {
                ConvertToDocx(file, directory);
            }

        }

        private static void ConvertToDocx(string file, string destinationDir)
        {
            var sourceHtmlFi = new FileInfo(file);
            Console.WriteLine("Converting " + sourceHtmlFi.Name);
            
            var destDocxFi = new FileInfo(Path.ChangeExtension(file, "docx"));
            
            var page = new CkePage();
            page.HtmlContent = File.ReadAllText(file);

            ConvertToDocx(page, false, destDocxFi.FullName);
        }

        /// <summary>
        /// Заменяет элементы "пользовательский стиль" на подчёркивание.
        /// </summary>
        /// <param name="element"></param>
        public static void ResetFill(XElement element)
        {
            foreach (var fillElem in element.Descendants("fill").ToArray())
            {
                var replaceValue = new XText(new string('_', fillElem.Value.Length));
                fillElem.AddAfterSelf(replaceValue);
                fillElem.Remove();
            }
        }

        /// <summary>
        /// Пробует прочитать атрибут ориентации и возвращает его значение. Если атрибута нет, возвращает значение по умолчанию - "landscape".
        /// </summary>
        /// <param name="document">XML-документ.</param>
        /// <returns></returns>
        public static string GetDocumentOrientation(XElement document)
        {
            string result = "portrait";
            var orientation = document.Descendants("orientation").FirstOrDefault();
            if (orientation == null || !orientation.HasAttribute("val"))
            {
                return result;
            }

            if (orientation.Attribute("val").Value == "landscape")
            {
                result = "landscape";
            }


            return result;
        }

        /// <summary>
        /// Конвертирует HTML-контент в DOCX файл и сохраняет его как аттачмент к документу.
        /// </summary>
        /// <param name="page">Идентификатор и контент документа.</param>
        /// <param name="fill">Сохранять ли значения, размеченные пользовательским выделением</param>
        /// <returns></returns>
        public static void ConvertToDocx(CkePage page, bool fill, string dest)
        {
            string htmlContent = page.HtmlContent;

            var xml = new HtmlToXml().TranslateToXml(htmlContent);
            var html = XmlExtension.GetXElement(xml.DocumentElement);
            string orientation = GetDocumentOrientation(html);
            html = html.Descendants("xmlcontent").FirstOrDefault();
            if (html == null)
            {
                return;
            }

            #region подготовка html к экспорту в ворд

            // Удаление атрибута 'tempid'
            html.Descendants().Attributes("tempid").Remove();

            //удаляем гиперссылки и якоря из контента
            html.Descendants("a").Where(a => a.HasAttribute("name")).Remove();

            var linksElem = html.Descendants("a").Where(a => a.HasAttribute("href")).ToList();
            foreach (var link in linksElem)
            {
                link.ReplaceWith(link.DescendantNodes());
            }

            //заменям теги heading на h1,h2...
            var headingElems = html.Descendants("heading");
            foreach (var heading in headingElems)
            {
                var level = heading.Attribute("level")?.Value ?? "1";
                heading.Name = "h" + level;
            }

            foreach (var table in html.Descendants("table"))
            {
                var width = table.Attribute("width");
                if (width != null)
                {
                    table.SetStyle("width", width.Value);
                    width.Remove();
                }

                var height = table.Attribute("height");
                if (height != null)
                {
                    if (!height.Value.EndsWith("%"))
                    {
                        table.SetStyle("height", height.Value);
                    }
                    height.Remove();
                }
            }


            #endregion

            html = new XElement("html",
                new XElement("head",
                new XElement("style")),
                new XElement("body", html));

            if (!fill)
            {
                //ResetFill(html);
            }

            HtmlToWmlConverterSettings settings = HtmlToWmlConverter.GetDefaultSettings();
            settings.GetImageHandler = GetImageHandler;
            if (orientation == "landscape")
            {
                settings.Orientation = PageOrientationValues.Landscape;
            }

            //try
            //{
                WmlDocument doc = HtmlToWmlConverter.ConvertHtmlToWml(defaultCss, userCss, html, settings);

                doc.SaveAs(dest);
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //    Console.WriteLine("Нажмите любую клавишу для продолжения.");
            //    Console.ReadKey();
            //}
        }

        private static readonly Regex ImageRegex = new Regex(@"/data/doc/image/(?<iid>-?\d+)\?moduleId=(?<mid>-?\d+)&id=(?<id>-?\d+)?$", RegexOptions.Compiled);

        private static byte[] GetImageHandler(XElement element)
        {
            string srcAttribute = (string)element.Attribute(XhtmlNoNamespace.src);
            byte[] result = null;

            if (srcAttribute.StartsWith("data:"))
            {
                var semiIndex = srcAttribute.IndexOf(';');
                var commaIndex = srcAttribute.IndexOf(',', semiIndex);
                var base64 = srcAttribute.Substring(commaIndex + 1);
                result = Convert.FromBase64String(base64);

                return result;
            }

            var match = ImageRegex.Match(srcAttribute);
            if (match.Success)
            {
                var imgId = Convert.ToInt32(match.Groups["iid"].Value);
                var image = GetImageFromDb(imgId);
                result = image.ImageData;
                //element.SetStyle("width", image.Width.ToString() + "px");
                //element.SetStyle("height", image.Height.ToString() + "px");
                //element.Add(
                //    new XAttribute("width", image.Width),
                //    new XAttribute("height", image.Height));

                return result;

                //var request = WebRequest.CreateHttp("http://web-arm-service1.aservices.tech" + srcAttribute.Substring(5));
                //request.Headers.Add("x-wa", "{\"wa_i\": 100011, \"wa_t\": \"local\", \"wa_l\":\"ru\"}");
                ////request.Headers.Add("X-Ist", Guid.NewGuid().ToString());
                ////request.Host = "web-arm-service1.aservices.tech";
                //using (var resp = request.GetResponse() as HttpWebResponse)
                //{
                //    using (var stream = resp.GetResponseStream())
                //    {
                //        if (stream == null)
                //        {
                //            return null;
                //        }

                //        result = new byte[resp.ContentLength];
                //        stream.Read(result, 0, result.Length);


                //        using (MemoryStream ms = new MemoryStream(result))
                //        {
                //            var bmp = new Bitmap(ms);
                //            bmp.Save("image.jpg");
                //        }

                //        return result;
                //    }
                //}
            }

            // Читаем из файла
            return result;
        }

        public static WarmImage GetImageFromDb(int imageId)
        {
            using (var conn = new SqlConnection("Data Source=srv12.sps.m1.amedia.tech;Initial Catalog=RBD_dev;User Id=service.webarm;PASSWORD=WhS7LIwtPKNO;"))
            {
                var result = conn.QueryFirstOrDefault<WarmImage>("SELECT * FROM dbo.[Image] img WHERE img.ImgID = @imageId", new {imageId});

                return result;
            }
        }

        // Given a document name, set the print orientation for 
        // all the sections of the document.
        public static void SetPrintOrientation(string fileName, PageOrientationValues newOrientation)
        {
            using (var document = WordprocessingDocument.Open(fileName, true))
            {
                bool documentChanged = false;

                var docPart = document.MainDocumentPart;
                var sections = docPart.Document.Descendants<SectionProperties>();

                foreach (SectionProperties sectPr in sections)
                {
                    bool pageOrientationChanged = false;

                    PageSize pgSz = sectPr.Descendants<PageSize>().FirstOrDefault();
                    if (pgSz != null)
                    {
                        // No Orient property? Create it now. Otherwise, just 
                        // set its value. Assume that the default orientation 
                        // is Portrait.
                        if (pgSz.Orient == null)
                        {
                            // Need to create the attribute. You do not need to 
                            // create the Orient property if the property does not 
                            // already exist, and you are setting it to Portrait. 
                            // That is the default value.
                            if (newOrientation != PageOrientationValues.Portrait)
                            {
                                pageOrientationChanged = true;
                                documentChanged = true;
                                pgSz.Orient = new EnumValue<PageOrientationValues>(newOrientation);
                            }
                        }
                        else
                        {
                            // The Orient property exists, but its value
                            // is different than the new value.
                            if (pgSz.Orient.Value != newOrientation)
                            {
                                pgSz.Orient.Value = newOrientation;
                                pageOrientationChanged = true;
                                documentChanged = true;
                            }
                        }

                        if (pageOrientationChanged)
                        {
                            // Changing the orientation is not enough. You must also 
                            // change the page size.
                            var width = pgSz.Width;
                            var height = pgSz.Height;
                            pgSz.Width = height;
                            pgSz.Height = width;

                            PageMargin pgMar = sectPr.Descendants<PageMargin>().FirstOrDefault();
                            if (pgMar != null)
                            {
                                // Rotate margins. Printer settings control how far you 
                                // rotate when switching to landscape mode. Not having those
                                // settings, this code rotates 90 degrees. You could easily
                                // modify this behavior, or make it a parameter for the 
                                // procedure.
                                var top = pgMar.Top.Value;
                                var bottom = pgMar.Bottom.Value;
                                var left = pgMar.Left.Value;
                                var right = pgMar.Right.Value;

                                pgMar.Top = new Int32Value((int)left);
                                pgMar.Bottom = new Int32Value((int)right);
                                pgMar.Left = new UInt32Value((uint)System.Math.Max(0, bottom));
                                pgMar.Right = new UInt32Value((uint)System.Math.Max(0, top));
                            }
                        }
                    }
                }
                if (documentChanged)
                {
                    docPart.Document.Save();
                }
            }
        }

        public class HtmlToWmlReadAsXElement
        {
            public static XElement ReadAsXElement(FileInfo sourceHtmlFi)
            {
                string htmlString = File.ReadAllText(sourceHtmlFi.FullName);
                XElement html = null;
                try
                {
                    html = XElement.Parse(htmlString);
                }
                catch (XmlException e)
                {
                    throw e;
                }

                // HtmlToWmlConverter expects the HTML elements to be in no namespace, so convert all elements to no namespace.
                html = (XElement)ConvertToNoNamespace(html);
                return html;
            }

            private static object ConvertToNoNamespace(XNode node)
            {
                XElement element = node as XElement;
                if (element != null)
                {
                    return new XElement(element.Name.LocalName,
                        element.Attributes().Where(a => !a.IsNamespaceDeclaration),
                        element.Nodes().Select(n => ConvertToNoNamespace(n)));
                }
                return node;
            }
        }

        static string defaultCss =
            @"html, address,
blockquote,
body, dd, div,
dl, dt, fieldset, form,
frame, frameset,
h1, h2, h3, h4,
h5, h6, noframes,
ol, p, ul, center,
dir, hr, menu, pre { display: block; unicode-bidi: embed }
li { display: list-item }
head { display: none }
table { display: table }
tr { display: table-row }
thead { display: table-header-group }
tbody { display: table-row-group }
tfoot { display: table-footer-group }
col { display: table-column }
colgroup { display: table-column-group }
td, th { display: table-cell }
caption { display: table-caption }
th { font-weight: bolder; text-align: center }
caption { text-align: center }
body { margin: auto; }
h1 { font-size: 2em; margin: auto; }
h2 { font-size: 1.5em; margin: auto; }
h3 { font-size: 1.17em; margin: auto; }
h4, p,
blockquote, ul,
fieldset, form,
ol, dl, dir,
menu { margin: auto }
a { color: blue; }
h5 { font-size: .83em; margin: auto }
h6 { font-size: .75em; margin: auto }
h1, h2, h3, h4,
h5, h6, b,
strong { font-weight: bolder }
blockquote { margin-left: 40px; margin-right: 40px }
i, cite, em,
var, address { font-style: italic }
pre, tt, code,
kbd, samp { font-family: monospace }
pre { white-space: pre }
button, textarea,
input, select { display: inline-block }
big { font-size: 1.17em }
small, sub, sup { font-size: .83em }
sub { vertical-align: sub }
sup { vertical-align: super }
table { border-spacing: 2px; }
thead, tbody,
tfoot { vertical-align: middle }
td, th, tr { vertical-align: inherit }
s, strike, del { text-decoration: line-through }
hr { border: 1px inset }
ol, ul, dir,
menu, dd { margin-left: 40px }
ol { list-style-type: decimal }
ol ul, ul ol,
ul ul, ol ol { margin-top: 0; margin-bottom: 0 }
u, ins { text-decoration: underline }
br:before { content: ""\A""; white-space: pre-line }
center { text-align: center }
:link, :visited { text-decoration: underline }
:focus { outline: thin dotted invert }
/* Begin bidirectionality settings (do not change) */
BDO[DIR=""ltr""] { direction: ltr; unicode-bidi: bidi-override }
BDO[DIR=""rtl""] { direction: rtl; unicode-bidi: bidi-override }
*[DIR=""ltr""] { direction: ltr; unicode-bidi: embed }
*[DIR=""rtl""] { direction: rtl; unicode-bidi: embed }

";

        static string userCss = @"
table {
  border-collapse: collapse;
  border-style: hidden;
}
td .cell_border_all,
th .cell_border_all {
    border-style: solid ;
    border-width: 1px;
    border-color: rgb(0, 0, 0)}
td .cell_border_bottom,
td .cell_border_bottom {
    border-bottom-style: solid ;
    border-bottom-width: 1px;
    border-bottom-color: rgb(0, 0, 0) ;
}
td .cell_border_right,
th .cell_border_right{
    border-right-style: solid ;
    border-right-width: 1px;
    border-right-color: rgb(0, 0, 0) ;
}
td .cell_border_left,
th .cell_border_left {
    border-left-style: solid ;
    border-left-width: 1px;
    border-left-color: rgb(0, 0, 0) ;
}
td .cell_border_top,
th .cell_border_top{
    border-top-style: solid ;
    border-top-width: 1px ;
    border-top-color: rgb(0, 0, 0) ;
}
td {
    padding: 5px;
}";
    }
}
