/***************************************************************************

Copyright (c) Microsoft Corporation 2010.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license
can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

***************************************************************************/

/***************************************************************************
 * IMPORTANT NOTE:
 * 
 * With versions 4.1 and later, the name of the HtmlConverter class has been
 * changed to WmlToHtmlConverter, to make it orthogonal with HtmlToWmlConverter.
 * 
 * There are thin wrapper classes, HtmlConverter, and HtmlConverterSettings,
 * which maintain backwards compat for code that uses the old name.
 * 
 * Other than the name change of the classes themselves, the functionality
 * in WmlToHtmlConverter is identical to the old HtmlConverter class.
***************************************************************************/

using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using RestSharp;

class HtmlConverterHelper
{
    const string documentUrlPattern = "document/(?<moduleid>[^/]+)/(?<id>[^/]+)(/(?<anchor>[^/]+))?";

    static void Main(string[] args)
    {
        var n = DateTime.Now;
        var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        tempDi.Create();

        var picDirPath = @"C:\Users\k.komarov\source\example\pic";
        foreach (var file in Directory.GetFiles(picDirPath, "док с картинками_?" + ".docx"))
        {
            ConvertToHtml(file, picDirPath);
        }
    }

    public static void ConvertToHtml(string file, string outputDirectory)
    {
        var fi = new FileInfo(file);
        Console.WriteLine(fi.Name);
        byte[] byteArray = File.ReadAllBytes(fi.FullName);
        using (var memoryStream = new MemoryStream())
        {
            memoryStream.Write(byteArray, 0, byteArray.Length);
            // Открываем документ
            using (var wDoc = WordprocessingDocument.Open(memoryStream, true))
            {
                var destFileName = new FileInfo(fi.Name.Replace(".docx", ".html"));
                if (outputDirectory != null && outputDirectory != string.Empty)
                {
                    var di = new DirectoryInfo(outputDirectory);
                    if (!di.Exists)
                    {
                        throw new OpenXmlPowerToolsException("Output directory does not exist");
                    }
                    destFileName = new FileInfo(Path.Combine(di.FullName, destFileName.Name));
                }
                var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
                int imageCounter = 0;

                var pageTitle = fi.FullName;
                var part = wDoc.CoreFilePropertiesPart;
                if (part != null)
                {
                    pageTitle = (string) part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fi.FullName;
                }
                
                var settings = new HtmlConverterSettings()
                {
                    AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                    PageTitle = pageTitle,
                    FabricateCssClasses = false,
                    CssClassPrefix = "pt-",
                    RestrictToSupportedLanguages = false,
                    RestrictToSupportedNumberingFormats = false,
                    ImageHandler = imageInfo => ImageHandler(imageInfo, ref imageCounter, imageDirectoryName)
                };
                
                XElement htmlElement = HtmlConverter.ConvertToHtml(wDoc, settings);
                var htmlDocument = PostProcessDocument(htmlElement);
                var htmlString = htmlDocument.ToString(SaveOptions.OmitDuplicateNamespaces);
                File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                File.WriteAllText(destFileName.FullName + ".xml", htmlString, Encoding.UTF8);

                // Produce HTML document with <!DOCTYPE html > declaration to tell the browser
                // we are using HTML5.
                //var htmlDocumentOriginal = new XDocument(
                //    new XDocumentType("html", null, null, null),
                //    HtmlConverter.ConvertToHtml(wDoc, settings));
                //var htmlStringOriginal = htmlDocumentOriginal.ToString(SaveOptions.None);
                //File.WriteAllText(string.Format("{0}\\{1}_html_converter.html", 
                //    destFileName.DirectoryName, Path.GetFileNameWithoutExtension(destFileName.Name)), htmlStringOriginal, Encoding.UTF8);
            }
        }
    }

    // Очищаем от всего, что нагенерил конвертер
    public static XElement PostProcessDocument(XElement html)
    {
        var bodyDiv = html.Element(XN("body")).Element(XN("div"));

        // Активный список
        XElement list = null;
        // Чтобы сравнивать с предыдущим
        var elements = bodyDiv.Elements().ToArray();
        for (int i = 0; i < elements.Length; i++)
        {
            var elem = elements[i];

            if (IsListElement(elem, out string listName))
            {
                if (list == null)
                {
                    list = new XElement(listName);
                    elem.AddBeforeSelf(list);
                }

                // Если элемент описывает список, трансформируем в <li>
                var listItem = TransformToListItemElement(elem);
                if (IsListElement(elements[i - 1]))
                {
                    // Если является вложенным элементом списка
                    if (i > 0 && IsInnerListElement(elem, elements[i - 1]))
                    {
                        // Создаём ul/ol список
                        var innerList = new XElement(listName);
                        list.Elements().Last().Add(innerList);
                        // Записываем его в общий объект
                        list = innerList;
                    }
                    // Если является верхним элементом списка
                    if (i > 0 && IsOuterListElement(elem, elements[i - 1]))
                    {
                        // Поднимается к <li>, в котором делали список, а затем к ul/ol, в котором элемент находился
                        if (list.Parent.Parent.Name == listName)
                        {
                            list = list.Parent.Parent;
                        }
                        // Иначе это какая-то бажжина (кривое форматирование списка в исходном документе, например)
                        // Пишем в основной список, потому что родительского тупо нет
                    }
                }
                // Наконец, добавляем элемент в список нужного уровня
                list.Add(listItem);
                // Ставим пометку, что элемент надо удалить
                elem.Add(new XAttribute("toDelete", true));

                continue;
            }
            else
            {
                if (list != null)
                {
                    list = null;
                }
            }

            TransformHeaders(elem);

            RemoveSpans(elem);
        }

        // Удаляем href ссылок
        foreach (var linkElem in bodyDiv.Descendants(XN("a")).ToArray())
        {
            linkElem.AddAfterSelf(linkElem.Value);
            linkElem.Remove();
        }

        // Ставим самые простые границы у таблиц, иначе они вообще без них приедут
        foreach (var tableElement in bodyDiv.Descendants(XN("table")))
        {
            tableElement.ReplaceAttributes(new XAttribute("class", "bdAll"));
        }

        CleanupStyles(bodyDiv);

        return bodyDiv;
    }

    /// <summary>
    /// Преобразование ссылок по url документа (lnktype="docbyurl")
    /// </summary>
    private static void TransformLinkDocByUrl(XElement elem, int linkNumber)
    {
        // проверка наличия атрибута "href"
        if (AttributeExists(elem, "href") == false)
        {
            //AddValidateonError(elem, "Исправьте ссылку");
            return;
        }

        // проверка соответствия шаблону
        var href = elem.Attribute("href");
        Match match = Regex.Match(href.Value.Trim(), documentUrlPattern);
        if (!match.Success)
        {
            //AddValidateonError(elem, "Укажите полную ссылку, которая бы начиналась с https://");
            return;
        }

        byte moduleTo;
        if (!byte.TryParse(match.Groups["moduleid"].Value.Trim(), out moduleTo))
        {
            //AddValidateonError(elem, "Измените тип (модуль) документа в ссылке");
            return;
        }

        int idTo;
        if (!int.TryParse(match.Groups["id"].Value.Trim(), out idTo))
        {
            //AddValidateonError(elem, "Измените ID документа в ссылке");
            return;
        }

        // пример <a class="doc" href="sp://num=2">
        // замена атрибутов в контенте
        elem.ReplaceAttributes(
            new XAttribute("class", "doc"),
            new XAttribute("href", "sp://num=" + linkNumber));
    }


    /// <summary>
    /// Преобразование ссылок на страницу в интернете (lnktype="weblink")
    /// Пример <!--> <a href="http://www.rostrud.info" title="www.rostrud.info"> -->
    /// </summary>
    private static void TransformLinkWeb(XElement elem)
    {
        // проверка наличия атрибута "href"
        if (AttributeExists(elem, "href") == false)
        {
            return;
        }

        elem.ReplaceAttributes(
            elem.Attribute("href"),
            elem.Attribute("title")
            );
    }

    private static void TransformHeaders(XElement elem)
    {
        // Определяем, является ли элемент заголовком
        int headerLevel = 1;
        if (IsHeaderElement(elem, ref headerLevel))
        {
            elem.Name = headerLevel == 0 ? "p" : "h3";
            elem.Attribute("style").Remove();
            elem.Add(new XAttribute("style", "text-align:center"));

            // Добавляем пометку об оглавлении
            var incw = new XElement("incw",
                // Атрибуты
                new XAttribute("class", "headers"),
                new XAttribute("level", headerLevel),
                // Дочерний элемент
                new XElement("tocitem",
                    new XAttribute("class", "title toc"),
                    new XAttribute("level", headerLevel),
                    new XAttribute("titxt", elem.Value)));
            elem.AddBeforeSelf(incw);
        }
    }

    private static bool IsHeaderElement(XElement elem, ref int headerLevel)
    {
        var result = false;

        // h3 - самый популярный
        if (elem.Name == XN("h3"))
        {
            result = true;
            headerLevel = 3;
        }
        else if (elem.Name == XN("h1"))
        {
            result = true;
            headerLevel = 1;
        }
        else if (elem.Name == XN("h2"))
        {
            result = true;
            headerLevel = 2;
        }
        else if (elem.Name == XN("h4"))
        {
            result = true;
            headerLevel = 4;
        }
        else if (elem.Name == XN("h5"))
        {
            result = true;
            headerLevel = 5;
        }
        else if (elem.Name == XN("h6"))
        {
            result = true;
            headerLevel = 6;
        }

        if (!result)
        {
            if (elem.Attribute("style") != null)
            {
                var centered = elem.Attribute("style").Value.Contains("text-align: center");
                if (centered)
                {
                    var span = elem.Element(XN("span"));
                    if (span != null && span.Attribute("style") != null)
                    {
                        headerLevel = 0;
                        result = span.Attribute("style").Value.Contains("font-weight: bold");
                    }
                }
            }
        }


        return result;
    }

    private static void RemoveSpans(XElement elem)
    {
        var spanElements = elem.Descendants(XN("span")).ToArray();

        RemoveSpans(spanElements);
    }

    private static void RemoveSpans(XElement[] spanElements)
    {
        XElement previousFormattedSpan = null;
        foreach (var span in spanElements)
        {
            if (!span.HasElements)
            {
                var spanStyle = span.Attribute("style").Value.Split(';');
                XElement spanWrap = null;
                if (spanStyle.Any(p => p.Equals("font-style: italic")))
                {
                    var innerWrap = new XElement("em");
                    if (spanWrap != null)
                    {
                        spanWrap.Add(innerWrap);
                    }

                    spanWrap = innerWrap;
                }
                if (spanStyle.Any(p => p.Equals("font-weight: bold")))
                {
                    var innerWrap = new XElement("strong");
                    if (spanWrap != null)
                    {
                        spanWrap.Add(innerWrap);
                    }

                    spanWrap = innerWrap;
                }

                if (spanWrap == null)
                {
                    if (span.Value == " " && previousFormattedSpan != null)
                    {
                        previousFormattedSpan.SetValue(previousFormattedSpan.Value + " ");
                    }
                    else
                    {
                        span.AddAfterSelf(span.Value);
                    }
                    previousFormattedSpan = null;
                }
                else
                {
                    spanWrap.Add(span.Value);
                    while (spanWrap.Parent != null)
                    {
                        spanWrap = spanWrap.Parent;
                    }
                    // Иногда появляются вот такие спаны
                    if (spanWrap.Value != " ")
                    {
                        span.AddAfterSelf(spanWrap);
                        previousFormattedSpan = spanWrap;
                    }

                }
                span.Remove();
            }
            else
            {
                if (span.Elements().Count() == 1)
                {
                    var node = span.Elements().First();
                    span.AddAfterSelf(node);
                    span.Remove();
                }
            }
        }
    }

    private static XElement TransformToListItemElement(XElement elem)
    {
        var result = new XElement("li");
        var spanElements = elem.Elements(XN("span")).ToArray();
        if (spanElements.Length > 1)
        {
            // Первый спан - иконка списка
            RemoveSpans(spanElements.Skip(1).ToArray());
            // При этом иконку важно сохранить для дальнейшего анализа вложенности
            foreach (var node in elem.Nodes().Skip(1))
            {
                result.Add(node);
            }
        }
        return result;
    }

    public static bool IsInnerListElement(XElement pElement, XElement previousElement)
    {
        //margin-left: 1.00in;
        var style = pElement.Attribute("style").Value.Split(';');
        string marginLeft = GetMarginLeftInches(style);
        if (marginLeft == null)
        {
            return false;
        }

        style = previousElement.Attribute("style").Value.Split(';');
        string marginLeftPrevious = GetMarginLeftInches(style);
        if (marginLeft == null)
        {
            return false;
        }

        // InvarianCulture, т.к. в российской культуре дробная часть отделяется запятой
        return Convert.ToSingle(marginLeft, CultureInfo.InvariantCulture) > Convert.ToSingle(marginLeftPrevious, CultureInfo.InvariantCulture);
    }

    private static bool IsOuterListElement(XElement pElement, XElement previousElement)
    {
        //margin-left: 1.00in;
        var style = pElement.Attribute("style").Value.Split(';');
        string marginLeft = GetMarginLeftInches(style);
        if (marginLeft == null)
        {
            return false;
        }

        style = previousElement.Attribute("style").Value.Split(';');
        string marginLeftPrevious = GetMarginLeftInches(style);
        if (marginLeft == null)
        {
            return false;
        }

        // InvarianCulture, т.к. в российской культуре дробная часть отделяется запятой
        return Convert.ToSingle(marginLeft, CultureInfo.InvariantCulture) < Convert.ToSingle(marginLeftPrevious, CultureInfo.InvariantCulture);
    }

    private static string GetMarginLeftInches(string[] style)
    {
        //"0.50in";
        string result = style.FirstOrDefault(p => p.StartsWith("margin-left"));
        if (result == null || result.Length < 15)
        {
            return null;
        }

        return result.Substring(12, result.Length - 2 - 12).Trim();
    }

    public static bool IsListElement(XElement elem, out string listName)
    {
        // Попробовать использовать ref, вместо out?
        listName = "";
        if (elem.Name != XN("p"))
        {
            return false;
        }
        var spanElement = elem.Elements(XN("span")).FirstOrDefault();
        if (spanElement == null)
        {
            return false;
        }
        var styleAttribute = spanElement.Attribute("style");
        if (styleAttribute != null)
        {
            listName = spanElement.Value.Length == 1 ? "ul" : "ol";

            return styleAttribute.Value.Contains("display: inline-block;") &&
                styleAttribute.Value.Contains("text-indent: 0;") &&
                styleAttribute.Value.Contains("width:");
        }

        return false;
        //return spanElement.Value == "" || spanElement.Value == ""
        //    || (spanElement.Value.Length > 1 && int.TryParse(spanElement.Value.Remove(spanElement.Value.Length - 1), out int listIndex));
    }

    public static bool IsListElement(XElement pElement)
    {
        var spanElement = pElement.Elements(XN("span")).FirstOrDefault();
        if (spanElement == null)
        {
            return false;
        }
        var styleAttribute = spanElement.Attribute("style");
        if (styleAttribute != null)
        {
            return styleAttribute.Value.Contains("display: inline-block;") &&
                styleAttribute.Value.Contains("text-indent: 0;") &&
                styleAttribute.Value.Contains("width:");
        }

        return false;
    }

    public static XName XN(string xName)
    {
        return XName.Get(xName, Xhtml.xhtml.NamespaceName);
    }

    private static void CleanupStyles(XElement bodyDiv)
    {
        foreach (var elem in bodyDiv.DescendantsAndSelf())
        {
            elem.Name = elem.Name.LocalName;
            var badAttribute = elem.Attribute("dir");
            if (badAttribute != null)
            {
                badAttribute.Remove();
            }
            badAttribute = elem.Attribute("lang");
            if (badAttribute != null)
            {
                badAttribute.Remove();
            }
            badAttribute = elem.Attribute("class");
            if (badAttribute != null)
            {
                if (elem.Name != "incw")
                {
                    badAttribute.Remove();
                }
            }
            badAttribute = elem.Attribute("style");
            if (badAttribute != null)
            {
                if (elem.Name == "p")
                {
                    if (badAttribute.Value != "text-align:center")
                    {
                        badAttribute.Remove();
                    }
                }
                else if (elem.Name != "h3")
                {
                    badAttribute.Remove();
                }
            }
        }
    }

    public static XElement ImageHandler(ImageInfo imageInfo, ref int imageCounter, string imageDirectoryName)
    {
        byte[] data = null;
        var localDirInfo = new DirectoryInfo(imageDirectoryName);
        if (!localDirInfo.Exists)
        {
            localDirInfo.Create();
        }
        ++imageCounter;

        if (imageInfo.Bitmap == null)
        {
            if (imageInfo.Url == null)
            {
                return null;
            }
            var uri = new Uri(imageInfo.Url);
            var client = new RestClient(uri);
            var response = client.Execute(new RestRequest());
            if (response.IsSuccessful)
            {
                Console.Write("\rЗагружено картинок: {0,4}", imageCounter);
                imageInfo.ContentType = response.ContentType;
                data = response.RawBytes;
            }
            else
            {
                Console.WriteLine("Не удалось загрузить {0}\r\n{1}", imageInfo.Url, response.ErrorMessage);
            }
        }

        string extension = imageInfo.ContentType.Split('/')[1].ToLower();
        ImageFormat imageFormat = null;
        if (extension == "png")
        {
            imageFormat = ImageFormat.Png;
        }
        else if (extension == "gif")
        {
            imageFormat = ImageFormat.Gif;
        }
        else if (extension == "bmp")
        {
            imageFormat = ImageFormat.Bmp;
        }
        else if (extension == "jpeg")
        {
            imageFormat = ImageFormat.Jpeg;
        }
        else if (extension == "tiff")
        {
            // Convert tiff to gif.
            extension = "gif";
            imageFormat = ImageFormat.Gif;
        }
        else if (extension == "x-wmf")
        {
            extension = "wmf";
            imageFormat = ImageFormat.Wmf;
        }

        // If the image format isn't one that we expect, ignore it,
        // and don't return markup for the link.
        if (imageFormat == null)
        {
            return null;
        }
        
        if (data == null)
        {
            // Копируем данные через мемори стрим из Bitmap в img.Data
            using (var stream = new MemoryStream())
            {
                imageInfo.Bitmap.Save(stream, imageFormat == ImageFormat.Wmf ? ImageFormat.Png : imageFormat);
                data = new byte[stream.Length];
                stream.Position = 0;
                stream.Read(data, 0, data.Length);
            }
        }

        string imageFileName = string.Format("{0}/image{1}.{2}", imageDirectoryName, imageCounter, extension);
        try
        {
            if (data != null)
            {
                File.WriteAllBytes(imageFileName, data);
            }
            else
            {
                imageInfo.Bitmap.Save(imageFileName, imageFormat);
            }
        }
        catch (ExternalException ex)
        {
            Console.WriteLine(ex.Message);
            return null;
        }

        var imageElement = new XElement(Xhtml.img,
            new XAttribute(NoNamespace.src, $"{localDirInfo.Name}/image{imageCounter.ToString()}.{extension}"),
            imageInfo.ImgStyleAttribute,
            imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);

        return imageElement;
    }

    /// <summary>
    /// Проверка наличия заданного атрибута у элемента
    /// </summary>
    private static bool AttributeExists(XElement elem, string name)
    {
        var src = elem.Attribute(name);
        if (src == null || string.IsNullOrWhiteSpace(src.Value))
        {
            return false;
        }

        return true;
    }
}
