﻿/***************************************************************************

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
using System.Collections.Generic;
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
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using HtmlConverter01;
using OpenXmlPowerTools;
using RestSharp;

class HtmlConverterHelper
{
    const string documentUrlPattern = "document/(?<moduleid>[^/]+)/(?<id>[^/]+)(/(?<anchor>[^/]+))?";

    static void Main(string[] args)
    {
        //ConsoleHelpers.ImportFromCsv("E:\\lobby.csv");
        var picDirPath = @"C:\Users\k.komarov\source\example\list";
        foreach (var file in Directory.GetFiles(picDirPath, "*" + ".docx"))
        {
            ConvertToHtml(file, picDirPath);
        }
    }

    public static void ConvertToHtml(string file, string outputDirectory)
    {
        var fi = new FileInfo(file);
        Console.WriteLine("------------------------------------------------------------");
        Console.WriteLine(fi.Name);
        Console.WriteLine("------------------------------------------------------------");
        byte[] byteArray = File.ReadAllBytes(fi.FullName);
        using (var memoryStream = new MemoryStream())
        {
            memoryStream.Write(byteArray, 0, byteArray.Length);
            // Открываем документ
            using (var wDoc = WordprocessingDocument.Open(memoryStream, true))
            {
                var destFileName = new FileInfo(fi.Name.Replace(".docx", ".html"));
                if (!string.IsNullOrEmpty(outputDirectory))
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
                    pageTitle = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fi.FullName;
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

                XElement htmlElement = ConsoleHelpers.GetXElement(wDoc, settings, "Конвертация docx -> html...\r\nОбработка контента...");

                var getCard = new GetNpdDocCard();
                var body = htmlElement.Element(XN("body")).Element(XN("div"));

                // очищаем формат
                //getCard.ClearFormatBefore(wordDocument);

                // формируем шапку
                getCard.FormatHeader(body);

                // форматируем контент
                //getCard.FormatContent(body);

                getCard.FormatAcceptance(body);

                getCard.FormatGrif(body);

                // формируем подпись
                bool signatureExists = getCard.FormatSignature(body);

                // формируем блок "регистрация в минюсте"
                getCard.CreateBlockRegistrationMinust(body, signatureExists);

                ConsoleHelpers.PostProcessAndSave(destFileName, htmlElement);
                Console.WriteLine();
                ConsoleHelpers.ConvertOriginalAndSave(wDoc, destFileName, settings);

                //Console.WriteLine("Нажмите любую клавишу");
                //Console.ReadKey();
            }
        }
    }
    
    public static XElement PostProcessDocument(XElement html)
    {
        var body = new XElement("div");

        var temp = html.Elements()
                .First(p => p.Name.LocalName == "body").Elements()
                .SelectMany(s => s.Elements())
                .Select(s => s.Value)
                .ToArray();

        // Выбираем элементы внутри дивов
        foreach (var elem in html.Elements()
                .First(p => p.Name.LocalName == "body").Elements()
                .SelectMany(s => s.Elements()))
        {
            // Иногда элементы обернуты дополнительным div-ом
            if (elem.Name.LocalName == "div")
            {
                // Таблицы всегда им обернуты - их оставляем как есть
                if (elem.Elements().Any() && elem.Elements().First().Name.LocalName == "table")
                {
                    body.Add(elem);
                }
                // Иначе добавляем все дочерние элементы
                else
                {
                    foreach (var child in elem.Elements())
                    {
                        body.Add(child);
                    }
                }
            }
            else
            {
                body.Add(elem);
            }
        }

        // Активный список
        XElement list = null;
        // Чтобы сравнивать с предыдущим
        var elements = body.Elements().ToArray();
        for (int i = 0; i < elements.Length; i++)
        {
            var elem = elements[i];
            if (IsListElement(elem, out string listName))
            {
                if (list == null)
                {
                    if (IsNotListStart(body, elem, out string startNumber))
                    {
                        var listNumber = elem.Attribute("abstractNumId").Value;
                        list = body.Descendants().First(p => (p.Name.LocalName == "ul" || p.Name.LocalName == "ol") && p.HasAttributeValue("listNumber", listNumber));

                        int currentIndex = i;
                        var elementsBuffer = new List<XElement>();
                        while (!elements[--currentIndex].HasAttribute("toDelete"))
                        {
                            elementsBuffer.Add(elements[currentIndex]);
                        }
                        elementsBuffer.Reverse();
                        list.Add(elementsBuffer);
                    }
                    else
                    {
                        list = new XElement(listName,
                            new XAttribute("listNumber", elem.Attribute("abstractNumId").Value));
                        if (startNumber != null)
                        {
                            list.Add(new XAttribute("start", startNumber));
                        }

                        var listClass = GetListClass(elem);
                        if (listClass != null)
                        {
                            list.Add(new XAttribute("class", listClass));
                        }
                        elem.AddBeforeSelf(list);
                    }
                }

                // Если элемент описывает список, трансформируем в <li>
                var listItem = TransformToListItemElement(elem, listName);
                if (i > 0 && IsListElement(elements[i - 1]) && IsDifferentLevel(elem, elements[i -1]))
                {
                    if (!IsInnerElement(list, elem, elements[i - 1]))
                    //if (list.Parent.Parent != null && list.Parent.Parent.HasAttributeValue("listNumber", elem.Attribute("abstractNumId").Value))
                    {
                        // Поднимается к <li>, в котором делали список, а затем к ul/ol, в котором элемент находился
                        if (list.Parent.Parent.Name == listName)
                        {
                            list = list.Parent.Parent;
                        }
                        // Иначе это какая-то бажжина (кривое форматирование списка в исходном документе, например)
                        // Пишем в основной список, потому что родительского тупо нет
                    }
                    else
                    {
                        // Создаём ul/ol список
                        var innerList = new XElement(listName,
                            new XAttribute("listNumber", elem.Attribute("abstractNumId").Value));
                        list.Elements().Last().Add(innerList);
                        // Записываем его в общий объект
                        list = innerList;
                    }

                    //// Если является вложенным элементом списка
                    //if (i > 0 && IsInnerListElement(elem, elements[i - 1]))
                    //{
                    //    // Создаём ul/ol список
                    //    var innerList = new XElement(listName,
                    //        new XAttribute("listNumber", elem.Attribute("abstractNumId").Value));
                    //    list.Elements().Last().Add(innerList);
                    //    // Записываем его в общий объект
                    //    list = innerList;
                    //}
                    //// Если является верхним элементом списка
                    //if (i > 0 && IsOuterListElement(elem, elements[i - 1]))
                    //{
                    //    // Поднимается к <li>, в котором делали список, а затем к ul/ol, в котором элемент находился
                    //    if (list.Parent.Parent.Name == listName)
                    //    {
                    //        list = list.Parent.Parent;
                    //    }
                    //    // Иначе это какая-то бажжина (кривое форматирование списка в исходном документе, например)
                    //    // Пишем в основной список, потому что родительского тупо нет
                    //}
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

        body.Elements().Where(p => p.HasAttributeValue("toDelete", "true")).Remove();

        //foreach (var p in body.Elements().ToArray())
        //{
        //    if (IsListElement(p))
        //    {
        //        list.Add(new XElement("li", p));
        //        p.Remove();
        //    }
        //}

        // Удаляем href ссылок
        foreach (var linkElem in body.Descendants(XN("a")).ToArray())
        {
            linkElem.AddAfterSelf(linkElem.Value);
            linkElem.Remove();
        }

        TransformTablesToCke(body);
        TransformParagprahs(body);

        CleanupStyles(body);

        return body;
    }

    private static bool IsNotListStart(XElement body, XElement elem, out string startNumber)
    {
        startNumber = null;
        var listItemRun = elem.Elements().Attributes("listItemRun").First().Value;
        if (listItemRun != "1")
        {
            var abstractNumId = elem.Attribute("abstractNumId").Value;
            if (!body.Descendants().Any(p => (p.Name.LocalName == "ul" || p.Name.LocalName == "ol") && p.HasAttributeValue("listNumber", abstractNumId)))
            {
                startNumber = listItemRun;
                return false;
            }
        }
        return listItemRun != "1";
    }

    private static bool IsInnerElement(XElement list, XElement elem, XElement previous)
    {
        if (list.Parent.Parent == null)
        {
            return true;
        }
        else
        {
            var parentList = list.Parent.Parent;
            if (parentList.HasAttributeValue("listNumber", elem.Attribute("abstractNumId").Value))
            {
                var elementNumber = elem.Elements().Attributes("listItemRun").First().Value;
                var previousNumber = previous.Elements().Attributes("listItemRun").First().Value;
                if (elementNumber.Length > previousNumber.Length && elementNumber.Contains(previousNumber))
                {
                    // Это подпункт
                    return true;
                }

                return false;
            }
            else
            {
                return true;
            }
        }
    }

    private static bool IsDifferentLevel(XElement elem, XElement previous)
    {
        if (elem.Attribute("abstractNumId").Value != previous.Attribute("abstractNumId").Value)
        {
            return true;
        }
        else
        {
            return elem.Elements().Attributes("listItemRun").First().Value.Count(p => p == '.') !=
                previous.Elements().Attributes("listItemRun").First().Value.Count(p => p == '.');
        }
    }

    private static void TransformParagprahs(XElement body)
    {
        foreach (var paraElem in body.Descendants()
            .Where(p => p.Name.LocalName == "p")
            .Where(p => p.HasAttribute("style")))
        {
            string textAlign = paraElem.GetStyle("text-align");

            paraElem.RemoveAttribute("style");
            if (textAlign != null)
            {
                if (textAlign == "justify")
                {
                    textAlign = "left";
                }
                paraElem.SetStyle("text-align", textAlign);
            }
        }
    }

    private static void TransformTablesToCke(XElement bodyDiv)
    {
        foreach (var tableElement in bodyDiv.Descendants().Where(p => p.Name.LocalName == "table"))
        {
            tableElement.Attributes().Where(p => p.Name != "class" && p.Name != "align").Remove();
            // Для CKE: у всех элементов-таблиц должен быть класс cke_show_border
            if (!tableElement.HasClass("cke_show_border"))
            {
                tableElement.AddClass("cke_show_border");
            }

            if (tableElement.Parent.HasAttribute("align") && tableElement.HasAttribute("align"))
            {
                tableElement.Parent.Attribute("align").Remove();
            }

            if (tableElement.HasAttribute("align"))
            {
                tableElement.Parent.Add(tableElement.Attribute("align"));
                tableElement.Attribute("align").Remove();
            }

            var map = BuildTableMap(tableElement);
            var mapValues = new List<List<string>>();
            foreach (var row in map)
            {
                mapValues.Add(new List<string>());
                foreach (var cell in row)
                {
                    mapValues.Last().Add(cell?.Value);
                }
            }

            var CLASSES = new Dictionary<string, string>()
            {
                { "top", "cell_border_top" },
                { "left", "cell_border_left" },
                { "right", "cell_border_right" },
                { "bottom", "cell_border_bottom" },
                { "all", "cell_border_all" }
            };
            var tBorders = new Dictionary<string, bool>()
            {
                { "top", true },
                { "left", true },
                { "right", true },
                { "bottom", true }
            };

            var cells = new List<XElement>();

            for (var rowIndex = 0; rowIndex < map.Count; rowIndex++)
            {
                var row = map[rowIndex];
                for (var colIndex = 0; colIndex < row.Count; colIndex++)
                {
                    var cell = row[colIndex];
                    if (!cells.Contains(cell))
                    {
                        var ckCell = cell;
                        // Удалять стили ckCell не надо - нам они ещё понадобятся
                        string height = row.Element.HasAttribute("height")
                            ? row.Element.Attribute("height").Value
                            : row.Element.GetStyle("height");
                        if (height == null)
                        {
                            height = ckCell.GetStyle("height");
                        }
                        string width = ckCell.HasAttribute("width")
                            ? ckCell.Attribute("width").Value
                            : ckCell.GetStyle("width");

                        string border = ckCell.GetStyle("border");
                        var borders = new Dictionary<string, string>()
                        {
                            { "top", null },
                            { "left", null },
                            { "right", null },
                            { "bottom", null }
                        };

                        foreach (var key in borders.Keys.ToList())
                        {
                            borders[key] = ckCell.GetStyle("border-" + key);

                            string nKey;
                            bool neighbourBorder = false;
                            bool byStyle = false;
                            int rowSpan;
                            int colSpan;
                            int neighborCol;
                            int neighbourRow;

                            switch (key)
                            {
                                case "top":
                                    {
                                        nKey = "bottom";
                                        byStyle = false;
                                        rowSpan = 1;
                                        colSpan = cell.GetColspan(1);
                                        neighborCol = colIndex;
                                        neighbourRow = rowIndex - 1;
                                        break;
                                    }
                                case "bottom":
                                    {
                                        nKey = "top";
                                        byStyle = true;
                                        rowSpan = 1;
                                        colSpan = cell.GetColspan(1);
                                        neighborCol = colIndex;
                                        neighbourRow = rowIndex + cell.GetRowspan(1);
                                        break;
                                    }
                                case "left":
                                    {
                                        nKey = "right";
                                        byStyle = false;
                                        rowSpan = cell.GetRowspan(1);
                                        colSpan = 1;
                                        neighborCol = colIndex - 1;
                                        neighbourRow = rowIndex;
                                        break;
                                    }
                                default: //"right"
                                    {
                                        nKey = "left";
                                        byStyle = true;
                                        rowSpan = cell.GetRowspan(1);
                                        colSpan = 1;
                                        neighborCol = colIndex + cell.GetColspan(1);
                                        neighbourRow = rowIndex;
                                        break;
                                    }
                            }


                            if (neighbourRow >= 0 && neighbourRow < map.Count && neighborCol >= 0 && neighborCol < map[neighbourRow].Count)
                            {
                                for (var i = 0; i < rowSpan; i++)
                                {
                                    var nRow = map[neighbourRow + i];
                                    for (var j = 0; j < colSpan; j++)
                                    {
                                        var neighbour = nRow[neighborCol + j];
                                        if (byStyle)
                                        {
                                            var cellStyle = neighbour.GetStyle("border");
                                            var sideStyle = neighbour.GetStyle("border-" + nKey);
                                            neighbourBorder = neighbourBorder || (sideStyle != "none" && (sideStyle != null || (cellStyle != null && cellStyle != "none")));
                                        }
                                        else
                                        {
                                            neighbourBorder = neighbourBorder || neighbour.HasClass(CLASSES["all"]) || neighbour.HasClass(CLASSES[nKey]);
                                        }
                                    }
                                }

                                if (neighbourBorder)
                                {
                                    // Главное, чтобы было не null
                                    borders[key] = "True";
                                }
                            }

                            // ToString() возвращает значение с большой буквы ("True"/"False")
                            borders[key] = (borders[key] != "none" &&
                                            ((borders[key] != null && borders[key] != "none") ||
                                            (border != null && border != "none"))).ToString();
                        }

                        //if (border == null && border != "none")
                        //{
                        //    if (borders.Any(p => p.Value == "none"))
                        //    {
                        //        border = null;
                        //    }
                        //}

                        //if (border == null)
                        //{
                        //    foreach (var key in borders.Keys.ToArray())
                        //    {
                        //        borders[key] = borders[key] == "none" ? null : borders[key];
                        //    }
                        //}
                        //else
                        //{
                        //    border = null;
                        //    foreach (var key in borders.Keys.ToArray())
                        //    {
                        //        borders[key] = (borders[key] != null && borders[key] != "none") ? borders[key] : null;
                        //    }
                        //}

                        bool borderAll = borders.All(p => p.Value == "True");

                        if (borderAll)//border != null)
                        {
                            ckCell.AddClass(CLASSES["all"]);
                        }
                        else
                        {
                            foreach (var pair in borders)
                            {
                                // TODO: Скорее всего, здесь достаточного только второго условия
                                if (borders[pair.Key] != null && borders[pair.Key] == "True")
                                {
                                    ckCell.AddClass(CLASSES[pair.Key]);
                                }
                            }
                            //if (rowIndex == 0 && (borders["top"] == null || borders["top"] == "False"))
                            //{
                            //    tBorders["top"] = false;
                            //}
                            //if (colIndex == 0 && (borders["left"] == null || borders["left"] == "False"))
                            //{
                            //    tBorders["left"] = false;
                            //}
                            //if (rowIndex == map.Count - 1 && (borders["bottom"] == null || borders["bottom"] == "False"))
                            //{
                            //    tBorders["bottom"] = false;
                            //}
                            //if (colIndex == row.Count - 1 && (borders["right"] == null || borders["right"] == "False"))
                            //{
                            //    tBorders["right"] = false;
                            //}
                        }

                        string verticalAlign = ckCell.GetStyle("vertical-align");
                        string horizontalAlign = ckCell.GetStyle("horizontal-align");

                        ckCell.RemoveAttribute("style");
                        ckCell.RemoveAttribute("width");
                        ckCell.RemoveStyle("height");
                        if (horizontalAlign != null)
                        {
                            ckCell.SetStyle("text-align", horizontalAlign);
                        }
                        if (verticalAlign != null)
                        {
                            ckCell.SetStyle("vertical-align", verticalAlign);
                        }
                        if (width != null)
                        {
                            ckCell.SetStyle("width", ConvertToPx(width));
                        }
                        if (height != null)
                        {
                            ckCell.SetStyle("height", ConvertToPx(height));
                        }
                        cells.Add(ckCell);
                    }
                }
            }

            // На таблицы мы этот класс не вешаем, иначе все внутренние границы тоже станут видимымит
            //var tAll = tBorders["top"] && tBorders["left"] && tBorders["right"] && tBorders["bottom"];
            //if (tAll)
            //{
            //    tableElement.AddClass(CLASSES["all"]);
            //}
            //else
            //{
            //    foreach (var keyPair in tBorders)
            //    {
            //        if (tBorders[keyPair.Key])
            //        {
            //            tableElement.AddClass(CLASSES[keyPair.Key]);
            //        }
            //    }
            //}


            var tableContent = tableElement.Elements().ToArray();
            var tbody = new XElement("tbody", tableContent);
            tableElement.Add(tbody);
            tableContent.Remove();
        }
    }

    private static string ConvertToPx(string style)
    {
        string pattern = @"([\d.,]+)\s?([a-zA-Z%]*)";
        var match = Regex.Match(style, pattern, RegexOptions.Compiled);
        if (match.Success)
        {
            var value = float.Parse(match.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture);
            if (match.Groups[2].Value == "pt")
            {
                value *= (float)1.33;
            }

            return value.ToString(CultureInfo.InvariantCulture) + "px";
        }

        return style;
    }

    private static XTable BuildTableMap(XElement table)//, int? startRow = null, int? startCell = null, int? endRow = null, int? endCell = null)
    {
        var tableRows = table.Elements(XN("tr")).ToList();

        var startRowIndex = 0;
        var startCellIndex = 0;
        var endRowIndex = tableRows.Count - 1;
        var endCellIndex = -1;

        // Row and Column counters.
        int rowCounter = -1;
        int columnCounter;

        var result = new XTable();

        for (var rowIndex = startRowIndex; rowIndex <= endRowIndex; rowIndex++)
        {
            rowCounter++;
            Trace.Write($"{rowCounter}. ");

            //!aMap[r] && (aMap[r] = []);
            //if (result.Count <= rowCounter)
            //{
            //    result.Add(new XList());
            //}
            if (result[rowCounter] == null)
            {
                result[rowCounter] = new XRow();
            }

            columnCounter = -1;

            for (var colIndex = startCellIndex; colIndex <= (endCellIndex == -1 ? (tableRows[rowIndex].Elements(XN("td")).Count() - 1) : endCellIndex); colIndex++)
            {
                var currentCell = tableRows[rowIndex].Elements(XN("td")).ElementAt(colIndex);
                Trace.Write(currentCell.Value + "; ");

                if (currentCell == null)
                {
                    break;
                }

                columnCounter++;
                //while (aMap[r][c])
                //    c++;
                //while (result[rowCounter].Count > columnCounter) //&& 
                while (result[rowCounter][columnCounter] != null)
                {
                    columnCounter++;
                }

                var colSpan = //isNaN( oCell.colSpan ) ? 1 : oCell.colSpan;
                    currentCell.GetColspan(1);
                //!int.TryParse(currentCell.Attribute("colspan") == null ? null : currentCell.Attribute("colspan").Value, out int n1) ? 1 : n1;
                var rowSpan = //isNaN(oCell.rowSpan) ? 1 : oCell.rowSpan;
                    currentCell.GetRowspan(1);
                    //!int.TryParse(currentCell.Attribute("rowspan") == null ? null : currentCell.Attribute("rowspan").Value, out int n2) ? 1 : n2;
                    

                // Если есть роуспан, копируем ячейку во все последующие строки
                for (var rs = 0; rs < rowSpan; rs++)
                {
                    if (rowIndex + rs > endRowIndex)
                    {
                        break;
                    }

                    //if (!aMap[r + rs])
                    //    aMap[r + rs] = [];
                    //if (result.Count <= rowCounter + rs)
                    //{
                    //    result.Add(new XList());
                    //}

                    if (result[rowCounter + rs] == null)
                    {
                        result[rowCounter + rs] = new XRow();
                    }

                    // Если есть колспан, копируем ячейку во все последущие столбцы
                    // (на каждой строке, которая есть в цикле роуспана)
                    for (var cs = 0; cs < colSpan; cs++)
                    {
                        // Неправильно
                        //while (result[rowCounter + rs].Count - 1 < columnCounter + cs)
                        //{
                        //    result[rowCounter + rs].Add(null);
                        //}

                        //if (result[rowCounter + rs].Count <= columnCounter + cs)
                        //{
                        //    result[rowCounter + rs].Add(tableRows[i].Elements(XN("td")).ElementAt(j));
                        //}
                        //else
                        //{
                            result[rowCounter + rs][columnCounter + cs] = tableRows[rowIndex].Elements(XN("td")).ElementAt(colIndex);
                        //}
                    }
                }

                columnCounter += colSpan - 1;

                if (endCellIndex != -1 && columnCounter >= endCellIndex)
                {
                    break;
                }
            }
            Trace.WriteLine("");
        }
        return result;
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
                if (span.Value.StartsWith("12. В случае несоблюдения"))
                {

                }

                var spanStyle = span.HasAttribute("style") ? span.Attribute("style").Value.Split(';') : new string[0];
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
                    // Если спан пустой ИЛИ пустой, но при этом после него есть слово
                    if (spanWrap.Value != " " ||
                        (spanWrap.Value == " " && span.ElementsAfterSelf().FirstOrDefault() != null))
                    {
                        var previousElement = span.ElementsBeforeSelf().LastOrDefault();
                        if (previousElement != null && previousElement.Name == spanWrap.Name)
                        {
                            previousElement.SetValue(previousElement.Value + spanWrap.Value);
                            spanWrap = previousElement;
                        }
                        else
                        {
                            span.AddAfterSelf(spanWrap);
                        }
                        previousFormattedSpan = spanWrap;
                    }

                }
                span.Remove();
            }
            else
            {
                // Считаем, что если внутри только один элемент, то всё ок
                // При этом игнорируем элементы "<span />" - иногда либа генерирует их
                if (span.Elements().Count(p => !(p.Name.LocalName == "span" && !p.Nodes().Any())) == 1)
                {
                    var node = span.Elements().First(p => !(p.Name.LocalName == "span" && !p.Nodes().Any()));
                    span.AddAfterSelf(node);
                    span.Remove();
                }
            }
        }
    }

    private static XElement TransformToListItemElement(XElement elem, string listName)
    {
        var spanElements = elem.Elements(XN("span")).ToArray();
        XElement result;
        // Если это нумерованный список и нет класса - это п
        if (listName == "ol" && GetListClass(elem) == null)
        {
            result = new XElement("p",
                new XAttribute("elementNumber", elem.Elements().Attributes("listItemRun").First().Value));
            if (spanElements.Length > 1)
            {
                RemoveSpans(spanElements.Skip(1).ToArray());
                result.Add(elem.Elements().First().Value + " ");
                foreach (var node in elem.Nodes().Skip(1))
                {
                    result.Add(node);
                }
            }
        }
        else
        {
            result = new XElement("li",
                new XAttribute("elementNumber", elem.Elements().Attributes("listItemRun").First().Value));
            if (spanElements.Length > 1)
            {
                // Первый спан - иконка списка
                RemoveSpans(spanElements.Skip(1).ToArray());
                // При этом иконку важно сохранить для дальнейшего анализа вложенности
                // TODO: Уже не важно
                foreach (var node in elem.Nodes().Skip(1))
                {
                    result.Add(node);
                }
            }
        }


        return result;
    }

    public static bool IsInnerListElement(XElement pElement, XElement previousElement)
    {
        // p abstractNumId="2" 
        // span listItemRun="1" 

        if (pElement.Attribute("abstractNumId").Value == previousElement.Attribute("abstractNumId").Value)
        {
            return false;
        }
        else
        {
            if (pElement.Elements().Attributes("listItemRun").First().Value == "1")
            {
                return true;
            }

            return false;
        }
        
        ////margin-left: 1.00in;
        //string marginLeft = pElement.GetStyle("margin-left");
        //if (marginLeft == null)
        //{
        //    return false;
        //}
        //if (marginLeft.EndsWith("in"))
        //{
        //    marginLeft = marginLeft.Remove(marginLeft.Length - 2);
        //}

        //string marginLeftPrevious = previousElement.GetStyle("margin-left");
        //if (marginLeft == null)
        //{
        //    return false;
        //}
        //if (marginLeftPrevious.EndsWith("in"))
        //{
        //    marginLeftPrevious = marginLeftPrevious.Remove(marginLeftPrevious.Length - 2);
        //}

        //// InvarianCulture, т.к. в российской культуре дробная часть отделяется запятой
        //return Convert.ToSingle(marginLeft, CultureInfo.InvariantCulture) > Convert.ToSingle(marginLeftPrevious, CultureInfo.InvariantCulture);
    }

    public static bool IsOuterListElement(XElement pElement, XElement previousElement)
    {
        if (pElement.Attribute("abstractNumId").Value == previousElement.Attribute("abstractNumId").Value)
        {
            return false;
        }
        else
        {
            if (pElement.Elements().Attributes("listItemRun").First().Value == "1")
            {
                return false;
            }

            return true;
        }

        ////margin-left: 1.00in;
        //string marginLeft = pElement.GetStyle("margin-left");
        //if (marginLeft == null)
        //{
        //    return false;
        //}
        //if (marginLeft.EndsWith("in"))
        //{
        //    marginLeft = marginLeft.Remove(marginLeft.Length - 2);
        //}

        //string marginLeftPrevious = previousElement.GetStyle("margin-left");
        //if (marginLeft == null)
        //{
        //    return false;
        //}
        //if (marginLeftPrevious.EndsWith("in"))
        //{
        //    marginLeftPrevious = marginLeftPrevious.Remove(marginLeftPrevious.Length - 2);
        //}

        //// InvarianCulture, т.к. в российской культуре дробная часть отделяется запятой
        //return Convert.ToSingle(marginLeft, CultureInfo.InvariantCulture) < Convert.ToSingle(marginLeftPrevious, CultureInfo.InvariantCulture);
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
        if (elem.Name.LocalName != "p")
        {
            return false;
        }
        // В этом спане обычно содержится 
        var spanElement = elem.Elements().FirstOrDefault(p => p.Name.LocalName == "span");
        if (spanElement == null)
        {
            return false;
        }


        if (spanElement.HasAttribute("style"))
        {
            listName = spanElement.Value.Length == 1 ? "ul" : "ol";

            //return spanElement.GetStyle("display") == "inline-block" &&
            //    spanElement.GetStyle("text-indent") == "0" &&
            //    spanElement.HasStyle("width") &&
            //    spanElement.Value.Length < 8;
        }

        //if (listName != "ul" && GetListClass(elem) == null)
        //{
        //    // Чтобы не склеивались
        //    spanElement.SetValue(spanElement.Value + " ");
        //    return false;
        //}

        return elem.Elements().Any(p => p.HasAttribute("listItemRun"));
        //return spanElement.Value == "" || spanElement.Value == ""
        //    || (spanElement.Value.Length > 1 && int.TryParse(spanElement.Value.Remove(spanElement.Value.Length - 1), out int listIndex));
    }

    public static string GetListClass(XElement elem)
    {
        var span = elem.Elements().FirstOrDefault(p => p.Name.LocalName == "span");
        if (span == null)
        {
            return null;
        }
        
        var upperAlphaRegex = new Regex(@"[A-Z]+\.", RegexOptions.Compiled);
        var lowerAlphaRegex = new Regex(@"[a-z]+\.", RegexOptions.Compiled);
        var numericRegex = new Regex(@"[0-9]+\.", RegexOptions.Compiled);

        if (upperAlphaRegex.IsMatch(span.Value))
        {
            return "c-list upper-latin";
        }
        else if (lowerAlphaRegex.IsMatch(span.Value))
        {
            return "c-list lower-latin";
        }
        else if (numericRegex.IsMatch(span.Value))
        {
            return "c-list normal";
        }
        else
        {
            // Это список с маркировкой, который нет в веб-арме
            return null;
        }
    }

    public static bool IsListElement(XElement pElement)
    {
        string listType;
        return IsListElement(pElement, out listType);
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
            if (elem.Name == "table" || elem.Name == "img" || elem.Name == "td")
            {
                continue;
            }
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
                if (elem.Name != "incw" && elem.Name != "ol")
                {
                    badAttribute.Remove();
                }
            }
            badAttribute = elem.Attribute("style");
            if (badAttribute != null)
            {
                if (elem.Name != "h3" && elem.Name != "p")
                {
                    badAttribute.Remove();
                }
            }

            badAttribute = elem.Attribute("elementNumber");
            if (badAttribute != null)
            {
                badAttribute.Remove();
            }

            badAttribute = elem.Attribute("listNumber");
            if (badAttribute != null)
            {
                badAttribute.Remove();
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
        Console.Write("\rОбработка картинок... {0,5}", imageCounter);

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
        switch (extension)
        {
            //case "x-wmf":
            case "png":
                extension = "png";
                imageFormat = ImageFormat.Png;
                break;
            case "gif":
                imageFormat = ImageFormat.Gif;
                break;
            case "bmp":
                imageFormat = ImageFormat.Bmp;
                break;
            case "jpeg":
                imageFormat = ImageFormat.Jpeg;
                break;
            case "tiff":
                // Convert tiff to gif
                extension = "gif";
                imageFormat = ImageFormat.Gif;
                break;
            case "x-wmf":
                extension = "wmf";
                imageFormat = ImageFormat.Wmf;
                break;
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

        if (imageInfo.ImgStyleAttribute == null)
        {
            var shapeElement = imageInfo.DrawingElement.Element(VML.shape);
            if (shapeElement != null)
            {
                imageInfo.ImgStyleAttribute = shapeElement.Attribute("style");
            }
        }

        var imageElement = new XElement("img",
            new XAttribute("src", $"{localDirInfo.Name}/image{imageCounter.ToString()}.{extension}"),
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
