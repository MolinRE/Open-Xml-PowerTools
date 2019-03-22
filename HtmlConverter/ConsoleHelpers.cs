using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace HtmlConverter01
{
    public class ConsoleHelpers
    {
        public static XElement GetXElement(WordprocessingDocument wDoc, HtmlConverterSettings settings, string msg)
        {
            Console.Write(msg);
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            XElement htmlElement = HtmlConverter.ConvertToHtml(wDoc, settings);
            stopwatch.Stop();
            Console.WriteLine(" ({0:0,0} мс.)", stopwatch.ElapsedMilliseconds);
            

            return htmlElement;
        }

        internal static void PostProcessAndSave(FileInfo destFileName, XElement doc)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            if (doc.Descendants(HtmlConverterHelper.XN("img")).Any())
            {
                Console.WriteLine();
            }

            Console.Write("Пост-обработка...");
            stopwatch.Restart();
            var bodyDiv = HtmlConverterHelper.PostProcessDocument(doc);
            stopwatch.Stop();
            Console.WriteLine(" ({0} мс.)", stopwatch.ElapsedMilliseconds);
            var bodyString = bodyDiv.ToString(SaveOptions.OmitDuplicateNamespaces);
            var html = new XDocument(
                new XDocumentType("html", null, null, null),
                new XElement("html",
                    new XElement("head",
                        new XElement("link",
                            new XAttribute("rel", "stylesheet"),
                            new XAttribute("type", "text/css"),
                            new XAttribute("href", "../styles.css"))),
                    new XElement("body", bodyDiv)));
            var htmlString = html.ToString(SaveOptions.OmitDuplicateNamespaces);
            File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
            File.WriteAllText(destFileName.FullName + ".xml", bodyString, Encoding.UTF8);
            Console.WriteLine("Документ \"{0}\" сохранен.", destFileName.Name);
        }

        internal static void ConvertOriginalAndSave(WordprocessingDocument wDoc, FileInfo destFileName, HtmlConverterSettings settings)
        {
            // Produce HTML document with<!DOCTYPE html > declaration to tell the browser
            // we are using HTML5.
            var htmlElement = new XDocument(
                new XDocumentType("html", null, null, null),
                GetXElement(wDoc, settings, "Конвертация docx -> html без пост-обработки...\r\nОбработка контента..."));
            var htmlString = htmlElement.ToString(SaveOptions.None);

            string fileName = Path.GetFileNameWithoutExtension(destFileName.Name) + "_html_converter.html";
            File.WriteAllText(destFileName.DirectoryName + "\\" + fileName, htmlString, Encoding.UTF8);
            Console.WriteLine("Документ \"{0}\" сохранен.", fileName);
        }

        internal static void PrintCard(GetNpdDocCard card)
        {
            string result =
                "------------------------------------------------------------" +
                (card.DocRegions.Any() ? $"\r\n|-Регион: {card.DocRegionString}" : "") +
                (string.IsNullOrEmpty(card.DocName) ? "" : $"\r\n|-Название: {(card.DocName.Length > 300 ? card.DocName.Remove(300) : card.DocName)}") +
                (card.DocTypes.Any() ? $"\r\n|-Тип(ы): {card.DocTypeString}" : "") +
                (card.DocLobbies.Any() ? $"\r\n|-Принявший(-ие) орган(ы): {card.DocLobbyString}" : "") +
                (card.DocDate.HasValue ? $"\r\n|-Дата документа: {card.DocDate:dd.MM.yyyy}" : "") +
                (card.DocNumbers.Any() ? $"\r\n|-Номер(а): {card.DocNumberString}" : "") +
                (string.IsNullOrEmpty(card.RegNumber) ? "" : $"\r\n|-Регистрационный №: {card.RegNumber}") +
                (card.RegDate.HasValue ? $"\r\n|-Дата регистрации: {card.RegDate:dd.MM.yyyy}" : "") +
                "\r\n------------------------------------------------------------";

            Console.WriteLine(result);
        }

        internal static void ImportFromCsv(string csvFilePath)
        {
            var stringBuilder = new StringBuilder();
            //foreach (var type in DocType.Common)
            //{
            //    stringBuilder.AppendFormat("{0};{1};{2}", type.Name, type.NameForRedactions, type.NamePluralForRedactions);
            //    stringBuilder.AppendLine();
            //}

            //File.WriteAllText("E:\\lobby.csv", stringBuilder.ToString(), Encoding.GetEncoding(1251));
            //return;

            foreach (var line in File.ReadLines(csvFilePath, Encoding.GetEncoding(1251)))
            {
                var tokens = line.Split(';');
                stringBuilder.AppendFormat("new DocType() {{ Name = \"{0}\", NameForRedactions = \"{1}\", NamePluralForRedactions = \"{2}\", NameR = \"{3}\" }},",
                    tokens[0], tokens[1], tokens[2], tokens[3]);
                stringBuilder.AppendLine();
            }

            string result = stringBuilder.ToString();
        }
    }
}
