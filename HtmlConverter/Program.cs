using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;

namespace HtmlConverter
{
    class Program
    {
        static char charToSkip = (char)0;

        static void Main(string[] args)
        {
            //ConsoleHelpers.ImportFromCsv("E:\\lobby.csv");
            //ParseBuffer("document.html");
            var picDirPath = @"C:\Users\k.komarov\source\example\pic";

            // Настраиваем кэш лобби, чтобы не лезть в базу
            WordImportDal.Lobbies = WordImportDal.GetAllLobbies();

            foreach (var file in Directory.GetFiles(picDirPath, "Руководство пользователя ЕИС (версия 8.2)" + ".docx"))
            {
                ConvertToHtml(file, picDirPath);
            }
        }

        private static void ParseBuffer(string fileName)
        {
            fileName = Path.Combine(@"C:\Users\k.komarov\source\example\clipboard", fileName);
            var htmlDocument = Common.ReadHtmlDocument(fileName);
            TransformHtmlCommon converter = new TransformHtmlCommon();
            var result = converter.PostProcessDocument(htmlDocument);

            result.Save(Path.ChangeExtension(fileName, "xml"));
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

                    var settings = new WmlToHtmlConverterSettings()
                    {
                        AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                        PageTitle = pageTitle,
                        FabricateCssClasses = false,
                        CssClassPrefix = "pt-",
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo => TransformHtmlCommon.ImageHandler(imageInfo, ref imageCounter, imageDirectoryName)
                    };

                    XElement htmlElement = ConsoleHelpers.GetXElement(wDoc, settings, "Конвертация docx -> html...\r\nОбработка контента...");

                    var getCard = new GetNpdDocCard();
                    var body = htmlElement
                        .Elements().First(p => p.Name.LocalName == "body")
                        .Elements().First(p => p.Name.LocalName == "div");

                    // формируем шапку и контент
                    getCard.FormatHeader(body);

                    getCard.FormatAcceptance(body);

                    getCard.FormatGrif(body);

                    // формируем подпись
                    bool signatureExists = getCard.FormatSignature(body);

                    // формируем блок "регистрация в минюсте"
                    getCard.CreateBlockRegistrationMinust(body, signatureExists);

                    ConsoleHelpers.PostProcessAndSave(destFileName, htmlElement);
                    Console.WriteLine();
                    //ConsoleHelpers.ConvertOriginalAndSave(wDoc, destFileName, settings);

                    if (charToSkip != 's')
                    {
                        Console.WriteLine("Нажмите любую клавишу (s = перестать спрашивать)");
                        charToSkip = Console.ReadKey().KeyChar;
                    }
                }
            }
        }
    }
}
