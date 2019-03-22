using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using HtmlConverter01;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlPowerTools;

namespace WebArmTests
{
    [TestClass]
    public class GetNpdCardTests
    {
        const string docNamesAreNotEqual = "Название документа не совпадает";
        const string docNameIsNotEmpty = "Название содержит текст";

        const string docLobbyCountAreNotEqual = "Количество органов не совпадает";
        const string docLobbiesAreNotEqual = "Название органа не совпадает";

        const string docDateIsNull = "Дата документа пустая";
        const string docDateIsNotNull = "Дата документа не пустая";
        const string docDatesAreNotEqual = "Дата документа не совпадает";

        const string docNumbersCountAreNotEqual = "Количество номеров не совпадает";
        const string docNumbersAreNotEqual = "Номер документа не совпадает";

        const string docCaseNumberIsNotEmpty = "Номер дела не пустой";
        const string docCaseNumbersAreNotEqual = "Номер дела не совпадает";

        const string docTypesCountAreNotEqual = "";
        const string docTypesAreNotEqual = "";

        const string docVersionDateIsNull = "Дата редакции пустая";
        const string docVersionDateIsNotNull = "Дата редакции не пустая";
        const string docVersionDatesAreNotEqual = "Дата редакции не совпадает";

        const string regDatesAreNotEqual = "Дата регистрации не совпадает";
        const string regDateIsNull = "Дата регистрации пустая";
        const string regDateIsNotNull = "Дата регистрации не пустая";
        const string regNumbersAreNotEqual = "Регистрационный № документа не совпадает";
        const string regNumberIsNotEmpty = "Регистрационный № документа содержит текст";

        private static GetNpdDocCard ConvertToHtml(string fileName)
        {
            var fi = new FileInfo(fileName);
            Console.WriteLine(fi.Name);
            byte[] byteArray = File.ReadAllBytes(fi.FullName);
            using (var memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                // Открываем документ
                using (var wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
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
                        ImageHandler = imageInfo => ImageHandler(imageInfo, ref imageCounter)
                    };

                    XElement htmlElement = HtmlConverter01.ConsoleHelpers.GetXElement(wDoc, settings, "Конвертация docx -> html...\r\nОбработка контента...");

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

                    return getCard;
                }
            }
        }

        public static XElement ImageHandler(ImageInfo imageInfo, ref int imageCounter)
        {
            var imageElement = new XElement(Xhtml.img,
                new XAttribute(NoNamespace.src, $"image-src"),
                imageInfo.ImgStyleAttribute,
                imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);

            return imageElement;
        }

        const string documentsDirectory = @"C:\Users\k.komarov\source\ad\MolinRE_Open-Xml-PowerTools\WebArmTests\SampleDocuments";

        [TestMethod]
        public void GetCardForStandard()
        {
            string fileName = Path.Combine(documentsDirectory, "пример эталон.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual("Об утверждении форм, содержания и порядка " +
                "представления отчетности об осуществлении органами государственной власти " +
                "субъектов российской федерации переданных полномочий российской федерации в области лесных отношений", 
                docCard.DocName, docNamesAreNotEqual);

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(1, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual);
            Assert.AreEqual("Министерство природных ресурсов и экологии РФ", docCard.DocLobbies[0].Name, docLobbiesAreNotEqual);

            Assert.IsTrue(docCard.DocDate.HasValue, docDateIsNull);
            Assert.AreEqual(new DateTime(2015, 12, 28).Date, docCard.DocDate.Value.Date, docDatesAreNotEqual);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(1, docCard.DocNumbers.Count, docNumbersCountAreNotEqual);
            Assert.AreEqual("565", docCard.DocNumbers[0], docNumbersAreNotEqual);

            Assert.AreEqual(string.Empty, docCard.DocCaseNumber, docCaseNumberIsNotEmpty);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(1, docCard.DocTypes.Count, docTypesCountAreNotEqual);
            Assert.AreEqual("ПРИКАЗ", docCard.DocTypes[0].Name, docTypesAreNotEqual);

            Assert.IsTrue(docCard.DocVersionDate.HasValue, docVersionDateIsNull);
            Assert.AreEqual(new DateTime(2017, 4, 3).Date, docCard.DocVersionDate.Value.Date, docVersionDatesAreNotEqual);

            Assert.IsTrue(docCard.RegDate.HasValue, regDateIsNull);
            Assert.AreEqual(new DateTime(2016, 3, 25).Date, docCard.RegDate.Value.Date, regDatesAreNotEqual);

            Assert.AreEqual("41569", docCard.RegNumber, regNumbersAreNotEqual);
        }

        [TestMethod]
        public void GetCardForExample1()
        {
            string fileName = Path.Combine(documentsDirectory, "пример1 приказ ФНС.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual(docCard.DocName, "О вводе в промышленную эксплуатацию программного обеспечения задачи " +
                "\"передача в банки документов, используемых налоговыми органами при реализации своих полномочий в " +
                "отношениях, регулируемых законодательством о налогах и сборах, и представление банками информации в " +
                "налоговые органы в электронном виде по телекоммуникационным каналам связи\" (\"банк-обмен\")", docNamesAreNotEqual);

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(1, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual);
            Assert.AreEqual(docCard.DocLobbies[0].Name, "Федеральная налоговая служба");

            Assert.IsTrue(docCard.DocDate.HasValue);
            Assert.AreEqual(docCard.DocDate.Value.Date, new DateTime(2011, 11, 29).Date);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(docCard.DocNumbers.Count, 1);
            Assert.AreEqual(docCard.DocNumbers[0], "ММВ-7-6/901@");

            Assert.AreEqual(string.Empty, docCard.DocCaseNumber, docCaseNumberIsNotEmpty);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(docCard.DocTypes.Count, 1);
            Assert.AreEqual(docCard.DocTypes[0].Name, "ПРИКАЗ");

            Assert.IsTrue(docCard.DocVersionDate.HasValue, docVersionDateIsNull);
            Assert.AreEqual(docCard.DocVersionDate.Value.Date, new DateTime(2014, 12, 25).Date, docVersionDatesAreNotEqual);

            Assert.IsTrue(docCard.RegDate.HasValue, regDateIsNull);
            Assert.AreEqual(docCard.RegDate.Value.Date, new DateTime(2016, 3, 25).Date, regDatesAreNotEqual);

            Assert.AreEqual(docCard.RegNumber, "41569", regNumbersAreNotEqual);
        }

        [TestMethod]
        public void GetCardForExample2()
        {
            string fileName = Path.Combine(documentsDirectory, "пример2 приказ ФНС неск изм.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual(docCard.DocName, "Об утверждении формы и формата заявления о применении " +
                "налоговой льготы участниками региональных инвестиционных проектов, для которых не " +
                "требуется включение в реестр участников региональных инвестиционных проектов, а " +
                "также порядка его передачи в электронной форме по телекоммуникационным каналам связи (образец)", docNamesAreNotEqual);

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(1, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual);
            Assert.AreEqual(docCard.DocLobbies[0].Name, "Федеральная налоговая служба");

            Assert.IsTrue(docCard.DocDate.HasValue);
            Assert.AreEqual(docCard.DocDate.Value.Date, new DateTime(2016, 12, 27).Date);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(docCard.DocNumbers.Count, 1);
            Assert.AreEqual(docCard.DocNumbers[0], "ММВ-7-3/719@");

            Assert.AreEqual(string.Empty, docCard.DocCaseNumber, docCaseNumberIsNotEmpty);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(docCard.DocTypes.Count, 1);
            Assert.AreEqual(docCard.DocTypes[0].Name, "ПРИКАЗ");

            Assert.IsTrue(docCard.DocVersionDate.HasValue, docVersionDateIsNull);
            Assert.AreEqual(docCard.DocVersionDate.Value.Date, new DateTime(2018, 10, 19).Date, docVersionDatesAreNotEqual);

            Assert.IsTrue(docCard.RegDate.HasValue, regDateIsNull);
            Assert.AreEqual(docCard.RegDate.Value.Date, new DateTime(2017, 01, 24).Date, regDatesAreNotEqual);

            Assert.AreEqual(docCard.RegNumber, "45366", regNumbersAreNotEqual);
        }

        [TestMethod]
        public void GetCardForExample3()
        {
            string fileName = Path.Combine(documentsDirectory, "пример3 два органа два номера.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual("Об утверждении состава и порядка представления федеральной " +
                "налоговой службой и федеральной таможенной службой сведений, предусмотренных подпунктом " +
                "1.1 пункта 1 статьи 151 налогового кодекса российской федерации", docCard.DocName, docNamesAreNotEqual);

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(2, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual);
            Assert.AreEqual("Министерство финансов РФ", docCard.DocLobbies[0].Name);
            Assert.AreEqual("Федеральная налоговая служба", docCard.DocLobbies[1].Name);

            Assert.IsTrue(docCard.DocDate.HasValue);
            Assert.AreEqual(new DateTime(2016, 5, 25).Date, docCard.DocDate.Value.Date);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(2, docCard.DocNumbers.Count, docNumbersCountAreNotEqual);
            Assert.AreEqual("72н", docCard.DocNumbers[0]);
            Assert.AreEqual("ММВ-7-15/335@", docCard.DocNumbers[1]);

            Assert.AreEqual(string.Empty, docCard.DocCaseNumber, docCaseNumberIsNotEmpty);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(1, docCard.DocTypes.Count);
            Assert.AreEqual("ПРИКАЗ", docCard.DocTypes[0].Name);

            Assert.IsFalse(docCard.DocVersionDate.HasValue, docVersionDateIsNotNull);

            Assert.IsTrue(docCard.RegDate.HasValue, regDateIsNull);
            Assert.AreEqual(new DateTime(2016, 6, 28).Date, docCard.RegDate.Value.Date, regDatesAreNotEqual);

            Assert.AreEqual("42660", docCard.RegNumber, regNumbersAreNotEqual);
        }

        [TestMethod]
        public void GetCardForExample4()
        {
            string fileName = Path.Combine(documentsDirectory, "пример4 три органа три номера.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual("О реализации положений постановления правительства российской " +
                "федерации от 21 октября 2004 г. n 573 \"о порядке и условиях финансирования процедур банкротства " +
                "отсутствующих должников\"", docCard.DocName, docNamesAreNotEqual);

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(3, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual);
            Assert.AreEqual("Федеральная налоговая служба", docCard.DocLobbies[0].Name);
            Assert.AreEqual("Министерство экономического развития и торговли РФ", docCard.DocLobbies[1].Name);
            Assert.AreEqual("Министерство финансов РФ", docCard.DocLobbies[2].Name);

            Assert.IsTrue(docCard.DocDate.HasValue);
            Assert.AreEqual(new DateTime(2005, 3, 10).Date, docCard.DocDate.Value.Date);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(docCard.DocNumbers.Count, 3);
            Assert.AreEqual("САЭ-3-19/80@", docCard.DocNumbers[0]);
            Assert.AreEqual("53", docCard.DocNumbers[1]);
            Assert.AreEqual("34н", docCard.DocNumbers[2]);

            Assert.AreEqual(string.Empty, docCard.DocCaseNumber, docCaseNumberIsNotEmpty);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(1, docCard.DocTypes.Count);
            Assert.AreEqual("ПРИКАЗ", docCard.DocTypes[0].Name);

            Assert.IsFalse(docCard.DocVersionDate.HasValue, docVersionDateIsNotNull);

            Assert.IsTrue(docCard.RegDate.HasValue, regDateIsNull);
            Assert.AreEqual(new DateTime(2005, 4, 18).Date, docCard.RegDate.Value.Date, regDatesAreNotEqual);

            Assert.AreEqual("6516", docCard.RegNumber, regNumbersAreNotEqual);
        }

        [TestMethod]
        public void GetCardForExample5()
        {
            string fileName = Path.Combine(documentsDirectory, "пример5 суд с делом.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual(docCard.DocName, string.Empty, docNameIsNotEmpty);

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(1, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual);
            Assert.AreEqual("Арбитражный суд Уральского округа", docCard.DocLobbies[0].Name);

            Assert.IsTrue(docCard.DocDate.HasValue);
            Assert.AreEqual(docCard.DocDate.Value.Date, new DateTime(2018, 10, 23).Date);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(docCard.DocNumbers.Count, 1);
            Assert.AreEqual(docCard.DocNumbers[0], "Ф09-7404/2017");

            Assert.AreEqual("А60-16021/2017", docCard.DocCaseNumber, docCaseNumbersAreNotEqual);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(docCard.DocTypes.Count, 1);
            Assert.AreEqual(docCard.DocTypes[0].Name, "ПОСТАНОВЛЕНИЕ");

            Assert.IsFalse(docCard.DocVersionDate.HasValue, docVersionDateIsNotNull);

            Assert.IsFalse(docCard.RegDate.HasValue, regDateIsNotNull);

            Assert.AreEqual(docCard.RegNumber, string.Empty, regNumberIsNotEmpty);
        }

        [TestMethod]
        public void GetCardForExample6()
        {
            string fileName = Path.Combine(documentsDirectory, "пример6 без названия.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual(docCard.DocName, string.Empty, docNameIsNotEmpty);

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(1, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual);
            Assert.AreEqual(docCard.DocLobbies[0].Name, "Министерство финансов РФ");

            Assert.IsTrue(docCard.DocDate.HasValue);
            Assert.AreEqual(docCard.DocDate.Value.Date, new DateTime(2017, 12, 13).Date);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(docCard.DocNumbers.Count, 1);
            Assert.AreEqual(docCard.DocNumbers[0], "02-07-07/83463");

            Assert.AreEqual(string.Empty, docCard.DocCaseNumber, docCaseNumberIsNotEmpty);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(docCard.DocTypes.Count, 1);
            Assert.AreEqual(docCard.DocTypes[0].Name, "ПИСЬМО");

            Assert.IsFalse(docCard.DocVersionDate.HasValue, docVersionDateIsNotNull);

            Assert.IsFalse(docCard.RegDate.HasValue, regDateIsNotNull);

            Assert.AreEqual(docCard.RegNumber, string.Empty, regNumberIsNotEmpty);
        }

        [TestMethod]
        public void GetCardForExample7a()
        {
            string fileName = Path.Combine(documentsDirectory, "пример7 - разные изменяющие1.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual(docCard.DocName, "Об утверждении формы и формата заявления о применении налоговой " +
                "льготы участниками региональных инвестиционных проектов, для которых не требуется включение " +
                "в реестр участников региональных инвестиционных проектов, а также порядка его передачи в " +
                "электронной форме по телекоммуникационным каналам связи", docNamesAreNotEqual);

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(1, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual);
            Assert.AreEqual(docCard.DocLobbies[0].Name, "Федеральная налоговая служба");

            Assert.IsTrue(docCard.DocDate.HasValue);
            Assert.AreEqual(docCard.DocDate.Value.Date, new DateTime(2016, 12, 27).Date);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(docCard.DocNumbers.Count, 1);
            Assert.AreEqual(docCard.DocNumbers[0], "ММВ-7-3/719@");

            Assert.AreEqual(string.Empty, docCard.DocCaseNumber, docCaseNumberIsNotEmpty);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(docCard.DocTypes.Count, 1);
            Assert.AreEqual(docCard.DocTypes[0].Name, "ПРИКАЗ");

            Assert.IsTrue(docCard.DocVersionDate.HasValue, docVersionDateIsNull);
            Assert.AreEqual(docCard.DocVersionDate.Value.Date, new DateTime(2018, 10, 21).Date, docVersionDatesAreNotEqual);

            Assert.IsTrue(docCard.RegDate.HasValue, regDateIsNull);
            Assert.AreEqual(docCard.RegDate.Value.Date, new DateTime(2017, 1, 24).Date, regDatesAreNotEqual);

            Assert.AreEqual(docCard.RegNumber, "45366", regNumbersAreNotEqual);
        }

        [TestMethod]
        public void GetCardForExample7b()
        {
            string fileName = Path.Combine(documentsDirectory, "пример7 - разные изменяющие2.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual(docCard.DocName, "Об утверждении формы и формата заявления о применении налоговой " +
                "льготы участниками региональных инвестиционных проектов, для которых не требуется включение " +
                "в реестр участников региональных инвестиционных проектов, а также порядка его передачи в " +
                "электронной форме по телекоммуникационным каналам связи", docNamesAreNotEqual);

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(1, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual);
            Assert.AreEqual(docCard.DocLobbies[0].Name, "Федеральная налоговая служба");

            Assert.IsTrue(docCard.DocDate.HasValue);
            Assert.AreEqual(docCard.DocDate.Value.Date, new DateTime(2016, 12, 27).Date);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(docCard.DocNumbers.Count, 1);
            Assert.AreEqual(docCard.DocNumbers[0], "ММВ-7-3/719@");

            Assert.AreEqual(string.Empty, docCard.DocCaseNumber, docCaseNumberIsNotEmpty);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(docCard.DocTypes.Count, 1);
            Assert.AreEqual(docCard.DocTypes[0].Name, "ПРИКАЗ");

            Assert.IsTrue(docCard.DocVersionDate.HasValue, docVersionDateIsNull);
            Assert.AreEqual(docCard.DocVersionDate.Value.Date, new DateTime(2018, 10, 21).Date, docVersionDatesAreNotEqual);

            Assert.IsTrue(docCard.RegDate.HasValue, regDateIsNull);
            Assert.AreEqual(docCard.RegDate.Value.Date, new DateTime(2017, 1, 24).Date, regDatesAreNotEqual);

            Assert.AreEqual(docCard.RegNumber, "45366", regNumbersAreNotEqual);
        }

        [TestMethod]
        public void GetCardForExample7c()
        {
            string fileName = Path.Combine(documentsDirectory, "пример7 суд без дела.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual(docCard.DocName, string.Empty, docNameIsNotEmpty);

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(1, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual);
            Assert.AreEqual(docCard.DocLobbies[0].Name, "Арбитражный суд Нижегородской области");

            Assert.IsTrue(docCard.DocDate.HasValue);
            Assert.AreEqual(docCard.DocDate.Value.Date, new DateTime(2015, 05, 19).Date);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(docCard.DocNumbers.Count, 1);
            Assert.AreEqual(docCard.DocNumbers[0], "А43-3580/2015");

            Assert.AreEqual(string.Empty, docCard.DocCaseNumber, docCaseNumberIsNotEmpty);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(docCard.DocTypes.Count, 1);
            Assert.AreEqual(docCard.DocTypes[0].Name, "ОПРЕДЕЛЕНИЕ");

            Assert.IsFalse(docCard.DocVersionDate.HasValue, docVersionDateIsNotNull);

            Assert.IsFalse(docCard.RegDate.HasValue, regDateIsNotNull);

            Assert.AreEqual(docCard.RegNumber, string.Empty, regNumberIsNotEmpty);
        }

        [TestMethod]
        public void GetCardForExample8a()
        {
            string fileName = Path.Combine(documentsDirectory, "пример 8 - АПЕЛЛЯЦИЯ.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual(docCard.DocName, "О возвращении кассационной жалобы", docNamesAreNotEqual);

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(1, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual);
            Assert.AreEqual(docCard.DocLobbies[0].Name, "Федеральный арбитражный суд Северо-Западного округа");

            Assert.IsTrue(docCard.DocDate.HasValue);
            Assert.AreEqual(docCard.DocDate.Value.Date, new DateTime(2014, 07, 07).Date);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(docCard.DocNumbers.Count, 1);
            Assert.AreEqual(docCard.DocNumbers[0], "А56-6885/2014");

            Assert.AreEqual(string.Empty, docCard.DocCaseNumber, docCaseNumberIsNotEmpty);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(docCard.DocTypes.Count, 1);
            Assert.AreEqual(docCard.DocTypes[0].Name, "ОПРЕДЕЛЕНИЕ");

            Assert.IsFalse(docCard.DocVersionDate.HasValue, docVersionDateIsNotNull);

            Assert.IsFalse(docCard.RegDate.HasValue, regDateIsNotNull);

            Assert.AreEqual(docCard.RegNumber, string.Empty, regNumberIsNotEmpty);
        }

        [TestMethod]
        public void GetCardForExample8b()
        {
            string fileName = Path.Combine(documentsDirectory, "пример 8 - КАССАЦИЯ.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual(docCard.DocName, "О возвращении кассационной жалобы");

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(1, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual);
            Assert.AreEqual(docCard.DocLobbies[0].Name, "Федеральный арбитражный суд Северо-Западного округа");

            Assert.IsTrue(docCard.DocDate.HasValue);
            Assert.AreEqual(docCard.DocDate.Value.Date, new DateTime(2014, 07, 07).Date);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(docCard.DocNumbers.Count, 1);
            Assert.AreEqual(docCard.DocNumbers[0], "А56-6885/2014");

            Assert.AreEqual(string.Empty, docCard.DocCaseNumber, docCaseNumberIsNotEmpty);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(docCard.DocTypes.Count, 1);
            Assert.AreEqual(docCard.DocTypes[0].Name, "ОПРЕДЕЛЕНИЕ");
            
            Assert.IsFalse(docCard.DocVersionDate.HasValue, docVersionDateIsNotNull);

            Assert.IsFalse(docCard.RegDate.HasValue, regDateIsNotNull);

            Assert.AreEqual(docCard.RegNumber, string.Empty, regNumberIsNotEmpty);
        }

        [TestMethod]
        public void GetCardForExample9()
        {
            string fileName = Path.Combine(documentsDirectory, "пример9 АОФР изменяющий hak_537.docx");
            var docCard = ConvertToHtml(fileName);

            Assert.AreEqual("О внесении изменения в приложение к постановлению Правительства " +
                "Республики Хакасия от 01.02.2018 № 39 «Об утверждении распределения субсидий из республиканского " +
                "бюджета Республики Хакасия бюджетам муниципальных образований на реализацию мероприятий, направленных " +
                "на поддержку и развитие систем коммунального комплекса в муниципальных образованиях Республики Хакасия, на 2018 год»",
                docCard.DocName, docNamesAreNotEqual);

            Assert.IsNotNull(docCard.DocLobbies);
            Assert.AreEqual(1, docCard.DocLobbies.Select(s => s.ID).Distinct().Count(), docLobbyCountAreNotEqual, docLobbyCountAreNotEqual);
            Assert.AreEqual("Правительство Республики Хакасия", docCard.DocLobbies[0].Name, docLobbiesAreNotEqual);

            Assert.IsTrue(docCard.DocDate.HasValue, docDateIsNull);
            Assert.AreEqual(new DateTime(2018, 11, 14).Date, docCard.DocDate.Value.Date, docDatesAreNotEqual);

            Assert.IsNotNull(docCard.DocNumbers);
            Assert.AreEqual(1, docCard.DocNumbers.Count, docNumbersCountAreNotEqual);
            Assert.AreEqual("537", docCard.DocNumbers[0], docNumbersAreNotEqual);

            Assert.AreEqual(string.Empty, docCard.DocCaseNumber, docCaseNumberIsNotEmpty);

            Assert.IsNotNull(docCard.DocTypes);
            Assert.AreEqual(1, docCard.DocTypes.Count, docTypesCountAreNotEqual);
            Assert.AreEqual("ПОСТАНОВЛЕНИЕ", docCard.DocTypes[0].Name, docTypesAreNotEqual);

            Assert.IsFalse(docCard.DocVersionDate.HasValue, docVersionDateIsNotNull);

            Assert.IsFalse(docCard.RegDate.HasValue, regDateIsNotNull);

            Assert.AreEqual(docCard.RegNumber, string.Empty, regNumberIsNotEmpty);
        }
    }
}
