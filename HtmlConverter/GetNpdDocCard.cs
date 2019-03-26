using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace HtmlConverter
{
    public class GetNpdDocCard
    {
        #region Паттерны и HTML-константы
        /// <summary>
        /// Паттерн для поиска дат вида "от dd MMMM yyyy"
        /// </summary>
        private const string DatePattern1 = @"(от)?\s*(?'date'\d{1,2}\s+[а-яА-Я]{3,10}\s+\d{4}\s+)";
        /// <summary>
        /// Паттерн для поиска дат вида "от dd MMMM yyyy [[г.][ода]]"
        /// </summary>
        private const string DatePattern1Custom = @"(от)?\s*(?'date'\d{1,2}\s+[а-яА-Я]{3,10}\s+\d{4}\s+)[г\.ода\s]*";
        /// <summary>
        /// Паттерн для поиска дат вида "от dd.MM.yyyy"
        /// </summary>
        private const string DatePattern2 = @"(от)?\s*(?'date'\d{1,2}\.\d{1,2}\.\d{4})";
        /// <summary>
        /// Паттерн для поиска номеров документа.
        /// </summary>
        private const string NumberPattern = @"(?:N|№)\s+((?:[^,\s]+(?: [А-Яа-я][,\s])?(?:,\s*)?))";
        /// <summary>
        /// Паттерн фамилии и инициалов (для поиска подписи)
        /// </summary>
        private const string SignaturePattern = @"([А-ЯЁ\-]{0,3}\.)\s?([А-ЯЁ\-]{0,3}\.)\s?([А-ЯЁа-яё-]+)";
        /// <summary>
        /// Паттерн фамилии и инициалов (с одним инициалом)
        /// </summary>
        private const string SignatureSmallPattern = @"([А-ЯЁ]\.)\s?([А-ЯЁа-яё-]{2,})";
        /// <summary>
        /// Шаблон для форматирования дат к нашему стандарту
        /// </summary>
        private const string DateTemplate = "d MMMM yyyy года";
        /// <summary>
        /// Переход на новую строку без создания параграфа (shift+enter)
        /// </summary>
        private readonly XElement _br = new XElement("br");
        /// <summary>
        /// Неразрывный пробел 
        /// </summary>
        private readonly char _nbsp = (char)160;
        #endregion

        #region Свойства карточки
        public string DocCaseNumber { get; set; }
        /// <summary>
        /// Дата документа
        /// </summary>
        public DateTime? DocDate { get; set; }
        /// <summary>
        /// Принявший(-ие) орган(ы)
        /// </summary>
        public List<DocLobbyFormat> DocLobbies { get; set; }
        /// <summary>
        /// Название (заголовок)
        /// </summary>
        public string DocName { get; set; }
        /// <summary>
        /// Номер(а)
        /// </summary>
        public List<string> DocNumbers { get; set; }
        /// <summary>
        /// Тип(ы)
        /// </summary>
        public List<DocType> DocTypes { get; set; }
        /// <summary>
        /// Редакция от
        /// </summary>
        public DateTime? DocVersionDate { get; set; }
        /// <summary>
        /// Дата регистрации документа в МинЮсте
        /// </summary>
        public DateTime? RegDate { get; set; }
        /// <summary>
        /// Регистрационный №
        /// </summary>
        public string RegNumber { get; set; }
        public List<RegionRf> DocRegions { get; set; }
        #endregion

        public string DocTypeString
        {
            get
            {
                return string.Join("; ", DocTypes.Select(s => s.ToString()));
            }
        }
        public string DocLobbyString
        {
            get
            {
                return string.Join("; ", DocLobbies.Select(s => s.ToString()));
            }
        }
        public string DocNumberString
        {
            get
            {
                return string.Join("; ", DocNumbers);
            }
        }
        public string DocRegionString
        {
            get
            {
                return string.Join("; ", DocRegions);
            }
        }

        public bool HasAutoCard { get; set; }

        public GetNpdDocCard()
        {
            DocName = string.Empty;
            RegNumber = string.Empty;
            DocCaseNumber = string.Empty;
            DocLobbies = new List<DocLobbyFormat>();
            DocNumbers = new List<string>();
            DocTypes = new List<DocType>();
            DocRegions = new List<RegionRf>();
            HasAutoCard = false;
        }

        /// <summary>
        /// Форматирование шапки и сбор метаданных
        /// </summary>
        public void FormatHeader(XElement documentBody)
        {
            if (documentBody.Elements().Count(p => p.Name.LocalName == "p" && p.Value.Trim().Length > 0) < 5)
            {
                return;
            }

            DeleteWarmCardTable(documentBody);

            // Получаем всех потомков - т.к. параграфы могут быть вложены в таблицу
            var wordParagraphs = documentBody
                .Descendants()
                .Where(p => p.Name.LocalName == "p" || (p.Name.LocalName.StartsWith("h") && p.Name.LocalName.Length == 2))
                .ToList();

            // 0 - определяем наличие атрибутов регистрации в минюсте
            var registrationParagraphs = wordParagraphs.Take(5).ToList();
            var registrationParagraph = registrationParagraphs
                .FirstOrDefault(p => p.Value.Trim().StartsWith("зарегистрировано ", StringComparison.CurrentCultureIgnoreCase));

            if (registrationParagraph != null)
            {
                string parRegText = registrationParagraph.Value;

                // 2.1 - дата
                DateTime regDate;
                if (TryMatchDate(parRegText, DatePattern1, out regDate))
                {
                    RegDate = regDate;
                }
                if (RegDate == null && TryMatchDate(parRegText, DatePattern2, out regDate))
                {
                    RegDate = regDate;
                }

                // 2.2 - номер            
                RegNumber = Regex.Match(parRegText, NumberPattern).Groups[1].Value;

                if (RegDate != null)
                {
                    // Удаляем
                    wordParagraphs.Remove(registrationParagraph);
                    registrationParagraph.Remove();
                    documentBody.Elements().TakeWhile(p => p.Value.Trim() == "").Remove();
                }
            }

            #region Выраниваем параграфы по центру
            wordParagraphs
                .Take(10)
                .Where(p => p.Value.StartsWith("от") && Regex.IsMatch(p.Value, NumberPattern))
                .SetStyle("text-align", "center");

            // Только первые 2 параграфа. 4 - очень много. Максимум - можно 4 попробовать.
            wordParagraphs.Take(2).SetStyle("text-align", "center");
            #endregion

            // 1 - узнаем границы шапки.
            // проверяем на наличие таблицы с датой
            var firstTable = documentBody.Descendants().FirstOrDefault(X.Table);

            var temp = wordParagraphs.Select(s => new { Style = s.GetStyle("text-align"), s.Value }).ToList();

            // Получаем параграфы вводной части
            // Берём до абзаца "Принят" - это подпись к закону
            var headerParagraphs = wordParagraphs
                .TakeWhile(p => IsHeaderParagraph(p, firstTable))
                .TakeWhile(p => p.Value.Trim().ToLower() != "(извлечение)")
                .TakeWhile(s => !s.Value.Trim().StartsWith("принят", StringComparison.CurrentCultureIgnoreCase) &&
                    !s.Value.Trim().StartsWith("одобрен", StringComparison.CurrentCultureIgnoreCase))
                .ToList();

            if (headerParagraphs.Count < 2)
            {
                return;
            }

            #region Если есть абзацы, которые входят в таблицу, то преобразыем табл в текст
            if (firstTable != null)
            {
                var tableParagraphs = firstTable.Descendants().Where(X.Paragraph);
                if (tableParagraphs.Any(p => headerParagraphs.Contains(p)))
                {
                    firstTable.Parent.Remove();
                }
            }
            #endregion

            // Удаляем верхнюю черту, если такая есть
            if (documentBody.Elements().Any(p => p.Name.LocalName == "div" && p.Value.Trim() == ""))
            {
                documentBody.Elements().First(p => p.Name.LocalName == "div" && p.Value.Trim() == "").Remove();
            }

            // Зачем-то делаем копию списка абзацев
            var headerParagraphsCopy = headerParagraphs.ToList();
            // текст абзацев
            var headerParagraphsTexts = headerParagraphsCopy.Select(s => s.Value.Trim().TrimNewLine()).ToList();

            var fullStringHeader = new StringBuilder();
            headerParagraphsTexts.ForEach(headerParagraphText => fullStringHeader.Append(headerParagraphText.Trim() + " "));

            #region 2 - определяем атрибуты
            // номер, дата, тип и орган

            // 2.1 - дата
            DateTime date;
            if (TryMatchDate(fullStringHeader.ToString(), DatePattern1, out date))
            {
                DocDate = date;
            }
            if (DocDate == null && TryMatchDate(fullStringHeader.ToString(), DatePattern2, out date))
            {
                DocDate = date;
            }

            // 2.2 - номер
            foreach (var line in headerParagraphsTexts)
            {
                var docNumberMatches = Regex.Matches(line, NumberPattern);
                foreach (Match docNumberMatch in docNumberMatches)
                {
                    var numbers = docNumberMatch.Groups[1].Value.Split(',')
                        .Select(s => s.Trim())
                        .Distinct()
                        .Where(p => !DocNumbers.Contains(p));
                    DocNumbers.AddRange(numbers);
                }
            }

            // Немного изменил оригинальную логику. Сначала находим все номера
            string docCaseNumber = Regex.Match(fullStringHeader.ToString(), "Дело\\s*" + NumberPattern).Groups[1].Value;
            if (DocNumbers.Contains(docCaseNumber))
            {
                // А затем записываем номер дела, если он он был найден и удаляем его из массива дел
                this.DocCaseNumber = docCaseNumber;
                DocNumbers.Remove(docCaseNumber);

                // Иначе нет гарантии, что все номера документа будут найдены
            }

            // А затем отрезаем мусор, который попал в номера документа, по количеству органов
            // 2.3 - орган
            DocLobbies = WordImportHandler.SearchLobbies(headerParagraphsTexts.Where(p => p.Trim() != ""));
            if (DocLobbies.Count < DocNumbers.Count && DocLobbies.Any())
            {
                DocNumbers.RemoveRange(DocLobbies.Count, DocNumbers.Count - DocLobbies.Count);
            }

            // 2.4 - тип
            var toFindType = headerParagraphsTexts
                .Where(p => p.Trim() != "")
                .Select(p => p.Replace("АПЕЛЛЯЦИОННОЕ", "").Replace("КАССАЦИОННОЕ", "").Trim())
                .Take(3 + (DocLobbies.Count == 0 ? 1 : DocLobbies.Count) * 2)
                .ToList();
            var docTypeWithMatchingText = WordImportHandler.GetDictionaryTypeByText(toFindType);
            // Если тип не удалось определить сходу, пробуем проанализировать строку над номер документа
            if (docTypeWithMatchingText.Count == 0)
            {
                string docTypeString = null;
                for (int i = 0; i < toFindType.Count; i++)
                {
                    if (Regex.IsMatch(toFindType[i], NumberPattern) && i > 0)
                    {
                        docTypeString = toFindType[i - 1];
                    }
                }

                // Для законов нужно пропарсить регион
                if (docTypeString != null && docTypeString.StartsWith("закон", StringComparison.CurrentCultureIgnoreCase))
                {
                    docTypeWithMatchingText.Add(WordImportHandler.GetType("закон"), docTypeString);
                    string regionName = docTypeString.Substring(docTypeString.IndexOf(' ') + 1);
                    var docTypeRegion = WordImportHandler.GetRegion(regionName);
                    if (docTypeRegion != null)
                    {
                        DocRegions.Add(docTypeRegion);
                    }
                }
            }
            DocTypes = docTypeWithMatchingText.Select(s => s.Key).ToList();

            // Для законов пробуем подхватить регион
            if (DocTypes.Any(p => p.Name.ToLower().Contains("закон")))
            {
                if (DocRegions.Count == 0)
                {
                    DocRegions = WordImportHandler.SearchRegions(toFindType);
                }

                // У законов только один номер - дополнительно очищаем, если попал мусор
                if (DocNumbers.Count > 1)
                {
                    DocNumbers.RemoveRange(1, DocNumbers.Count - 1);
                }
            }

            // 2.5 - название
            // Получаем позицию, с которой будем склеивать заголовок
            // Подразумеваем, что это позиция, на которой находится номер документа
            // Т.к. сразу после него идёт название

            var docNumberCount = DocNumbers.Count;
            var nameParaStartIndex = 0;
            foreach (var headerPara in headerParagraphsCopy)
            {
                if (docNumberCount == 0)
                {
                    nameParaStartIndex = headerParagraphsCopy.IndexOf(headerPara);
                    break;
                }
                var isMatch = Regex.IsMatch(headerPara.Value, NumberPattern);
                if (isMatch && docNumberCount > 0)
                {
                    docNumberCount--;

                    // Если на последнем элементе цикла
                    if (docNumberCount == 0)
                    {
                        nameParaStartIndex = headerParagraphsCopy.IndexOf(headerPara);
                    }
                }
            }

            // Получаем позицию, в которой находится тип документа
            int firstIndexOfDocTypeMatchingText = 0;
            foreach (var keyPair in docTypeWithMatchingText)
            {
                for (int j = 0; j < headerParagraphsTexts.Count; j++)
                {
                    if (headerParagraphsTexts[j].Trim() != "")
                    {
                        if (keyPair.Value ==
                            headerParagraphsTexts[j].Replace("АПЕЛЛЯЦИОННОЕ", "").Replace("КАССАЦИОННОЕ", "").Trim())
                        {
                            firstIndexOfDocTypeMatchingText = j;
                        }
                    }
                }
            }

            // Обычно номер документа идёт ПОСЛЕ типа, но иногда это не так
            if (nameParaStartIndex < firstIndexOfDocTypeMatchingText)
            {
                nameParaStartIndex = firstIndexOfDocTypeMatchingText + 1;
            }
            // А иногда под типом есть ещё и регион (и пустая строка)
            var region = DocRegions.FirstOrDefault();
            if (region != null)
            {
                int offset = 0;
                string text = headerParagraphsTexts[nameParaStartIndex];
                if (text.Trim() == "")
                {
                    offset = 1;
                    text = headerParagraphsTexts[nameParaStartIndex + 1];
                }

                if (text.Equals(region.RegionName, StringComparison.CurrentCultureIgnoreCase) || text.Equals(region.RegionNameForDoc, StringComparison.CurrentCultureIgnoreCase))
                {
                    nameParaStartIndex += 1 + offset;
                }
            }

            nameParaStartIndex--;

            if (headerParagraphsCopy.Count < 2)
            {
                return;
            }

            // Дата документа обычно идёт вместе с номером и поэтому отсеивается
            // Но это не всегда так
            if (Regex.IsMatch(headerParagraphsCopy[nameParaStartIndex + 1].Value, DatePattern1,
                    RegexOptions.IgnoreCase | RegexOptions.Compiled) ||
                Regex.IsMatch(headerParagraphsCopy[nameParaStartIndex + 1].Value, DatePattern2,
                    RegexOptions.IgnoreCase | RegexOptions.Compiled))
            {
                nameParaStartIndex++;
            }
            #endregion

            var paragraphNames = headerParagraphsCopy
                .Skip(nameParaStartIndex)
                .ToList();

            // исключаем абзац с "Дело"
            if (paragraphNames.Any(s => s.Value.Trim().StartsWith("Дело", StringComparison.CurrentCultureIgnoreCase)))
            {
                paragraphNames = paragraphNames
                    .SkipWhile(s => !s.Value.Trim().StartsWith("Дело", StringComparison.CurrentCultureIgnoreCase))
                    .ToList();
            }
            // берем до абзаца с текстом "(в ред."
            if (paragraphNames.Any(s => s.Value.Trim().StartsWith("(в ред", StringComparison.CurrentCultureIgnoreCase)))
            {
                paragraphNames = paragraphNames
                    .TakeWhile(s => !s.Value.Trim().StartsWith("(в ред", StringComparison.CurrentCultureIgnoreCase))
                    .ToList();
            }

            // если сведения об изменяющих находятся на одной строке с названием
            foreach (var paragraph in paragraphNames.ToList())
            {
                string value = paragraph.Value.Trim();
                var index = value.IndexOf("(в ред.");
                if (index > 0)
                {
                    paragraph.Value = value.Remove(index);

                    var redaction = new XElement(paragraph);
                    redaction.Value = value.Substring(index);
                    paragraph.AddAfterSelf(redaction);
                    headerParagraphsCopy.Insert(headerParagraphsCopy.IndexOf(paragraph), redaction);
                }
            }

            string[] docNameParts = null;

            if (paragraphNames.Count > 0)
            {
                int offset = 1;
                //if (Regex.IsMatch(paragraphNames[0].Value, DatePattern1, RegexOptions.IgnoreCase | RegexOptions.Compiled) ||
                //    Regex.IsMatch(paragraphNames[0].Value, DatePattern2, RegexOptions.IgnoreCase | RegexOptions.Compiled))
                //{
                //    // Надо пересмотреть логику. Названия часто содержат дату. 
                //    // Зачастую название состоит из одного абзаца.
                //    // И часто оно не очень длинное (до 100 символов), так что проверка на количество плохо канает
                //    if (paragraphNames.Count > 1 && paragraphNames.Skip(1).All(p => p.Value.Trim() != ""))
                //    {
                //        offset = 1;
                //    }

                //    // Если после абзаца в котором был матч идёт пустота - это точно абзац с датой-номером, пропускаем его и следущий
                //    if (paragraphNames.Count > 1 && paragraphNames[1].Value.Trim() == "")
                //    {
                //        offset = 2;
                //    }
                //}

                // Если больше 1 абзаца - пропускаем первый, т.к. это номер с датой
                docNameParts = paragraphNames
                    .Skip(offset)
                    .SelectMany(ExtractText)
                    .Where(p => p.Length > 0)
                    .ToArray();

                if (docNameParts.Any(p => p.Length > 1))
                {
                    // Если все буквы прописные, то заменяем на строчные + первая заглавная.
                    if (docNameParts.All(p => p.ToUpper() == p))
                    {
                        docNameParts[0] = char.ToUpper(docNameParts[0][0]) + docNameParts[0].Substring(1).ToLower();
                        for (int i = 1; i < docNameParts.Length; i++)
                        {
                            docNameParts[i] = docNameParts[i].ToLower();
                        }
                    }

                    DocName = string.Join(" ", docNameParts).Trim();
                }
            }

            // 3 - блок "В редакции"
            var redactionParagraphs = headerParagraphsCopy
                .SkipWhile(p => !p.Value.Trim().StartsWith("(в ред", StringComparison.CurrentCultureIgnoreCase))
                .ToList();

            // Если блок "в редакции" идёт после сведений об изменяющих, то он мог не попасть в headerParagraphs
            // TODO Подумать над тем, чтобы включить его в headerParagraphs
            if (redactionParagraphs.Count == 0)
            {
                redactionParagraphs = wordParagraphs
                    .SkipWhile(p => !p.Value.Trim().StartsWith("(в ред", StringComparison.CurrentCultureIgnoreCase))
                    .TakeWhile(p => p.Value.Trim() != "")
                    .ToList();
            }

            redactionParagraphs = redactionParagraphs
                .TakeWhile(p => redactionParagraphs.IndexOf(p) <= 0 || redactionParagraphs[redactionParagraphs.IndexOf(p) - 1].Value.Trim().EndsWith(")"))
                .ToList();

            // Получаем дату изменений (версия документа)
            // А так же формируем сведения об изменяющих редакциях
            var redactions = GetRedactions(redactionParagraphs);

            XElement header = GenerateHeader(docNameParts);
            // Сведения об изменяющих надо вставлять после сведений о принятии
            headerParagraphs.Last(p => p.Parent != null).AddAfterSelf(redactions);

            // Удаляем старую шапку
            if (!header.IsEmpty)
            {
                headerParagraphs.Where(p => p.Parent != null).Remove();
            }
            // Вставляем новую шапку
            foreach (var elem in header.Elements().Reverse())
            {
                documentBody.AddFirst(elem);
            }

            HasAutoCard = true;
            ConsoleHelpers.PrintCard(this);
        }

        private static void DeleteWarmCardTable(XElement documentBody)
        {
            var tableElement = documentBody.Descendants().FirstOrDefault(p => p.Name.LocalName == "table");
            if (tableElement != null)
            {
                var firstCell = tableElement.Descendants().FirstOrDefault(p => p.Name.LocalName == "td");
                if (firstCell != null && firstCell.Value == "Модуль")
                {
                    tableElement.Parent.Remove();
                }
            }
        }

        private XElement GenerateHeader(IEnumerable<string> docNameParts)
        {
            if (DocTypes.Count == 0)
            {
                return new XElement("div");
            }

            var header = new XElement("div");
            if (DocLobbies.Count > 1)
            {
                int i = -1;
                while (++i < DocLobbies.Count)
                {
                    header.Add(new XElement("p",
                        new XAttribute("id", "doclobby"),
                        new XAttribute("class", "header"),
                        new XAttribute("style", "text-align: center"),
                        new XElement("strong", DocLobbies[i].NameHeader)));

                    if (DocDate.HasValue && DocNumbers.Any())
                    {
                        header.Add(new XElement("p",
                            new XAttribute("id", "docdate"),
                            new XAttribute("class", "header"),
                            new XAttribute("style", "text-align: center"),
                            new XElement("strong", string.Format("от {0} № {1}",
                                DocDate.Value.ToString(DateTemplate),
                                DocNumbers.Count > i ? DocNumbers[i] : DocNumbers.Last()))));
                    }
                }

                if (DocTypes.Any())
                {
                    header.Add(new XElement("p",
                        new XAttribute("id", "doctype"),
                        new XAttribute("class", "header"),
                        new XAttribute("style", "text-align: center"),
                        new XElement("strong", DocTypes.First().Name, DocRegions.Any() ? " " + DocRegions.First().RegionNameR : "")));
                }
            }
            else
            {
                foreach (var lobby in DocLobbies)
                {
                    header.Add(new XElement("p",
                            new XAttribute("id", "doclobby"),
                            new XAttribute("class", "header"),
                            new XAttribute("style", "text-align: center"),
                            new XElement("strong", lobby.NameHeader.ToUpper())));
                }

                foreach (var type in DocTypes)
                {
                    header.Add(new XElement("p",
                        new XAttribute("id", "doctype"),
                        new XAttribute("class", "header"),
                        new XAttribute("style", "text-align: center"),
                        new XElement("strong", type.Name, DocRegions.Any() ? " " + DocRegions.First().RegionNameR : "")));
                }

                if (DocDate.HasValue && DocNumbers.Any())
                {
                    header.Add(new XElement("p",
                        new XAttribute("id", "docnumber"),
                        new XAttribute("class", "header"),
                        new XAttribute("style", "text-align: center"),
                        new XElement("strong", string.Format("от {0} № {1}", DocDate.Value.ToString(DateTemplate), DocNumbers.First()))));
                }
            }
            if (DocNumbers.Any() && DocCaseNumber != "")
            {
                header.Add(new XElement("p",
                        new XAttribute("id", "docnumber"),
                        new XAttribute("class", "header"),
                        new XAttribute("style", "text-align: center"),
                        new XElement("strong", string.Format("Дело № {0}", DocCaseNumber))));
            }
            if (DocName != "")
            {
                header.Add(new XElement("p",
                        new XAttribute("id", "docname"),
                        new XAttribute("class", "header"),
                        new XAttribute("style", "text-align: center"),
                        docNameParts.Select(s => new XElement("strong", s, _br))));
            }

            return header;
        }

        /// <summary>
        /// Возвращает параграф с блоком "Сведения об  изменяющих"
        /// </summary>
        /// <param name="redactionParagraphs">Перечисление элементов, которые содержат сведения об изменяющих</param>
        /// <returns></returns>
        private XElement GetRedactions(IEnumerable<XElement> redactionParagraphs)
        {
            if (redactionParagraphs == null || !redactionParagraphs.Any())
            {
                return null;
            }

            const string line69 = "_____________________________________________________________________";

            var result = new XElement("p",
                new XAttribute("class", "redactions"),
                new XElement("span", line69, _br),
                new XElement("span", "Документ с изменениями, внесенными:", _br));

            var paragraphTexts = redactionParagraphs.Select(p => p.Value.Trim()).Where(p => p != "").ToArray();
            // Удаляем "(в ред."
            paragraphTexts[0] = Regex.Replace(paragraphTexts[0], "\\(в ред[^ ]+ ?", "");
            if (paragraphTexts.Last().EndsWith(")"))
            {
                // Удаляем закрывающую скобку
                paragraphTexts[paragraphTexts.Length - 1] = paragraphTexts[paragraphTexts.Length - 1].Remove(paragraphTexts[paragraphTexts.Length - 1].Length - 1);
            }

            try
            {
                string date = Regex.Matches(string.Join(", ", paragraphTexts), @"от\s+(\d\d\.\d\d\.\d\d\d\d)\s+")
                    .Cast<Match>()
                    .Last().Groups[1].Value;

                DocVersionDate = DateTime.Parse(date);
            }
            catch { }

            string patternDateNum = @"от\s+\d\d\.\d\d\.\d\d\d\d\s+(N|№)\s+[^,]+";
            string lobbyName = "";

            foreach (string text in paragraphTexts)
            {
                foreach (string redaction in text.Split(',').Select(s => s.Trim()).Where(p => !string.IsNullOrEmpty(p)))
                {
                    if (!string.IsNullOrEmpty(Regex.Replace(redaction, patternDateNum, "").Trim()))
                    {
                        lobbyName = Regex.Replace(redaction, patternDateNum, "").Trim();

                        var docType = DocType.Common.FirstOrDefault(p =>
                            lobbyName.StartsWith(p.Name, StringComparison.CurrentCultureIgnoreCase) ||
                            lobbyName.StartsWith(p.NamePluralForRedactions, StringComparison.CurrentCultureIgnoreCase));

                        // Поправляем название органа
                        if (docType != null)
                        {
                            if (lobbyName.StartsWith(docType.NamePluralForRedactions, StringComparison.CurrentCultureIgnoreCase))
                            {
                                lobbyName = lobbyName.Substring(docType.NamePluralForRedactions.Length);
                            }
                            else if (lobbyName.StartsWith(docType.NameR, StringComparison.CurrentCultureIgnoreCase))
                            {
                                lobbyName = lobbyName.Substring(docType.NameR.Length);
                            }
                            else
                            {
                                lobbyName = lobbyName.Substring(docType.Name.Length);
                            }

                            lobbyName = docType.NameForRedactions + " " + lobbyName.Trim();
                        }
                    }

                    Regex.Matches(redaction, patternDateNum)
                        .Cast<Match>()
                        .Select(s => Regex.Replace(s.Value, "от\\s+(\\d\\d)\\.(\\d\\d)\\.(\\d\\d\\d\\d)\\s+(?:г.|года\\s+)?",
                        m =>
                        "от " + new DateTime(
                            Convert.ToInt32(m.Groups[3].Value),
                            Convert.ToInt32(m.Groups[2].Value),
                            Convert.ToInt32(m.Groups[1].Value)
                            ).ToString(DateTemplate) + " "))
                        .ToList()
                        .ForEach(x => result.Add(new XElement("span", lobbyName + " " + x.Replace(" N ", " № "), _br)));
                }
            }
            
            result.Add(new XElement("span", line69));
            redactionParagraphs.Remove();

            return result;
        }

        private bool IsHeaderParagraph(XElement paragprah, XElement table)
        {
            var textAlign = paragprah.GetStyle("text-align");

            // Вроде бы вырвавние по умолчанию - по левому краю
            if (textAlign == null)
            {
                // TODO: Срочно пересмотреть эту логику
                return paragprah.Value.Trim() == "" || IsWithinTable(paragprah, table);
            }

            // Выравнивание по левому краю - точно не у заголовков
            if (textAlign == "left")
            {
                return IsWithinTable(paragprah, table);
                //return false;
            }

            // Иногда заголовки бывают по ширине
            if (textAlign.Contains("justify"))
            {
                // Но если он пустой, то это не заголовок
                return paragprah.Value.Trim() == "";
            }

            return true;
        }

        private bool IsWithinTable(XElement paragprah, XElement table)
        {
            return table != null && table.Descendants().Any(p => p == paragprah);
        }

        /// <summary>
        /// Форматирует блок "Сведения об изменяюших"
        /// </summary>
        public void FormatAcceptance(XElement documentBody)
        {
            // Слегка изменённый паттерн стандартной даты ("дд ММММ гггг [[г.][ода]]")
            const string dateTemplate = @"(?'date'\d{1,2}\s+[а-яА-Я]{3,10}\s+\d{4}\s+)[г\.ода\s]*";

            var paragraphs = documentBody.Elements()
                .Where(X.Paragraph)
                .SkipWhile(p => p.HasAttribute("class"))
                .ToList();

            var result = new XElement("div");
            List<XElement> acceptance;
            do
            {
                acceptance = new List<XElement>();
                // Пропускаем, пока не содержит слова "принят", "одобрен"
                foreach (var p in paragraphs
                    .SkipWhile(p => !p.Value.Trim().StartsWith("принят", StringComparison.CurrentCultureIgnoreCase) &&
                                    !p.Value.Trim().StartsWith("одобрен", StringComparison.CurrentCultureIgnoreCase))
                    .Take(8))
                {
                    // Берём пока не будет пустая строка,
                    // Либо пока не будет выравнивание по другому краю
                    if (p.Value.Trim() != "")
                    {
                        acceptance.Add(p);
                        if (!NextElementHasSameStyle(p))
                        {
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }

                if (acceptance.Any())
                {
                    var firstWord = acceptance.First().Value.Split(' ')[0];
                    if (!firstWord.Equals("принят", StringComparison.CurrentCultureIgnoreCase) &&
                        !firstWord.Equals("одобрен", StringComparison.CurrentCultureIgnoreCase))
                    {
                        // break - потому что абзацы принятия идут друг за другом,
                        // как только один неверный - они кончились
                        break;
                    }

                    var acceptanceElem = new XElement("p", new XAttribute("class", "acceptance"));
                    acceptanceElem.SetStyle("text-align", "right");
                    foreach (var elem in acceptance.SelectMany(ExtractText))
                    {
                        string value = elem;
                        // Парсим дату и приводим к стандартному виду
                        DateTime date;
                        if (TryMatchDate(value, dateTemplate, out date))
                        {
                            value = Regex.Replace(value, dateTemplate, date.ToString(DateTemplate + " "));
                        }
                        else if (TryMatchDate(value, DatePattern2, out date))
                        {
                            value = Regex.Replace(value, DatePattern2, date.ToString(DateTemplate));
                        }

                        acceptanceElem.Add(new XElement("span", value.Trim(), _br));
                    }
                    result.Add(acceptanceElem);

                    int skipCount = paragraphs.IndexOf(acceptance.First());
                    int takeCount = paragraphs.IndexOf(acceptance.Last()) + 1;
                    if (paragraphs[takeCount].Value.Trim() == "")
                    {
                        takeCount++;
                    }

                    takeCount -= skipCount;

                    // Удаляем из документа
                    paragraphs.Skip(skipCount).Take(takeCount).Remove();
                    // Удаляем из массива
                    paragraphs.RemoveRange(skipCount, takeCount);

                    // Идём дальше
                    acceptance = paragraphs
                        .TakeWhile(p => p.Value.Trim() != "").ToList();
                }
            } while (acceptance.Any());

            if (result.Elements().Any())
            {
                var first = documentBody.Elements()
                    .Where(X.Paragraph)
                    .LastOrDefault(p => p.HasClass("header"));

                if (first == null)
                {
                    first = documentBody.Elements()
                        .Where(X.Paragraph)
                        .First();
                }

                first.AddAfterSelf(result.Elements().ToArray());
            }
        }

        internal XElement GetPrevious(List<XElement> elems, int index, bool skipEmpty)
        {
            if (index == 0)
            {
                return null;
            }

            if (!skipEmpty)
            {
                return elems[index - 1];
            }

            while (index > 0)
            {
                index--;
                if (elems[index] != null && elems[index].Value.Trim() != "")
                {
                    break;
                }
            }

            return elems[index];
        }

        /// <summary>
        /// Извлекает из указанной последовательности элементы, которые составляют подпись документа.
        /// </summary>
        /// <param name="wordParagraphs">Последовательность элементов параграфов документа (с конца).</param>
        /// <returns></returns>
        private List<XElement> GetSignatureElements(IEnumerable<XElement> wordParagraphs)
        {
            var result = new List<XElement>();
            foreach (var paragraph in wordParagraphs)
            {
                if (paragraph.Value.Trim().Length > 0)
                {
                    var textAlign = paragraph.GetStyle("text-align") ?? "left";
                    if (textAlign != "right" && textAlign != "left")
                    {
                        break;
                    }

                    if (result.Count > 0 && textAlign == "left")
                    {
                        break;
                    }

                    if (textAlign == "right")
                    {
                        if (result.Any() && paragraph.Descendants().Count() == 1)
                        {
                            result.First().AddFirst(paragraph.Elements(), _br);
                            // Т.к. мы берём внутренности элемента, дальше в цикле мы его уже поймаем - удаляем сразу
                            paragraph.Remove();
                        }
                        else
                        {
                            result.Insert(0, paragraph);
                        }
                    }
                }
                else
                {
                    if (result.Any() && result.First().Value.Trim().Length > 0)
                    {
                        result.Insert(0, paragraph);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Оформление подписи. Возвращает True если подпись найдена и сформирована
        /// </summary>
        public bool FormatSignature(XElement documentBody)
        {
            var wordParagraphs = documentBody
                .Elements()
                .Where(X.Paragraph)
                .ToList();

            var debug = wordParagraphs.Select(s => s.Value).ToArray();

            var grifIndex = 0;
            var signParagraphs = new List<XElement>();
            // Идём по списку параграфов
            for (int i = 0; i < wordParagraphs.Count; i++)
            {
                // Если встречаем гриф, то берём все предыдущие параграфы,
                if (wordParagraphs[i].HasClass("grif"))
                {
                    grifIndex = i + 1;
                    // Берём все элементы сверху списка, которые по правому краю
                    // Берём, пока не встретится элемент не поправому краю и найдём хотя бы один элемент

                    // Пробуем поймать подпись, которая выровнена по левому краю
                    var previousPararagph = GetPrevious(wordParagraphs, i, true);
                    if (Regex.IsMatch(previousPararagph.Value, SignaturePattern))
                    {
                        previousPararagph.SetStyle("text-align", "right");
                    }

                    var temp = GetSignatureElements(wordParagraphs.Take(i).Reverse());
                    signParagraphs.AddRange(temp);

                    // У которых идёт подряд text-align=right
                }
                else if (i == wordParagraphs.Count - 1)
                {
                    i++;
                    var temp = GetSignatureElements(wordParagraphs.Skip(grifIndex).Take(i).Reverse());
                    signParagraphs.AddRange(temp);
                }
            }

            // Сначала проверяем по упрощенному паттерну
            if (!signParagraphs.Any() || !signParagraphs.Any(p => Regex.IsMatch(p.Value, SignatureSmallPattern)))
            {
                return false;
            }

            // Иногда приезжает пустой параграф
            if (signParagraphs.First().Value.Trim().Length == 0)
            {
                signParagraphs.Remove(signParagraphs.First());
            }
            
            foreach (var signParagraph in signParagraphs.Distinct().ToList())
            {
                var signature = new XElement("p", new XAttribute("class", "signature"));
                signature.SetStyle("text-align", "right");

                var etext = ExtractText(signParagraph);
                foreach (var line in etext.Where(p => p.Trim().Length > 0))
                {
                    string text = FormatNamesInSignature(line);
                    signature.Add(new XText(text), _br);
                }
                signParagraph.AddAfterSelf(signature);

                if (signParagraph.Parent != null)
                {
                    signParagraph.Remove();
                }
            }

            return true;
        }

        public string FormatNamesInSignature(string text)
        {
            text = text.Trim();
            
            // Здесь уже сперва проверяем на нормальные инициалы
            if (Regex.IsMatch(text, SignaturePattern))
            {
                text = Regex.Replace(text, SignaturePattern, m =>
                {
                    string nameAndMiddleName = m.Groups[1].Value + m.Groups[2].Value;
                    string surname = m.Groups[3].Value;
                    surname = char.ToUpper(surname[0]) + surname.Substring(1).ToLower();

                    // Соединяем неразрывным пробелом
                    return nameAndMiddleName + _nbsp + surname;
                });
            }
            // Затем на упрощенные
            else if (Regex.IsMatch(text, SignatureSmallPattern))
            {
                text = Regex.Replace(text, SignatureSmallPattern, m =>
                {
                    string name = m.Groups[1].Value;
                    string surname = m.Groups[2].Value;
                    surname = char.ToUpper(surname[0]) + surname.Substring(1).ToLower();

                    // Соединяем неразрывным пробелом
                    return name + _nbsp + surname;
                });
            }

            return text;
        }

        /// <summary>
        /// Формирует блок Регистрация в минюсте
        /// </summary>
        /// <param name="signatureExists">Признак того, что блок формируем в конце. Если не задан, то блок формируем в начале.</param>
        public void CreateBlockRegistrationMinust(XElement documentBody, bool signatureExists)
        {
            if (string.IsNullOrEmpty(RegNumber) || !RegDate.HasValue)
            {
                return;
            }

            XElement insertAfter = null;
            if (signatureExists)
            {
                insertAfter = documentBody.Descendants("p").Where(p => p.HasClass("signature")).LastOrDefault();
            }

            if (insertAfter == null)
            {
                insertAfter = documentBody.Descendants().LastOrDefault(X.Paragraph);
            }

            var regParagraph = new XElement("p",
                    new XElement("span", "Зарегистрировано", _br),
                    new XElement("span", "в Министерстве юстиции", _br),
                    new XElement("span", "Российской Федерации", _br),
                    new XElement("span", RegDate.Value.ToString(DateTemplate), _br),
                    new XElement("span", "регистрационный № " + RegNumber, _br)
                    );
            regParagraph.ReplaceAttributes(insertAfter.Attributes().ToArray());
            regParagraph.SetStyle("text-align", "left");

            insertAfter.AddAfterSelf(regParagraph);
        }

        public void FormatGrif(XElement documentBody)
        {
            var paragraphs = documentBody.Elements()
                .Where(X.Paragraph)
                .SkipWhile(p => p.HasAttribute("class"))
                .ToList();

            var offset = 0;
            var grifParagraphs = new List<XElement>();
            foreach (var elem in paragraphs
               .Skip(offset)
               .SkipWhile(p => !IsGrif(p))
               .ToList())
            {
                // Мы нашли, где начинается гриф
                grifParagraphs.Add(elem);
                // Берём элементы, пока следущий не будет отличаться по стилю
                if (!NextElementHasSameStyle(elem) || elem.Elements().Count() > 1)
                {
                    // Второе условие - для случаев, когда текст вообще не форматирован
                    // Пробуем отличить, по признаку, что переход на новый абзац - это другой элемент документа
                    break;
                }
            }
            
            while (grifParagraphs.Any())
            {
                var grifElem = new XElement("p",
                    new XAttribute("class", "grif"));
                grifElem.SetStyle("text-align", "right");

                foreach (var elem in grifParagraphs.Where(p => p.Value.Trim() != ""))
                {
                    // Приложение может быть уже автоформатировано через шифт-энтер
                    // Тогда внутри elem будуте спаны и <br />&#x200e;

                    // Атрибут lang приезжает из docx при конвертации (но по идее, может и не приехать, кто его знает)
                    var etext = ExtractText(elem);
                    foreach (var line in ExtractText(elem))
                    {
                        string text = line;
                        DateTime date;
                        if (TryMatchDate(text, DatePattern1Custom, out date))
                        {
                            text = Regex.Replace(text, DatePattern1Custom, date.ToString("от " + DateTemplate + " "));
                        }
                        else if (TryMatchDate(text, DatePattern2, out date))
                        {
                            text = Regex.Replace(text, DatePattern2, date.ToString("от " + DateTemplate + " "));
                        }

                        if (Regex.IsMatch(text, NumberPattern, RegexOptions.IgnoreCase | RegexOptions.Compiled))
                        {
                            var number = Regex.Match(text, NumberPattern, RegexOptions.Compiled).Groups[1].Value;
                            text = Regex.Replace(text, NumberPattern, "№ " + number);
                        }

                        grifElem.Add(new XElement("span", text, _br));
                    }
                }

                offset = paragraphs.IndexOf(grifParagraphs.Last()) + 1;

                // Вставляем новый гриф
                grifParagraphs.Last().AddAfterSelf(grifElem);
                // Удаляем старый гриф
                grifParagraphs.Remove();

                grifParagraphs.Clear();
                foreach (var elem in paragraphs
                    .Skip(offset)
                    .SkipWhile(p => !IsGrif(p))
                    .ToList())
                {
                    grifParagraphs.Add(elem);
                    if (!NextElementHasSameStyle(elem) || elem.Elements().Count() > 1)
                    {
                        break;
                    }
                }

            } 
        }

        /// <summary>
        /// Достает текст из параграфа таким образом, что каждая строка в параграфе - это элемент массива.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        internal List<string> ExtractText(XElement paragraph)
        {
            var docNameParts = new List<string> { "" };

            foreach (var span in paragraph.Descendants()
                .Where(p => p.Name.LocalName == "span"))
            {
                if (span.Elements().Any(p => p.Name.LocalName == "br"))
                {
                    // По идее, x200e бывает только рядом с <br />
                    // Новая строка
                    docNameParts.Add("");
                }
                else
                {
                    // TODO: Подумать об использовании атрибута lanf
                    // Дописываем в последний элемент
                    int last = docNameParts.Count - 1;
                    if (docNameParts[last].Length > 0)
                    {
                        string value = docNameParts[last];
                        // Иногда попадаются пустые спаны и получаются двойные пробелы.
                        if (docNameParts[last][value.Length - 1] != ' ' && !IsPunctuation(span.Value))
                        {
                            docNameParts[docNameParts.Count - 1] += " ";
                        }
                    }
                    docNameParts[docNameParts.Count - 1] += span.Value.Trim();
                }

                if (docNameParts.Last() != "")
                {
                    if ((span.NextNode is XElement) && ((XElement)span.NextNode).Name.LocalName == "br")
                    {
                        // Новая строка
                        docNameParts.Add("");
                    }
                }
            }

            if (docNameParts.Last() == "")
            {
                docNameParts.Remove(docNameParts.Last());
            }

            return docNameParts;
        }        

        private bool IsPunctuation(string value)
        {
            if (value.Trim().Length < 1)
            {
                return false;
            }
            var ch = value.Trim()[0];
            return ch == ',' ||
                ch == '.' ||
                ch == ';' ||
                ch == ':' ||
                ch == '!' ||
                ch == '?';
        }

        internal bool IsGrif(XElement paragraph)
        {
            var elem = paragraph;
            bool center = elem.GetStyle("text-align") == "center";
            // На случай, если оформлено без переноса на новую строку
            if (paragraph.Elements().Any(p => p.Name.LocalName == "span"))
            {
                elem = paragraph.Elements().First(p => p.Name.LocalName == "span");
            }

            bool bold = elem.GetStyle("font-weight") == "bold";

            string text = elem.Value.Trim();
            if (text.Length < 4)
            {
                return false;
            }

            var split = text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (split[0] == "Приложение" && !(center || bold))
            {
                return true;
            }

            if (split[0] == "Форма" && !(center || bold))
            {
                return true;
            }
            
            // Договорились, что однокоренные слова не считаем - поэтому в регулярке стоит "конец строки"
            return Regex.IsMatch(split[0].ToLower(), "[у]твержд[её]н[аыо]?$", RegexOptions.Compiled);
        }

        internal bool NextElementHasSameStyle(XElement elem)
        {
            var nextElem = elem.ElementsAfterSelf().FirstOrDefault();
            if (nextElem == null)
            {
                return false;
            }

            var style = nextElem.Attribute("style");
            if (style == null)
            {
                return false;
            }

            if (nextElem.GetStyle("text-align") != elem.GetStyle("text-align")
                || nextElem.GetStyle("font-size") != elem.GetStyle("font-size")
                || nextElem.GetStyle("font-family") != elem.GetStyle("font-family"))
            {
                return false;
            }

            var span = elem.Element("span");
            var nextSpan = nextElem.Element("span");
            if (span != null && nextSpan != null)
            {
                return nextSpan.GetStyle("font-weight") == span.GetStyle("font-weight");
            }

            // Если дошли до сюда и всё ещё совпадает, то походу и правда такой же

            return true;
        }

        /// <summary>
        /// Защита от опечаток. 
        /// Пробует распознать указанный паттерн даты в указанной строке и в случае успеха, пробует сконвертировать его в объект <see cref="DateTime"/>.
        /// </summary>
        /// <param name="text">Текст для поиска даты.</param>
        /// <param name="datePattern">Паттерн даты.</param>
        /// <param name="dateTime">Результат преобразования найденного текста в дату.</param>
        /// <returns></returns>
        internal bool TryMatchDate(string text, string datePattern, out DateTime dateTime)
        {
            if (Regex.IsMatch(text, datePattern, RegexOptions.IgnoreCase | RegexOptions.Compiled))
            {
                if (DateTime.TryParse(Regex.Match(text, datePattern, RegexOptions.IgnoreCase).Groups["date"].Value, out dateTime))
                {
                    return true;
                }
            }

            dateTime = DateTime.MinValue;
            return false;
        }
    }
}
