using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HtmlConverter01
{
    public class WordImportHandler
    {
        /// <summary>
        /// Поиск полного названия органа по тексту из параграфов
        /// </summary>
        /// <param name="paragraphs">Текста параграфов</param>
        /// <returns></returns>
        public static List<DocLobbyFormat> SearchLobbies(IEnumerable<string> paragraphs)
        {
            Console.WriteLine("Поиск органов в базе " + WordImportDal.DataSource + ":");
            int count = 1;
            var result = new List<DocLobbyFormat>();
            foreach (var text in paragraphs.Where(p1 => p1.Length > 3 && !p1.All(p2 => char.IsDigit(p2) || p2 == '.')))
            {
                var dbLobby = WordImportDal.GetListLobbyByText(text);
                result.AddRange(dbLobby);

                LogMatch(ref count, text, dbLobby.LastOrDefault());
            }

            return result;
        }

        /// <summary>
        /// Поиск полного типа документа по тексту из параграфов
        /// </summary>
        /// <param name="paragraphs">Текста параграфов</param>
        /// <returns></returns>
        public static Dictionary<DocType, string> GetDictionaryTypeByText(IEnumerable<string> paragraphs)
        {
            Console.WriteLine("Поиск типов в базе " + WordImportDal.DataSource + ":");
            int count = 1;
            var result = new Dictionary<DocType, string>();
            foreach (var text in paragraphs)
            {
                foreach (var docType in WordImportDal.GetListTypeByText(text))
                {
                    result.Add(docType, text);

                    LogMatch(ref count, text, docType);
                }
            }

            return result;
        }

        /// <summary>
        /// Получает из базы тип документа с заданным названием
        /// </summary>
        /// <returns></returns>
        public static DocType GetType(string text)
        {
            return WordImportDal.GetListTypeByText(text).FirstOrDefault();
        }

        /// <summary>
        /// Получает из базы регион с заданным именем
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        public static RegionRf GetRegion(string query)
        {
            return WordImportDal.GetRegionsByText(query).FirstOrDefault();
        }

        /// <summary>
        /// Поиск в базе регионов по перечисленным текстам
        /// </summary>
        /// <param name="queries"></param>
        /// <returns></returns>
        public static List<RegionRf> SearchRegions(IEnumerable<string> queries)
        {
            Console.WriteLine("Поиск регионов в базе " + WordImportDal.DataSource + ":");
            int count = 1;
            var result = new List<RegionRf>();
            foreach (var text in queries)
            {
                var dbRegion = WordImportDal.GetRegionsByText(text);
                result.AddRange(dbRegion);

                LogMatch(ref count, text, dbRegion.LastOrDefault());
            }

            return result;
        }

        /// <summary>
        /// Логгирует найденное совпадение в базе.
        /// </summary>
        /// <param name="counter"></param>
        /// <param name="source"></param>
        /// <param name="match"></param>
        private static void LogMatch(ref int counter, string source, object match)
        {
            if (match != null)
            {
                Console.WriteLine("{0}. {1}\n\t-> {2}", counter++, source, match.ToString());
            }
            else
            {
                Console.WriteLine("{0}. [x] {1}", counter++, source);
            }
        }
    }
}
