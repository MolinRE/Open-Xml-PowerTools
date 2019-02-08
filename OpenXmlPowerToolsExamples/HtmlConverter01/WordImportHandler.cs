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
            var result = new List<DocLobbyFormat>();
            foreach (var text in paragraphs)
            {
                result.AddRange(WordImportDal.GetListLobbyByText(text));
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
            var result = new Dictionary<DocType, string>();
            foreach (var text in paragraphs)
            {
                foreach (var docType in WordImportDal.GetListTypeByText(text))
                {
                    result.Add(docType, text);
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
            var result = new List<RegionRf>();
            foreach (var text in queries)
            {
                result.AddRange(WordImportDal.GetRegionsByText(text));
            }

            return result;
        }
    }
}
