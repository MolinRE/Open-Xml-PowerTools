using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using System.Data;

namespace HtmlConverter01
{
    class WordImportDal
    {
        static string srv25 = "Data Source=srv25; Initial Catalog=KLM_2;Persist Security Info=True; user id=EditCAD; password=BwVzdZwrlJ";
        static string srv12 = "Data Source=SRV12;Initial Catalog=RBD_dev;User Id=service.webarm;PASSWORD=WhS7LIwtPKNO";

        /// <summary>
        /// Получает список органов по переданному тексту
        /// </summary>
        /// <param name="pText"></param>
        /// <returns></returns>
        public static List<DocLobbyFormat> GetListLobbyByText(string pText)
        {
            using (var conn = new SqlConnection(srv12))
            {
                var result = conn.Query<DocLobbyFormat>("webarm.WordImport_GetLobbyByText", 
                    new { lobbyName = new DbString() { Value = pText, IsFixedLength = false, Length = 256, IsAnsi = true } },
                    commandType: CommandType.StoredProcedure);

                return result.ToList();
            }
        }

        /// <summary>
        /// Загрузка списка органов по переданному тексту
        /// </summary>
        /// <param name="pText"></param>
        /// <returns></returns>
        public static List<DocType> GetListTypeByText(string pText)
        {
            using (var conn = new SqlConnection(srv12))
            {
                var result = conn.Query<DocType>("webarm.WordImport_GetDocTypeByText", 
                    new { typeName = new DbString() { Value = pText, IsFixedLength = false, Length = -1, IsAnsi = true } },
                    commandType: CommandType.StoredProcedure);

                return result.ToList();
            }
        }

        public static List<RegionRf> GetRegionsByText(string text)
        {
            using (var conn = new SqlConnection(srv12))
            {
                var result = conn.Query<RegionRf>("webarm.[WordImport_GetRegionRfByText]",
                    new { regionName = new DbString() { Value = text, IsFixedLength = false, Length = 256, IsAnsi = true } },
                    commandType: CommandType.StoredProcedure);

                return result.ToList();
            }
        }
    }
}