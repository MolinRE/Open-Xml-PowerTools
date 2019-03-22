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
        private static readonly string prodServerName = "srv25";
        private static readonly string devServerName = "srv12";
        private static readonly string prodDbName = "KLM_2";
        private static readonly string devDbName = "RBD_dev";
        private static readonly string userName = "k.komarov";
        private static readonly string userPassword = "w6q1r1q!";

        private static readonly bool useProd = false;

        internal static string DataSource
        {
            get { return useProd ? prodServerName : devServerName; }
        }

        internal static string InitialCatalog
        {
            get { return useProd ? prodDbName : devDbName; }
        }

        private static string ConnectionString
        {
            get
            {
                return string.Format("Data Source={0};Initial Catalog={1};User Id={2};Password={3};{4}",
                    DataSource, InitialCatalog, userName, userPassword,
                    useProd ? "Persist Security Info=True;" : "");
            }
        }

        /// <summary>
        /// Получает список органов по переданному тексту
        /// </summary>
        /// <param name="pText"></param>
        /// <returns></returns>
        public static List<DocLobbyFormat> GetListLobbyByText(string pText)
        {
            using (var conn = new SqlConnection(ConnectionString))
            {
                var result = conn.Query<DocLobbyFormat>("webarm.WordImport_GetLobbyByText", 
                    new
                    {
                        lobbyName = new DbString() { Value = pText, IsFixedLength = false, Length = 256, IsAnsi = true }
                    },
                    commandType: CommandType.StoredProcedure);

                //if (!result.Any())
                //{
                //    result = conn.Query<DocLobbyFormat>("webarm.WordImport_GetLobbyByText",
                //    new
                //    {
                //        lobbyName = new DbString() { Value = pText, IsFixedLength = false, Length = 256, IsAnsi = true },
                //        replace = true
                //    },
                //    commandType: CommandType.StoredProcedure);
                //}

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
            using (var conn = new SqlConnection(ConnectionString))
            {
                var result = conn.Query<DocType>("webarm.WordImport_GetDocTypeByText", 
                    new { typeName = new DbString() { Value = pText, IsFixedLength = false, Length = -1, IsAnsi = true } },
                    commandType: CommandType.StoredProcedure);

                return result.ToList();
            }
        }

        public static List<RegionRf> GetRegionsByText(string text)
        {
            using (var conn = new SqlConnection(ConnectionString))
            {
                var result = conn.Query<RegionRf>("webarm.[WordImport_GetRegionRfByText]",
                    new { regionName = new DbString() { Value = text, IsFixedLength = false, Length = 256, IsAnsi = true } },
                    commandType: CommandType.StoredProcedure);

                return result.ToList();
            }
        }
    }
}