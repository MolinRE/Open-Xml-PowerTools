using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using System.Data;

namespace HtmlConverter
{
    public class WordImportDal
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

        public static List<DocLobbyFormatDto> Lobbies { get; set; }

        public static List<DocLobbyFormatDto> GetAllLobbies()
        {
            using (var conn = new SqlConnection(ConnectionString))
            {
                var sql = @"
SELECT dl.LobbyID as ID
, dl.LobbyNameFull as NameHeaderOriginal
, ISNULL(dl.LobbyNameRToDocName, dl.LobbyNameFull) as NameToCard
, dl.LobbyNameFull as NameHeader
, dl.LobbyNameFull as [Name]
, rll.RegionID
FROM dbo.DocLobby dl
LEFT JOIN dbo.RegionLobbyLnk rll
	ON rll.LobbyID = dl.LobbyID";

                var result = conn.Query<DocLobbyFormatDto>(sql);

                return result.ToList();
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
                List<DocLobbyFormatDto> lobbies = null;
                if (Lobbies != null)
                {
                    lobbies = Lobbies.Where(p => SearchLobbyPredicate(p, pText)).ToList();
                }
                else
                {
                    lobbies = conn.Query<DocLobbyFormatDto>("webarm.WordImport_GetLobbyByText",
                        new
                        {
                            lobbyName = new DbString() { Value = pText, IsFixedLength = false, Length = 256, IsAnsi = true }
                        },
                        commandType: CommandType.StoredProcedure).ToList();
                }

                var result = new List<DocLobbyFormat>();
                foreach (var lobby in lobbies)
                {
                    var item = result.FirstOrDefault(p => p.ID == lobby.ID);
                    if (item == null)
                    {
                        item = new DocLobbyFormat();
                        item.ID = lobby.ID;
                        item.NameHeader = lobby.NameHeader;
                        item.NameHeaderOriginal = lobby.NameHeaderOriginal;
                        item.NameToCard = lobby.NameToCard;

                        item.RegionsIds = new List<int?>();
                        result.Add(item);
                    }

                    item.RegionsIds.Add(lobby.RegionID);
                }

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

        public static bool SearchLobbyPredicate(DocLobbyFormatDto lobby, string text)
        {
            return lobby.NameHeader.Equals(text, StringComparison.InvariantCultureIgnoreCase)
                || lobby.NameHeader.Equals(text.ToUpper().Replace("РОССИЙСКОЙ ФЕДЕРАЦИИ", "РФ"), StringComparison.InvariantCultureIgnoreCase)
                || lobby.NameHeader.Equals(text + " РФ", StringComparison.InvariantCultureIgnoreCase);
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