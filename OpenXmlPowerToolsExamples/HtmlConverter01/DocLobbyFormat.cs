using System;

namespace HtmlConverter01
{
    public class DocLobbyFormat
    {
        /// <summary>
        /// ID.
        /// </summary>
        public int ID { get; set; }

        /// <summary>
        /// Название.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Название в шапке оригинала.
        /// </summary>
        public string NameHeaderOriginal { get; set; }

        /// <summary>
        /// Название в карточку (DocName)
        /// </summary>
        public string NameToCard { get; set; }

        /// <summary>
        /// Название в шапку
        /// </summary>
        public string NameHeader
        {
            get { return Name.ToUpper(); }
        }

        /// <summary>
        /// Относится к судебке
        /// </summary>
        public bool IsArbitr { get; set; }

        public int RegionID { get; set; }

        public override string ToString()
        {
            return Name + " (" + ID + ")";
        }
    }

    /// <summary>
    /// Базовый класс для словарей
    /// </summary>    
    public class BaseDictElement
    {
        public BaseDictElement()
        {
        }

        public BaseDictElement(int aId)
        {
            this.ID = aId;
        }

        public BaseDictElement(int aId, string aName)
        {
            this.ID = aId;
            this.Name = aName;
        }

        /// <summary>
        /// ID
        /// </summary>
        public int ID { get; set; }

        /// <summary>
        /// Название
        /// </summary>
        public string Name { get; set; }

        public override bool Equals(Object value)
        {
            if (value is BaseDictElement == false) return false;
            BaseDictElement dic = (BaseDictElement)value;
            return (dic.ID == this.ID);
        }

        public override int GetHashCode()
        {
            return this.ID.GetHashCode();
        }

        public override string ToString()
        {
            return Name + " (" + ID + ")";
        }
    }

    /// <summary>
    /// Тип документа
    /// </summary>
    public class DocType : BaseDictElement
    {
        internal static DocType[] Common = GetDocTypes();

        public DocType() : base() { }

        public DocType(int aId) : base(aId) { }

        /// <summary>
        /// Наименование типа документа
        /// </summary>
        public new string Name { get; set; }

        /// <summary>
        /// Дата создания
        /// </summary>
        public DateTime? CreateDate { get; set; }

        /// <summary>
        /// Дата изменения
        /// </summary>
        public DateTime? ModifyDate { get; set; }

        /// <summary>
        /// Признак наличия связей с документами
        /// </summary>
        public bool HasLnk;

        /// <summary>
        /// Наличие связей с НПД
        /// </summary>
        public string HasDocLnk
        {
            get
            {
                if (this.HasLnk)
                {
                    return "+";
                }
                else
                {
                    return String.Empty;
                }
            }
            set { }
        }

        /// <summary>
        /// Наименование (большой лист)
        /// </summary>
        public string NameForBL { get; set; }

        /// <summary>
        /// Наименование (маленький лист)
        /// </summary>
        public string NameForSL { get; set; }

        /// <summary>
        /// Наименование (бэклинки)
        /// </summary>
        public string NameForBacklink { get; set; }

        /// <summary>
        /// Название во множественном числе и творительном падеже
        /// </summary>
        public string NamePluralForRedactions { get; set; }

        /// <summary>
        /// Название в блок "сведения об изменяюших" (творительный падеж)
        /// </summary>
        public string NameForRedactions { get; set; }


        /// <summary>
        /// Название в родительного падеже
        /// </summary>
        public string NameR { get; set; }

        internal static DocType[] GetDocTypes()
        {
            return new DocType[]
            {
                new DocType() { Name = "закон", NameForRedactions = "Законом", NamePluralForRedactions = "Законов", NameR = "закона" },
                new DocType() { Name = "указ", NameForRedactions = "Указом", NamePluralForRedactions = "Указов", NameR = "указа" },
                new DocType() { Name = "приказ", NameForRedactions = "приказом", NamePluralForRedactions = "приказов", NameR = "приказа" },
                new DocType() { Name = "постановление", NameForRedactions = "постановлением", NamePluralForRedactions = "постановлений", NameR = "постановления" },
                new DocType() { Name = "распоряжение", NameForRedactions = "распоряжением", NamePluralForRedactions = "распоряжений", NameR = "распоряжения" },
                new DocType() { Name = "решение", NameForRedactions = "решением", NamePluralForRedactions = "решений", NameR = "решения" },
                new DocType() { Name = "нормативный правовой акт", NameForRedactions = "нормативным правовым актом", NamePluralForRedactions = "нормативных правовых актов", NameR = "нормативного правового акта" },
                new DocType() { Name = "протокол", NameForRedactions = "протоколом", NamePluralForRedactions = "протоколов", NameR = "протокола" },
                new DocType() { Name = "соглашение", NameForRedactions = "соглашением", NamePluralForRedactions = "соглашений", NameR = "соглашения" },
                new DocType() { Name = "договор", NameForRedactions = "договором", NamePluralForRedactions = "договоров", NameR = "договора" },
                new DocType() { Name = "международный договор", NameForRedactions = "международным договором", NamePluralForRedactions = "международных договоров", NameR = "международного договора" },
                new DocType() { Name = "изменение", NameForRedactions = "изменением", NamePluralForRedactions = "изменений", NameR = "изменения" },
                new DocType() { Name = "дополнение", NameForRedactions = "дополнением", NamePluralForRedactions = "дополнений", NameR = "дополнения" },
                new DocType() { Name = "письмо", NameForRedactions = "письмом", NamePluralForRedactions = "писем", NameR = "письма" },
                new DocType() { Name = "совместное письмо", NameForRedactions = "совместным письмом", NamePluralForRedactions = "совместных писем", NameR = "совместного письма" },
                new DocType() { Name = "федеральный закон", NameForRedactions = "Федеральным законом", NamePluralForRedactions = "федеральных законов", NameR = "федерального закона" },
                new DocType() { Name = "федеральный конституционный закон", NameForRedactions = "Федеральным конституционным законом", NamePluralForRedactions = "федеральных конституционных законов", NameR = "федерального конституационального закона" },
                new DocType() { Name = "международный протокол", NameForRedactions = "международным протоколом", NamePluralForRedactions = "международных протоколов", NameR = "международного прокотола" },
                new DocType() { Name = "отраслевое соглашение", NameForRedactions = "отраслевым соглашением", NamePluralForRedactions = "отраслевых соглашений", NameR = "отраслевого соглашения" },
                new DocType() { Name = "указание", NameForRedactions = "указанием", NamePluralForRedactions = "указаний", NameR = "указания" }
            };
        }

        public override string ToString()
        {
            return this.Name + " (" + this.ID + ")";
        }
    }

    public class RegionRf
    {
        /// <summary>
        /// Идентификатор
        /// </summary>
        public int RegionID { get; set; }

        /// <summary>
        /// Название
        /// </summary>
        public string RegionName { get; set; }

        /// <summary>
        /// Код региона
        /// </summary>
        public int RegionCode { get; set; }

        public bool Active { get; set; }

        public string RegionNameR
        {
            get
            {
                return RegionNameForDoc.ToUpper();
            }
        }

        public string RegionNameP { get; set; }

        public string RegionNameForDoc { get; set; }

        public override string ToString()
        {
            return RegionName + " (" + RegionID + ")";
        }
    }
}
