using System.IO;

namespace HtmlConverter01.Models
{

    public abstract class AbstractUpload
    {
        public string Key { get; set; }

        public int? Id { get; set; }
        

        public string Name { get; set; }

        public string FileName { get; set; }

        /// <summary>
        /// Возвращает расширение файла, включая точку.
        /// </summary>
        public string FileExtension { get { return Path.GetExtension(FileName); } }

        /// <summary>
        /// Возвращает имя файла без расширения.
        /// </summary>
		public string FileNameWithoutExtension { get { return Path.GetFileNameWithoutExtension(FileName); } }

        public string ContentType { get; set; }

        public int? MimeTypeId { get; set; }

        protected internal bool CheckKey = false;
        

        public override string ToString()
        {
            return string.Format("{0} ({1})", Name, FileName);
        }

    }
}
