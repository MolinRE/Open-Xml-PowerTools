using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HtmlConverter01.Models
{
    public class UploadImageBase64 : AbstractUpload
    {
        public int? Width { get; set; }
        public int? Height { get; set; }
        public string ImageBase64 { get; set; }

        public UploadImageBase64()
        {

        }

        public UploadImageBase64(byte moduleId, int id)
        {
        }
    }
}
