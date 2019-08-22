using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxConverter
{
    public class WarmImage
    {
        public int ImgID { get; set; }
        public string ImageName { get; set; }
        public int MimeID { get; set; }
        public byte[] ImageData { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public Guid VerID { get; set; }
    }
}
