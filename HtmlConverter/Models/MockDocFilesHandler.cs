using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HtmlConverter01.Models
{
    public class MockDocFilesHandler : IDocFilesHandler
    {
        static Random gen = new Random();
        public BoolIdResult SetDocumentImageBase64(AbstractUpload upload)
        {
            Console.WriteLine($"Image {upload} uploaded");
            return new BoolIdResult(true, gen.Next(-50000, -30000));
        }
    }
}
