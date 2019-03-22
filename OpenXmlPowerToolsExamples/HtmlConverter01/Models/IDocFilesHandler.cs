using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HtmlConverter01.Models
{
    interface IDocFilesHandler
    {
        BoolIdResult SetDocumentImageBase64(AbstractUpload upload);
    }
}
