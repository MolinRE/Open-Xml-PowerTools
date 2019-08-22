using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebArmModels;
using WebArmModels.Upload;

namespace HtmlConverter01.Models
{
    interface IDocFilesHandler
    {
        BoolIdResult SetDocumentImageBase64(AbstractUpload upload);
    }
}
