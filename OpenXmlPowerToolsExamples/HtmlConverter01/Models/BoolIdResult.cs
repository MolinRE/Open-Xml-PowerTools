using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HtmlConverter01.Models
{

    /// <summary>Логический результат</summary>
	public class BoolResult
    {
        public bool Result { get; set; }

        public BoolResult()
        {

        }

        public BoolResult(bool result)
        {
            Result = result;
        }

        public static implicit operator bool(BoolResult result)
        {
            return result.Result;
        }
    }

    /// <summary>
    /// Логический результат с 32-битным Id.
    /// </summary>
	public class BoolIdResult : BoolResult
    {
        public int? Id { get; set; }

        public BoolIdResult()
        {

        }

        public BoolIdResult(bool result)
            : base(result)
        {

        }

        public BoolIdResult(bool result, int id)
        {
            Result = result;
            Id = id;
        }

        public override string ToString()
        {
            return Result && Id.HasValue
                 ? Id.ToString()
                 : Result.ToString();
        }
    }
}
