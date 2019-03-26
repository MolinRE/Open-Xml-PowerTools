using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HtmlConverter
{
    public static class StringExtensions
    {
        /// <summary>
        /// Удаляет из конца строки шифт-энтер
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string TrimNewLine(this string s)
        {
            if (s.EndsWith("#x200e"))
            {
                s = s.Remove(s.IndexOf("#x200e"));
            }

            return s;
        }
    }
}
