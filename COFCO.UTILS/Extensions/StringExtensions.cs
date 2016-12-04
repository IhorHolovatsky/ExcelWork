using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace COFCO.UTILS.Extensions
{
    /// <summary>
    /// Contains extensions methods
    /// </summary>
    public static class StringExtensions
    {
        /// <summary>
        /// Parses string to int?
        /// </summary>
        public static int? ParseToInt(this string inputString)
        {
            if (string.IsNullOrWhiteSpace(inputString))
            {
                return null;
            }

            int returnValue;

            if (int.TryParse(inputString, out returnValue))
            {
                return returnValue;
            }

            return null;
        }

    }
}
