using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsparser
{
    class StringUtil
    {
        public static int ConvertToInt32(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0;
            }

            return Convert.ToInt32(value);
        }
    }
}
