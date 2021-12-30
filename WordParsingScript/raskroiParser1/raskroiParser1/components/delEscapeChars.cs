using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace raskroiParser1.components
{
    public static class delEscapeChars
    {
        static List<string> patterns = new List<string>
        {
            @"\r",
            @"\a",
            @"\t",
            @"\v",
            @"\f",
            @"\n"
        };
        public static void delEscapeCharsFunc(ref string str)
        {
            string target = "";
            string curStr = str;

            foreach (string pattern in patterns)
            {
                System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(pattern);
                curStr = reg.Replace(curStr, target);
            }
            str = curStr;
        }
    }
}
