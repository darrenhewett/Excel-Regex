using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System.Text.RegularExpressions;

namespace Excel_Regex
{
    public static class xRegex
    {
        [ExcelFunction(Description = "Test regular expression matches. \n" + 
            "Refer to MS Regular Expression Language Reference for more details")]
        public static bool RegexMatches([ExcelArgument("Input string")] string input,
            [ExcelArgument("Regex Pattern")] string pattern)
        {
            Regex rgx = new Regex(pattern);
            MatchCollection matches = rgx.Matches(input);
            if (matches.Count > 0)
                return true;
            else
                return false;
        }

        [ExcelFunction(Description = "Get regular expression match. \n" +
            "Refer to MS Regular Expression Language Reference for more details")]
        public static object RegexExtract([ExcelArgument("Input string")]string input,
            [ExcelArgument("Regex Pattern")] string pattern,
            [ExcelArgument("(Optional) group number to return. Default = 1.")]
            Object matchNo)
        {
            int n = 0;
            if (matchNo is ExcelMissing)
                n = 1; // default match
            else
            {
                try
                {
                    n = Convert.ToInt32(matchNo);
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: '{0}'", e);
                    return ExcelError.ExcelErrorNum;
                }
            }
            
            Regex rgx = new Regex(pattern);
            MatchCollection matches = rgx.Matches(input);

            // return error if requested group number doesn't exist
            if (n > 1 && matches[0].Groups.Count <= n)
                return ExcelError.ExcelErrorNum;

            // Return the first group if groups were found, otherwise return the match
            if (matches[0].Groups.Count > 1) 
                return matches[0].Groups[n].Value.ToString();
            else
                return matches[0].Value.ToString();
        }
    }

    public class AddIn : IExcelAddIn
    {

        public void AutoOpen()
        {
            IntelliSenseServer.Register();
        }

        public void AutoClose()
        {
        }
    }
}
