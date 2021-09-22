using JkwExtensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public static class Utils
    {
        public static int ColumnNameToNumber(string columnName)
        {
            var result = Enumerable.Range(1, columnName.Length - 1)
                .Sum(i => (int)Math.Pow(26, i));

            result += columnName.ToUpper().Reverse()
                .Select((chr, i) => new { Chr = chr, Index = i })
                .Sum(x => (x.Chr - 'A') * (int)Math.Pow(26, x.Index));

            return result + 1;
        }

        public static string ColumnNumberToName(int columnNumber, string result = null)
        {
            columnNumber--;
            var newChar = new string(new[] { (char)(columnNumber % 26 + 'A') });
            result = (result ?? string.Empty).Insert(0, newChar);

            if (columnNumber < 26)
            {
                return result;
            }
            else
            {
                return ColumnNumberToName(columnNumber / 26, result);
            }
        }
    }
}
