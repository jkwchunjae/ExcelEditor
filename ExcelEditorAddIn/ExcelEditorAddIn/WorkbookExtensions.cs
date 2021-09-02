using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public static class WorkbookExtensions
    {
        public static IEnumerable<Excel.Worksheet> SheetList(this Excel.Workbook book)
        {
            foreach (Excel.Worksheet sheet in book.Sheets)
            {
                yield return sheet;
            }
        }

        public static IEnumerable<Excel.Style> StyleList(this Excel.Workbook book)
        {
            foreach (Excel.Style style in book.Styles)
            {
                yield return style;
            }
        }
    }
}
