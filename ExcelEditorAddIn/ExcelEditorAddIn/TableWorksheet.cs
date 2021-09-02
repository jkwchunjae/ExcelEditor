using EeCommon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public class TableWorksheet
    {
        public void OpenFile(ITableDocument tableDocument, string jsonFilePath)
        {
            var book = Globals.ThisAddIn.Application.Workbooks.Add();
            var sheet = book.SheetList().First();

            // title
            for (var column = 1; column <= tableDocument.Keys.Count; column++)
            {
                Excel.Range cell = sheet.Cells[1, column];
                cell.Value2 = tableDocument.Keys[column - 1];
            }

            // values
            if (tableDocument.Any)
            {
                Excel.Range minCell = sheet.Cells[2, 1];
                Excel.Range maxCell = sheet.Cells[1 + tableDocument.Length, tableDocument.Keys.Count];
                Excel.Range valuesRange = sheet.get_Range(minCell.Address, maxCell.Address);
                valuesRange.Value2 = tableDocument.Values;
            }

            book.Activate();
            sheet.Activate();
        }
    }
}
