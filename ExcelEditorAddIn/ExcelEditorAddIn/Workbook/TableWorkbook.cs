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
    public class TableWorkbook : BaseWorkbook
    {
        public ITableDocument TableDocument { get; private set; }
        public TableWorkbook(ITableDocument document, string jsonFilePath)
            : base(document, jsonFilePath)
        {
            TableDocument = document;
        }

        public void OpenFile()
        {
            Workbook = Globals.ThisAddIn.Application.Workbooks.Add();
            MainWorksheet = Workbook.SheetList().First();

            var tableDocument = TableDocument;
            var book = Workbook;
            var sheet = MainWorksheet;

            // title
            for (var column = 1; column <= tableDocument.Keys.Count; column++)
            {
                Excel.Range cell = sheet.Cells[1, column];
                cell.Value2 = tableDocument.Keys[column - 1];
            }

            // values
            if (tableDocument.Any)
            {
                var minCell = sheet.Cell(2, 1);
                var maxCell = sheet.Cell(1 + tableDocument.Length, tableDocument.Keys.Count);
                Excel.Range valuesRange = sheet.Range[minCell, maxCell];
                valuesRange.Value2 = tableDocument.Values;
            }

            book.Activate();
            sheet.Activate();
        }
    }
}
