using EeCommon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public class TableWorksheet : BaseWorksheet
    {
        ITableElement TableElement { get; }
        public TableWorksheet(ITableElement element, BaseWorkbook workbook, Excel.Worksheet worksheet)
            : base(element, workbook, worksheet)
        {
            TableElement = element;

            SpreadElement(element);
        }

        private void SpreadElement(ITableElement tableElement)
        {
            var sheet = Worksheet;

            // title
            for (var column = 1; column <= tableElement.Keys.Count; column++)
            {
                var cell = sheet.Cell(1, column);
                cell.Value2 = tableElement.Keys[column - 1];
            }

            // values
            if (tableElement.Any)
            {
                var minCell = sheet.Cell(2, 1);
                var maxCell = sheet.Cell(1 + tableElement.Length, tableElement.Keys.Count);
                Excel.Range valuesRange = sheet.Range[minCell, maxCell];
                valuesRange.Value2 = tableElement.Values;
            }
        }
    }
}
