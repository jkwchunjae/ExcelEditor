using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public static class WorksheetExtensions
    {
        public static Excel.Range Cell(this Excel.Worksheet sheet, int row1, int column1)
        {
            return sheet.Cells[row1, column1];
        }
    }
}
