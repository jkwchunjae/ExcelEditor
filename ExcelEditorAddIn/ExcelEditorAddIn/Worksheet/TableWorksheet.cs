using EeCommon;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
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
            AttachEvents();
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

                for (var row = 2; row <= 1 + tableElement.Length; row++)
                {
                    for (var column = 1; column <= tableElement.Keys.Count; column++)
                    {
                        var cell = sheet.Cell(row, column);
                        var element = tableElement.Elements[row - 2, column - 1];
                        Elements.Add((cell, element));
                    }
                }
            }
        }

        private void AttachEvents()
        {
            Worksheet.BeforeDoubleClick += Worksheet_BeforeDoubleClick;
            Worksheet.Change += Worksheet_Change;
        }

        private void Worksheet_Change(Excel.Range Target)
        {
        }

        private void Worksheet_BeforeDoubleClick(Excel.Range Target, ref bool Cancel)
        {
            if (TryGetExistElement(Target, out var element))
            {
                if (element.Type == ElementType.Table)
                {
                    Cancel = true;
                }
                else if (element.Type == ElementType.Array)
                {
                    Cancel = true;
                }
                else if (element.Type == ElementType.Object)
                {
                    Cancel = true;
                }
                else
                {
                    Cancel = false;
                }
            }
        }
    }
}
