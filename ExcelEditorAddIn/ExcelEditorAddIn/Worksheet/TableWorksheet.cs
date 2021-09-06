using EeCommon;
using JkwExtensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public class TableWorksheet : BaseWorksheet
    {
        ITableElement TableElement { get; }

        List<(string PropertyName, ElementType ElementType)> ColumnPropertyInfo = new List<(string PropertyName, ElementType ElementType)>();

        public TableWorksheet(ITableElement element, BaseWorkbook workbook, Excel.Worksheet worksheet)
            : base(element, workbook, worksheet)
        {
            TableElement = element;

            SpreadElement();
            AttachEvents();
        }

        private void SpreadElement()
        {
            var sheet = Worksheet;
            var table = TableElement;

            // title
            ColumnPropertyInfo = table.Properties.ToList();

            ColumnPropertyInfo.ForEach((propertyInfo, index) =>
            {
                var column = index + 1;
                var cell = sheet.Cell(1, column);
                cell.Value2 = propertyInfo.PropertyName;
            });

            // values
            if (table.Any)
            {
                var minCell = sheet.Cell(2, 1);
                var maxCell = sheet.Cell(1 + table.Length, table.Properties.Count());
                Excel.Range valuesRange = sheet.Range[minCell, maxCell];
                var values = table.GetExcelArray(ColumnPropertyInfo.Select(x => x.PropertyName));
                valuesRange.Value2 = values;
            }

            Elements = table.Elements
                .Select((obj, i) => new { IObject = obj, Index = i })
                .SelectMany(x => x.IObject.Properties.Select(property =>
                {
                    var propertyIndex = ColumnPropertyInfo.FindIndex(e => e.PropertyName == property.Key);
                    var column = propertyIndex + 1;
                    var row = x.Index + 2;
                    var cell = sheet.Cell(row, column);
                    return (cell, property.Value);
                }))
                .ToList();
        }

        private void AttachEvents()
        {
            Worksheet.BeforeDoubleClick += Worksheet_BeforeDoubleClick;
            Worksheet.Change += Worksheet_Change;
        }

        private void Worksheet_Change(Excel.Range Target)
        {
            if (IsInArea(Target) == false)
            {
                // 영역 바깥을 수정한 경우
                return;
            }

            object previousValue = null;

            try
            {
                if (TryGetExistElement(Target, out var element))
                {
                    // 이미 있는 값을 수정하는 경우
                    if (element.Type == ElementType.Value)
                    {
                        var valueElement = (IValueElement)element;
                        previousValue = valueElement.GetExcelValue();
                        valueElement.UpdateValue((object)Target.Value, (object)Target.Value2);
                        OnChange();
                    }
                    else // Array, Object, Table
                    {
                    }
                }
                else if (TryGetElementInfo(Target, out var parentObject, out var fieldName, out var elementType))
                {
                    // 없던 값을 새로 쓰는 경우
                    // 부모 객체 정보를 얻어온다.
                    if (elementType == ElementType.Value)
                    {
                        IElement newValue = TableElement.CreateValueElement((object)Target.Value, (object)Target.Value2);
                        parentObject.Add(fieldName, newValue);
                        OnChange();
                    }
                    else // Array, Object, Table
                    {
                    }
                }
            }
            catch (RequireNumberException)
            {
                MessageBox.Show("숫자 형식을 입력해야 합니다.");
                Target.Value = previousValue;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
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

        private bool TryGetElementInfo(Excel.Range cell,
            out IObjectElement objectElement,
            out string propertyName,
            out ElementType elementType)
        {
            if (IsInArea(cell))
            {
                objectElement = TableElement.Elements.ElementAt(cell.Row - 2);
                (propertyName, elementType) = ColumnPropertyInfo[cell.Column - 1];

                return true;
            }
            objectElement = null;
            propertyName = null;
            elementType = ElementType.Null;
            return false;
        }

        private bool IsInArea(Excel.Range cell)
        {
            // Elements 최대 최소 범위 안에 들어있어야 함.
            var minRow = Elements.Min(x => x.Cell.Row);
            var maxRow = Elements.Max(x => x.Cell.Row);
            var minColumn = Elements.Min(x => x.Cell.Column);
            var maxColumn = Elements.Max(x => x.Cell.Column);

            if (cell.Row >= minRow && cell.Row <= maxRow)
            {
                if (cell.Column >= minColumn && cell.Column <= maxColumn)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
