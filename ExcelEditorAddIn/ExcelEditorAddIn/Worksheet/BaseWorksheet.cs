using EeCommon;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ExcelEditorAddIn.ColumnSetting;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public class BaseWorksheet
    {
        public ElementType ElementType => Element.Type;
        public IElement Element { get; }
        public BaseWorkbook Workbook { get; }
        public Excel.Worksheet Worksheet { get; }
        protected string Path { get; }
        protected CommandBars CommandBars => Globals.ThisAddIn.Application.CommandBars;
        protected Metadata Metadata { get; }

        protected List<(Excel.Range Cell, IElement Element)> Elements;

        public event EventHandler Changed;

        protected ColumnSetting _columnSetting;

        public BaseWorksheet(IElement element, BaseWorkbook workbook, Excel.Worksheet worksheet, string path, Metadata metadata)
        {
            Element = element;
            Workbook = workbook;
            Worksheet = worksheet;
            Metadata = metadata;
            Path = path;

            AttachEvents_Base();
        }

        private void AttachEvents_Base()
        {
            Worksheet.BeforeRightClick += Worksheet_BeforeRightClick_Base;
        }

        private void Worksheet_BeforeRightClick_Base(Excel.Range Target, ref bool Cancel)
        {
            var contextMenuInfo = ContextMenuInfo.Make(Target.Address);

            if (contextMenuInfo == null)
            {
                Cancel = true;
                return;
            }
            switch (contextMenuInfo)
            {
                case SingleCellMenuInfo info:
                    Cancel = BeforeSingleCellRightClick(info);
                    break;
                case CellsMenuInfo info:
                    Cancel = BeforeCellsRightClick(info);
                    break;
                case ColumnMenuInfo info:
                    Cancel = BeforeColumnRightClick(info);
                    break;
                case RowMenuInfo info:
                    Cancel = BeforeRowRightClick(info);
                    break;
            }
        }

        protected virtual bool BeforeSingleCellRightClick(SingleCellMenuInfo info)
        {
            return false;
        }

        protected virtual bool BeforeCellsRightClick(CellsMenuInfo info)
        {
            return false;
        }

        protected virtual bool BeforeColumnRightClick(ColumnMenuInfo info)
        {
            return false;
        }

        protected virtual bool BeforeRowRightClick(RowMenuInfo info)
        {
            return false;
        }

        public virtual void UpdateMetadata()
        {
        }

        protected bool TryGetElement(Excel.Range cell, out IElement element)
        {
            if (Elements.Any(x => x.Cell.Address == cell.Address))
            {
                (_, element) = Elements.First(x => x.Cell.Address == cell.Address);
                return true;
            }
            element = null;
            return false;
        }

        protected bool TryGetExistElement(Excel.Range cell, out IElement element)
        {
            if (TryGetElement(cell, out element))
            {
                if (element != null)
                {
                    return true;
                }
            }
            element = null;
            return false;
        }

        protected bool IsInArea(Excel.Range cell)
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

        protected void OnChange()
        {
            Changed?.Invoke(this, null);
        }
    }
}
