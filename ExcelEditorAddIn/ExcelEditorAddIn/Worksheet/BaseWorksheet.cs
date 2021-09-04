using EeCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public class BaseWorksheet
    {
        public ElementType ElementType => Element.Type;
        public IElement Element { get; }
        public BaseWorkbook Workbook { get; }
        public Excel.Worksheet Worksheet { get; }

        protected List<(Excel.Range Cell, IElement Element)> Elements;

        public BaseWorksheet(IElement element, BaseWorkbook workbook, Excel.Worksheet worksheet)
        {
            Element = element;
            Workbook = workbook;
            Worksheet = worksheet;
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
    }
}
