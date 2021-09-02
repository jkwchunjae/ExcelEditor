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

        public BaseWorksheet(IElement element, BaseWorkbook workbook, Excel.Worksheet worksheet)
        {
            Element = element;
            Workbook = workbook;
            Worksheet = worksheet;
        }
    }
}
