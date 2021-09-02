using EeCommon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public class BaseWorkbook
    {
        public ElementType ElementType => Element.Type;
        public string JsonFilePath { get; protected set; }
        public IElement Element { get; protected set; }
        public Excel.Workbook Workbook { get; protected set; }
        public Excel.Worksheet MainWorksheet { get; protected set;  }

        public BaseWorkbook(IElement element, string jsonFilePath)
        {
            JsonFilePath = jsonFilePath;
            Element = element;
        }
    }
}
