using EeCommon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public class BaseWorkbook
    {
        public DocumentType DocumentType => Document.Type;
        public string JsonFilePath { get; protected set; }
        public IDocument Document { get; protected set; }
        public Excel.Workbook Workbook { get; protected set; }
        public Excel.Worksheet MainWorksheet { get; protected set;  }

        public BaseWorkbook(IDocument document, string jsonFilePath)
        {
            JsonFilePath = jsonFilePath;
            Document = document;
        }
    }
}
