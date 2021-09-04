using EeCommon;
using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public class BaseWorkbook
    {
        public ElementType ElementType => Element.Type;
        public string JsonFilePath { get; protected set; }
        public IElement Element { get; protected set; }
        public Excel.Workbook Workbook { get; protected set; }
        public BaseWorksheet MainWorksheet { get; protected set;  }

        public event EventHandler<Excel.Workbook> WorkbookCreated;

        public BaseWorkbook(IElement element, string jsonFilePath)
        {
            JsonFilePath = jsonFilePath;
            Element = element;
        }

        public virtual void Open() { }

        protected Excel.Workbook MakeWorkbook()
        {
            Workbook = Globals.ThisAddIn.Application.Workbooks.Add();

            WorkbookCreated?.Invoke(this, Workbook);
            AttachEvents();

            return Workbook;
        }

        private void AttachEvents()
        {
            Workbook.BeforeSave += Workbook_BeforeSave;
        }

        private void Workbook_BeforeSave(bool SaveAsUI, ref bool Cancel)
        {
            var text = Element.GetSaveText();
            File.WriteAllText(JsonFilePath, text);
        }
    }
}
