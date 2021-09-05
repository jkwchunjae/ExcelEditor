using EeCommon;
using System;
using System.IO;
using System.Windows.Forms;
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

        protected bool Dirty = false;


        public event EventHandler<Excel.Workbook> WorkbookCreated;
        public event EventHandler<BaseWorkbook> Closed;

        public BaseWorkbook(IElement element, string jsonFilePath)
        {
            JsonFilePath = jsonFilePath;
            Element = element;
        }

        public virtual void Open() { }

        protected Excel.Workbook MakeWorkbook()
        {
            Workbook = Globals.ThisAddIn.Application.Workbooks.Add();
            CreateWorkbookFile();

            WorkbookCreated?.Invoke(this, Workbook);

            return Workbook;
        }

        private void CreateWorkbookFile()
        {
            var workbookPath = PathOf.TemporaryFilePath(Path.GetFileNameWithoutExtension(JsonFilePath));
            if (!Directory.Exists(PathOf.LocalRootDirectory))
            {
                Directory.CreateDirectory(PathOf.LocalRootDirectory);
            }

            if (File.Exists(workbookPath))
            {
                try
                {
                    File.Delete(workbookPath);
                }
                catch
                {
                    MessageBox.Show("Opened another file.");
                    throw;
                }
            }

            Workbook.SaveAs(workbookPath);
        }

        public void Save()
        {
            Workbook.Save();
        }

        public void AttachEvents()
        {
            Workbook.BeforeSave += Workbook_BeforeSave;
            Workbook.BeforeClose += Workbook_BeforeClose;
        }

        private void Workbook_BeforeSave(bool SaveAsUI, ref bool Cancel)
        {
            if (Dirty)
            {
                var text = Element.GetSaveText();
                File.WriteAllText(JsonFilePath, text);
            }
        }

        private void Workbook_BeforeClose(ref bool Cancel)
        {
            Closed?.Invoke(this, this);
        }
    }
}
