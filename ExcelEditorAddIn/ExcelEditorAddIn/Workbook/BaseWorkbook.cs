using EeCommon;
using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public class BaseWorkbook
    {
        public ElementType ElementType => Element.Type;
        public string FilePath { get; protected set; }
        public IElement Element { get; protected set; }
        public Excel.Workbook Workbook { get; protected set; }
        public BaseWorksheet MainWorksheet { get; protected set;  }

        protected Metadata Metadata { get; private set; }

        protected bool Dirty = false;


        public event EventHandler<Excel.Workbook> WorkbookCreated;
        public event EventHandler<BaseWorkbook> Closed;

        public BaseWorkbook(IElement element, string filePath)
        {
            FilePath = filePath;
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
            var workbookPath = PathOf.TemporaryFilePath(Path.GetFileNameWithoutExtension(FilePath));
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

        protected bool TryOpenMetadata(out Metadata metadata)
        {
            var metadataPath = PathOf.MetadataFilePath(FilePath);

            if (!File.Exists(metadataPath))
            {
                metadata = null;
                return false;
            }

            try
            {
                var metadataText = File.ReadAllText(metadataPath, Encoding.UTF8);
                metadata = Element.Deserialize<Metadata>(metadataText);
                Metadata = metadata;
                return true;
            }
            catch
            {
                metadata = null;
                return false;
            }
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

        private void DetachEvents()
        {
            Workbook.BeforeSave -= Workbook_BeforeSave;
            Workbook.BeforeClose -= Workbook_BeforeClose;
        }

        private void Workbook_BeforeSave(bool SaveAsUI, ref bool Cancel)
        {
            if (Dirty)
            {
                var text = Element.GetSaveText();
                File.WriteAllText(FilePath, text);

                if (Metadata != null)
                {
                    File.WriteAllText(PathOf.MetadataFilePath(FilePath), Element.Serialize(Metadata));
                }

                Dirty = false;
            }
        }

        private void Workbook_BeforeClose(ref bool Cancel)
        {
            if (Dirty)
            {
                var result = MessageBox.Show("저장하지 않은 변경내역이 있습니다.\r\n저장하시겠습니까?", "Excel Editor", MessageBoxButtons.YesNoCancel);
                if (result == DialogResult.Yes)
                {
                    Save();
                    Closed?.Invoke(this, this);
                }
                else if (result == DialogResult.No)
                {
                    DetachEvents();
                    Save();
                    Cancel = false;
                    Closed?.Invoke(this, this);
                }
                else // if (result == DialogResult.Cancel)
                {
                    Cancel = true;
                }
            }
            else
            {
                Closed?.Invoke(this, this);
            }
            //Closed?.Invoke(this, this);
        }
    }
}
