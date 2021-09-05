using EeCommon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEditorAddIn
{
    public class TableWorkbook : BaseWorkbook
    {
        public ITableElement TableElement { get; private set; }
        public TableWorkbook(ITableElement tableElement, string jsonFilePath)
            : base(tableElement, jsonFilePath)
        {
            TableElement = tableElement;

            WorkbookCreated += TableWorkbook_WorkbookCreated;
        }

        public override void Open()
        {
            MakeWorkbook();

            MainWorksheet = new TableWorksheet(TableElement, this, Workbook.SheetList().First());

            Workbook.Activate();
        }

        private void TableWorkbook_WorkbookCreated(object sender, Excel.Workbook workbook)
        {
        }
    }
}
