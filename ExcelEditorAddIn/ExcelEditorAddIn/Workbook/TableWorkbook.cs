﻿using EeCommon;
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

        public TableWorkbook(ITableElement tableElement, string filePath)
            : base(tableElement, filePath)
        {
            TableElement = tableElement;

            WorkbookCreated += TableWorkbook_WorkbookCreated;
        }

        public override void Open()
        {
            MakeWorkbook();
            var metadata = OpenMetadata();

            MainWorksheet = new TableWorksheet(
                element: TableElement,
                workbook: this,
                worksheet: Workbook.SheetList().First(),
                path: "/",
                metadata: metadata);
            MainWorksheet.Changed += (s, a) => Dirty = true;

            Workbook.Activate();
        }

        private void TableWorkbook_WorkbookCreated(object sender, Excel.Workbook workbook)
        {
        }
    }
}
