using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.IO;
using EeJson;
using EeCommon;

namespace ExcelEditorAddIn
{
    public partial class ThisAddIn
    {
        List<BaseWorkbook> _workbookData = new List<BaseWorkbook>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private bool AlreadyOpened(string jsonFilePath, out BaseWorkbook workbookData)
        {
            var found = _workbookData.FirstOrDefault(x => x.JsonFilePath == jsonFilePath);
            workbookData = null;
            if (found != null)
            {
                workbookData = found;
            }
            return found != null;
        }

        public void OpenFile(OpenFileDialog openFileDialog)
        {
            // TODO: json syntax error
            var filePath = openFileDialog.FileName;

            if (AlreadyOpened(filePath, out var workbookData))
            {
                workbookData.Workbook.Activate();
                return;
            }

            using (var reader = new StreamReader(openFileDialog.OpenFile()))
            {
                var jsonText = reader.ReadToEnd();
                var baseDocument = new JsonBaseDocument(jsonText);
                if (baseDocument.Type == DocumentType.Table)
                {
                    var jsonTableDocument = new JsonTableDocument(baseDocument);
                    var tableWorksheet = new TableWorkbook(jsonTableDocument, filePath);
                    _workbookData.Add(tableWorksheet);
                    tableWorksheet.OpenFile();
                }
            }
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
