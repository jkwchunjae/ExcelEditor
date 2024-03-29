﻿using System;
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
            var found = _workbookData.FirstOrDefault(x => x.FilePath == jsonFilePath);
            workbookData = null;
            if (found != null)
            {
                workbookData = found;
            }
            return found != null;
        }

        public void OpenFile(string filePath)
        {
            // TODO: json syntax error
            Recents.Update(filePath);

            if (AlreadyOpened(filePath, out var workbookData))
            {
                try
                {
                    workbookData.Workbook.Activate();
                    return;
                }
                catch (Exception ex)
                {
                    _workbookData.Remove(workbookData);
                    MessageBox.Show(ex.Message);
                }
            }

            var text = File.ReadAllText(filePath);

            // if json format
            OpenJson(text, filePath);
        }

        private void OpenJson(string jsonText, string filePath)
        {
            var baseElement = new JsonBaseElement(jsonText);

            if (baseElement.Type == ElementType.Table)
            {
                var jsonTableElement = new JsonTableElement(baseElement);
                OpenJsonTable(jsonTableElement, filePath);
            }
        }

        private void OpenJsonTable(JsonTableElement jsonTableElement, string jsonFilePath)
        {
            var workbookData = new TableWorkbook(jsonTableElement, jsonFilePath);
            workbookData.Open();
            workbookData.Save();
            workbookData.AttachEvents();
            workbookData.Closed += (_s, wData) => _workbookData.Remove(wData);

            _workbookData.Add(workbookData);
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
