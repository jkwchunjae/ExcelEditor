using EeCommon;
using EeJson;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelEditorAddIn
{
    public partial class ExcelEditorRibbon
    {
        private void ExcelEditorRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void JsonOpenButton_Click(object sender, RibbonControlEventArgs e)
        {
            openFileDialog1.Filter = "json file (*.json)|*.json";
            var result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK || result == DialogResult.Yes)
            {
                Globals.ThisAddIn.OpenFile(openFileDialog1);
            }
        }
    }
}
