using EeCommon;
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
        private static readonly string RecentLabel = "Recents";

        private void ExcelEditorRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            InitRecents();
        }

        private void JsonOpenButton_Click(object sender, RibbonControlEventArgs e)
        {
            openFileDialog1.Filter = "json file (*.json)|*.json";
            var result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK || result == DialogResult.Yes)
            {
                Globals.ThisAddIn.OpenFile(openFileDialog1.FileName);
            }
        }

        #region Recents
        private void InitRecents()
        {
            RecentsDropdown.SelectionChanged += RecentsDropdown_SelectionChanged;
            RecentsDropdown.Label = string.Empty;

            LoadRecents();

            Recents.ItemUpdated += (_, __) => LoadRecents();
        }

        private void LoadRecents()
        {
            RecentsDropdown.Items.Clear();

            var emptyItem = Factory.CreateRibbonDropDownItem();
            emptyItem.Label = RecentLabel;
            RecentsDropdown.Items.Add(emptyItem);

            foreach (var recentData in Recents.Items)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = recentData.FilePath;
                RecentsDropdown.Items.Add(item);
            }
        }

        private void RecentsDropdown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            var filePath = RecentsDropdown.SelectedItem.Label;
            if (filePath != RecentLabel)
            {
                Globals.ThisAddIn.OpenFile(filePath);
            }
        }
        #endregion
    }
}
