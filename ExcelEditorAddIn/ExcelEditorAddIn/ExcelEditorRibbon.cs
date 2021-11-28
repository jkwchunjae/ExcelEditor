using EeCommon;
using JkwExtensions;
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
            //InitFavorites();
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

            Recents.ItemUpdated += (_, recents) => LoadRecents(recents);
        }

        private void LoadRecents(List<RecentItem> items = null)
        {
            RecentsDropdown.Items.Clear();

            var emptyItem = Factory.CreateRibbonDropDownItem();
            emptyItem.Label = RecentLabel;
            RecentsDropdown.Items.Add(emptyItem);

            foreach (var recentData in items ?? Recents.Items)
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

        #region Favorites
        private void InitFavorites()
        {
            LoadFavorites();

            Favorites.ItemUpdated += (_, favorites) => LoadFavorites(favorites);

            Recents.ItemUpdated += (_, recents) =>
            {
                var favorites = Favorites.Items;
                var remainItems = recents
                    .Where(rItem => favorites.Empty(f => f.FilePath == rItem.FilePath))
                    .ToList();
            };
        }

        private void LoadFavorites(List<FavoriteItem> items = null)
        {
            if (FavoriteGroup.Items.Any())
            {
                FavoriteGroup.Items.Clear();
            }

            foreach (var favorite in items ?? Favorites.Items)
            {
                var button = Factory.CreateRibbonButton();
                button.Label = favorite.Nickname;
                button.ScreenTip = favorite.FilePath;
                button.Click += (_, __) => Globals.ThisAddIn.OpenFile(favorite.FilePath);
                FavoriteGroup.Items.Add(button);
            }
        }
        #endregion
    }
}
