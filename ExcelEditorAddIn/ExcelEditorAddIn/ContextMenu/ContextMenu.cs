using JkwExtensions;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelEditorAddIn
{
    public interface IContextMenu
    {
        string Name { get; }

        void Show();
    }

    public class ContextMenu_Column : IContextMenu
    {
        public event EventHandler AddProperty;
        public event EventHandler RemoveProperty;

        public string Name { get; private set; }

        private CommandBar _commandBar;

        public ContextMenu_Column(CommandBar commandBar)
        {
            Name = nameof(ContextMenu_Column);

            _commandBar = commandBar;
            CreateButtons();
        }

        private void CreateButtons()
        {
            var addPropertyButton = (CommandBarButton)_commandBar.Controls.Add(
                Type: MsoControlType.msoControlButton,
                Temporary: true);
            addPropertyButton.Caption = "Add Property";
            //addPropertyButton.FaceId = 2;// https://bettersolutions.com/vba/ribbon/face-ids-2003.htm
            addPropertyButton.Click += (CommandBarButton button, ref bool Cancel)
                => AddProperty?.Invoke(button, null);
        }

        public void Show()
        {
            _commandBar.ShowPopup();
        }
    }
}
