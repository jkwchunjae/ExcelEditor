using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEditorAddIn
{
    public static class ContextMenuFactory
    {
        public static ContextMenu_Column CreateColumnMenu(string id)
        {
            var menuName = nameof(ContextMenu_Column) + id;

            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            var commandBar = commandBars.Find(bar => bar.Name == menuName);

            if (commandBar != null)
            {
                return new ContextMenu_Column(commandBar);
            }

            var newBar = commandBars.Add(Name: menuName,
                                      Position: MsoBarPosition.msoBarPopup,
                                      Temporary: true);
            var menu = new ContextMenu_Column(newBar);

            return menu;
        }
    }
}
