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
        private static List<(string MenuName, IContextMenu ContextMenu)> _contextCache = new List<(string MenuName, IContextMenu)>();

        public static ContextMenu_Column CreateColumnMenu(string id)
        {
            var menuName = nameof(ContextMenu_Column) + id;

            if (_contextCache.Any(x => x.MenuName == menuName))
            {
                return (ContextMenu_Column)_contextCache
                    .First(x => x.MenuName == menuName)
                    .ContextMenu;
            }

            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            var bar = commandBars.Add(
                Name: menuName,
                Position: MsoBarPosition.msoBarPopup,
                Temporary: true);
            var menu = new ContextMenu_Column(bar);

            _contextCache.Add((menuName, menu));

            return menu;
        }
    }
}
