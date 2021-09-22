using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace ExcelEditorAddIn
{
    public static class CommandBarsExtensions
    {
        public static CommandBar Find(this CommandBars bars, Func<CommandBar, bool> func)
        {
            foreach (CommandBar bar in bars)
            {
                if (func(bar))
                    return bar;
            }
            return null;
        }
    }
}
