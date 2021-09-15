using JkwExtensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelEditorAddIn
{
    public static class ContextMenu
    {
        public static ContextMenuInfo Make(string address)
        {
            // $E$7: SingleCell
            // $F$4:$G$10: Cells
            // $C:$C : Column
            // $1:$1 : Row

            var column = @"\$([A-Za-z]+)";
            var row = @"\$(\d+)";
            var cell = $@"{column}{row}"; // \$[A-Za-z]+\$\d+

            var singleCellPattern = $@"^{cell}$";
            if (Regex.IsMatch(address, singleCellPattern))
            {
                return new SingleCellMenuInfo(address);
            }

            var cellsPattern = $@"^{cell}\:{cell}$"; // ^\$[A-Za-z]+\$\d+\:\$[A-Za-z]+\$\d+$
            if (Regex.IsMatch(address, cellsPattern))
            {
                return new CellsMenuInfo(address);
            }

            var rowPattern = $@"^{row}\:{row}$"; // ^\$\d+\:\$\d+$
            if (Regex.IsMatch(address, rowPattern))
            {
                var match = Regex.Match(address, rowPattern);
                var beginRow = match.Groups[1].Value.ToInt();
                var endRow = match.Groups[2].Value.ToInt();
                return new RowMenuInfo(address, beginRow, endRow);
            }

            var columnPattern = $@"^{column}\:{column}$"; // ^\$[A-Za-z]+\:\$[A-Za-z]+$
            if (Regex.IsMatch(address, columnPattern))
            {
                var match = Regex.Match(address, columnPattern);
                var beginColumn = match.Groups[1].Value;
                var endColumn = match.Groups[2].Value;
                return new ColumnMenuInfo(address, beginColumn, endColumn);
            }

            return new UnknownMenuInfo(address);
        }
    }
}
