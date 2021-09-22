using JkwExtensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelEditorAddIn
{
    public enum ContextMenuType
    {
        Unknown,
        SingleCell,
        Cells,
        Row,
        Column,
    }

    public abstract class ContextMenuInfo
    {
        public ContextMenuType Type { get; protected set; }
        public string Address { get; protected set; }

        public ContextMenuInfo(ContextMenuType type, string address)
        {
            Type = type;
            Address = address;
        }

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
                var beginColumn = Utils.ColumnNameToNumber(match.Groups[1].Value);
                var endColumn = Utils.ColumnNameToNumber(match.Groups[2].Value);
                return new ColumnMenuInfo(address, beginColumn, endColumn);
            }

            return new UnknownMenuInfo(address);
        }
    }

    public class SingleCellMenuInfo : ContextMenuInfo
    {
        public SingleCellMenuInfo(string address)
            : base(ContextMenuType.SingleCell, address)
        {
        }
    }

    public class CellsMenuInfo : ContextMenuInfo
    {
        public CellsMenuInfo(string address)
            : base(ContextMenuType.Cells, address)
        {
        }
    }

    public class ColumnMenuInfo : ContextMenuInfo
    {
        public int BeginColumn { get; private set; }
        public int EndColumn { get; private set; }
        public ColumnMenuInfo(string address, int beginColumn, int endColumn)
            : base(ContextMenuType.Column, address)
        {
            BeginColumn = beginColumn;
            EndColumn = endColumn;
        }
    }

    public class RowMenuInfo : ContextMenuInfo
    {
        public int BeginRow { get; private set; }
        public int EndRow { get; private set; }
        public RowMenuInfo(string address, int beginRow, int endRow)
            : base(ContextMenuType.Row, address)
        {
            BeginRow = beginRow;
            EndRow = endRow;
        }
    }

    public class UnknownMenuInfo : ContextMenuInfo
    {
        public UnknownMenuInfo(string address)
            : base(ContextMenuType.Unknown, address)
        {
        }
    }
}
