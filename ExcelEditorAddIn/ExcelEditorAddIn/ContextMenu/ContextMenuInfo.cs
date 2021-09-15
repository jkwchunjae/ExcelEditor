using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        public string BeginColumn { get; private set; }
        public string EndColumn { get; private set; }
        public ColumnMenuInfo(string address, string beginColumn, string endColumn)
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
