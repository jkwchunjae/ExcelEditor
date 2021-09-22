using ExcelEditorAddIn;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace NUnitTestProject1
{
    public class Utils_GetContextMenuType_Test
    {
        [Test]
        public void GetContextMenuType_SingleCell_1()
        {
            var address = "$E$1";

            var info = ContextMenuInfo.Make(address) as SingleCellMenuInfo;

            Assert.AreEqual(ContextMenuType.SingleCell, info.Type);
        }

        [Test]
        public void GetContextMenuType_SingleCell_2()
        {
            var address = "$AB$1123";

            var info = ContextMenuInfo.Make(address) as SingleCellMenuInfo;

            Assert.AreEqual(ContextMenuType.SingleCell, info.Type);
        }

        [Test]
        public void GetContextMenuType_Cells_1()
        {
            var address = "$E$1:$F$10";

            var info = ContextMenuInfo.Make(address) as CellsMenuInfo;

            Assert.AreEqual(ContextMenuType.Cells, info.Type);
        }

        [Test]
        public void GetContextMenuType_Cells_2()
        {
            var address = "$A$11:$BB$100";

            var info = ContextMenuInfo.Make(address) as CellsMenuInfo;

            Assert.AreEqual(ContextMenuType.Cells, info.Type);
        }

        [Test]
        public void GetContextMenuType_Columns_1()
        {
            var address = "$E:$F";

            var info = ContextMenuInfo.Make(address) as ColumnMenuInfo;

            Assert.AreEqual(ContextMenuType.Column, info.Type);
            Assert.AreEqual("E", info.BeginColumn);
            Assert.AreEqual("F", info.EndColumn);
        }

        [Test]
        public void GetContextMenuType_Columns_2()
        {
            var address = "$E:$EE";

            var info = ContextMenuInfo.Make(address) as ColumnMenuInfo;

            Assert.AreEqual(ContextMenuType.Column, info.Type);
            Assert.AreEqual("E", info.BeginColumn);
            Assert.AreEqual("EE", info.EndColumn);
        }

        [Test]
        public void GetContextMenuType_Rows_1()
        {
            var address = "$3:$10";

            var info = ContextMenuInfo.Make(address) as RowMenuInfo;

            Assert.AreEqual(ContextMenuType.Row, info.Type);
            Assert.AreEqual(3, info.BeginRow);
            Assert.AreEqual(10, info.EndRow);
        }

        [Test]
        public void GetContextMenuType_Unknown_1()
        {
            var address = "$:$10";

            var info = ContextMenuInfo.Make(address) as UnknownMenuInfo;

            Assert.AreEqual(ContextMenuType.Unknown, info.Type);
        }

        [Test]
        public void GetContextMenuType_Unknown_2()
        {
            var address = "$E$";

            var info = ContextMenuInfo.Make(address);

            Assert.AreEqual(ContextMenuType.Unknown, info.Type);
        }

        [Test]
        public void GetContextMenuType_Unknown_3()
        {
            var address = "$E$1:";

            var info = ContextMenuInfo.Make(address);

            Assert.AreEqual(ContextMenuType.Unknown, info.Type);
        }

        [Test]
        public void GetContextMenuType_Unknown_4()
        {
            var address = "$1:$";

            var info = ContextMenuInfo.Make(address);

            Assert.AreEqual(ContextMenuType.Unknown, info.Type);
        }


    }
}
