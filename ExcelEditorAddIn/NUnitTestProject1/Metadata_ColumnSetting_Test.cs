﻿using EeCommon;
using ExcelEditorAddIn;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NUnitTestProject1
{
    public class Metadata_ColumnSetting_Test
    {
        private readonly List<(string PropertyName, ElementType ElementType)> _properties = new List<(string PropertyName, ElementType ElementType)>
        {
            ("Column1", ElementType.Value),
            ("Column2", ElementType.Value),
            ("Column3", ElementType.Value),
            ("Column4", ElementType.Value),
        };

        [Test]
        public void PropertyInfo_Order_Test1()
        {
            var columnSetting = new ColumnSetting()
            {
                Info = new List<ColumnSetting.OrderWidth>()
                {
                    new ColumnSetting.OrderWidth { Name = "Column1", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column2", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column3", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column4", Width = 100 },
                },
            };

            var ordered = _properties
                .OrderBy(property => property, new ColumnComparer(columnSetting))
                .Select(x => x.PropertyName)
                .ToList();

            Assert.AreEqual(_properties.Count, ordered.Count);
            Assert.AreEqual("Column1", ordered[0]);
            Assert.AreEqual("Column2", ordered[1]);
            Assert.AreEqual("Column3", ordered[2]);
            Assert.AreEqual("Column4", ordered[3]);
        }

        [Test]
        public void PropertyInfo_Order_Test2()
        {
            var columnSetting = new ColumnSetting()
            {
                Info = new List<ColumnSetting.OrderWidth>()
                {
                    new ColumnSetting.OrderWidth { Name = "Column2", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column1", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column4", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column3", Width = 100 },
                },
            };

            var ordered = _properties
                .OrderBy(property => property, new ColumnComparer(columnSetting))
                .Select(x => x.PropertyName)
                .ToList();

            Assert.AreEqual(_properties.Count, ordered.Count);
            Assert.AreEqual("Column2", ordered[0]);
            Assert.AreEqual("Column1", ordered[1]);
            Assert.AreEqual("Column4", ordered[2]);
            Assert.AreEqual("Column3", ordered[3]);
        }

        [Test]
        public void PropertyInfo_Order_Test3()
        {
            var columnSetting = new ColumnSetting()
            {
                Info = new List<ColumnSetting.OrderWidth>()
                {
                    new ColumnSetting.OrderWidth { Name = "Column1", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column2", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column4", Width = 100 },
                },
            };

            var ordered = _properties
                .OrderBy(property => property, new ColumnComparer(columnSetting))
                .Select(x => x.PropertyName)
                .ToList();

            Assert.AreEqual(_properties.Count, ordered.Count);
            Assert.AreEqual("Column1", ordered[0]);
            Assert.AreEqual("Column2", ordered[1]);
            Assert.AreEqual("Column4", ordered[2]);
            Assert.AreEqual("Column3", ordered[3]);
        }

        [Test]
        public void PropertyInfo_Order_Test4()
        {
            var columnSetting = new ColumnSetting()
            {
                Info = new List<ColumnSetting.OrderWidth>()
                {
                    new ColumnSetting.OrderWidth { Name = "Column1", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column2", Width = 100 },
                },
            };

            var ordered = _properties
                .OrderBy(property => property, new ColumnComparer(columnSetting))
                .Select(x => x.PropertyName)
                .ToList();

            Assert.AreEqual(_properties.Count, ordered.Count);
            Assert.AreEqual("Column1", ordered[0]);
            Assert.AreEqual("Column2", ordered[1]);
            Assert.AreEqual("Column3", ordered[2]);
            Assert.AreEqual("Column4", ordered[3]);
        }

        [Test]
        public void PropertyInfo_Order_Test5()
        {
            var columnSetting = new ColumnSetting()
            {
                Info = new List<ColumnSetting.OrderWidth>()
                {
                    new ColumnSetting.OrderWidth { Name = "Column2", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column3", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column4", Width = 100 },
                },
            };

            var ordered = _properties
                .OrderBy(property => property, new ColumnComparer(columnSetting))
                .Select(x => x.PropertyName)
                .ToList();

            Assert.AreEqual(_properties.Count, ordered.Count);
            Assert.AreEqual("Column2", ordered[0]);
            Assert.AreEqual("Column3", ordered[1]);
            Assert.AreEqual("Column4", ordered[2]);
            Assert.AreEqual("Column1", ordered[3]);
        }

        [Test]
        public void PropertyInfo_Order_Test6()
        {
            var columnSetting = new ColumnSetting()
            {
                Info = new List<ColumnSetting.OrderWidth>()
                {
                    new ColumnSetting.OrderWidth { Name = "Column2", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column4", Width = 100 },
                    new ColumnSetting.OrderWidth { Name = "Column3", Width = 100 },
                },
            };

            var ordered = _properties
                .OrderBy(property => property, new ColumnComparer(columnSetting))
                .Select(x => x.PropertyName)
                .ToList();

            Assert.AreEqual(_properties.Count, ordered.Count);
            Assert.AreEqual("Column2", ordered[0]);
            Assert.AreEqual("Column4", ordered[1]);
            Assert.AreEqual("Column3", ordered[2]);
            Assert.AreEqual("Column1", ordered[3]);
        }

        [Test]
        public void PropertyInfo_Order_Test7()
        {
            var columnSetting = new ColumnSetting()
            {
                Info = null,
            };

            var ordered = _properties
                .OrderBy(property => property, new ColumnComparer(columnSetting))
                .Select(x => x.PropertyName)
                .ToList();

            Assert.AreEqual(_properties.Count, ordered.Count);
            Assert.AreEqual("Column1", ordered[0]);
            Assert.AreEqual("Column2", ordered[1]);
            Assert.AreEqual("Column3", ordered[2]);
            Assert.AreEqual("Column4", ordered[3]);
        }

        [Test]
        public void PropertyInfo_Order_Test8()
        {
            ColumnSetting columnSetting = null;

            var ordered = _properties
                .OrderBy(property => property, new ColumnComparer(columnSetting))
                .Select(x => x.PropertyName)
                .ToList();

            Assert.AreEqual(_properties.Count, ordered.Count);
            Assert.AreEqual("Column1", ordered[0]);
            Assert.AreEqual("Column2", ordered[1]);
            Assert.AreEqual("Column3", ordered[2]);
            Assert.AreEqual("Column4", ordered[3]);
        }
    }
}
