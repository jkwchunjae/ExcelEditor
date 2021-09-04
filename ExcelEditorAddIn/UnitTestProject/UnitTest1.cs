using EeJson;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using System;
using System.Linq;

namespace UnitTestProject
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var table = new dynamic[]
            {
                new { Name = "abc", Power = 1234 },
                new { Name = "Item", Price = 1200 },
            };

            var jsonText = JsonConvert.SerializeObject(table);

            var baseElement = new JsonBaseElement(jsonText);
            var tableDocument = new JsonTableElement(baseElement);

            Assert.AreEqual(2, tableDocument.Length);
            Assert.AreEqual(3, tableDocument.Properties.Count());
            Assert.IsTrue(tableDocument.Properties.Any(x => x.PropertyName == "Name"));
            Assert.IsTrue(tableDocument.Properties.Any(x => x.PropertyName == "Power"));
            Assert.IsTrue(tableDocument.Properties.Any(x => x.PropertyName == "Price"));
        }

        [TestMethod]
        public void TestMethod2()
        {
            var table = new dynamic[]
            {
                new { Name = "abc", Power = 1234 },
                new { Name = "Item", Price = 1200 },
            };

            var jsonText = JsonConvert.SerializeObject(table);

            var baseElement = new JsonBaseElement(jsonText);
            var tableDocument = new JsonTableElement(baseElement);
        }
    }
}
