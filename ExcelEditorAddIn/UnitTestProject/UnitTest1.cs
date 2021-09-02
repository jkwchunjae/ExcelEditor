using EeJson;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using System;

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
            Assert.AreEqual(3, tableDocument.Keys.Count);
            Assert.IsTrue(tableDocument.Keys.Contains("Name"));
            Assert.IsTrue(tableDocument.Keys.Contains("Power"));
            Assert.IsTrue(tableDocument.Keys.Contains("Price"));
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
