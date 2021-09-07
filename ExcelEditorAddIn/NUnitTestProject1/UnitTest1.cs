using EeJson;
using Newtonsoft.Json;
using NUnit.Framework;
using System.Linq;

namespace NUnitTestProject1
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void Test1()
        {
            Assert.Pass();
        }

        [Test]
        public void Test2()
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
    }
}