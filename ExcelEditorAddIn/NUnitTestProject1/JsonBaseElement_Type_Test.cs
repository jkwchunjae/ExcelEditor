using EeCommon;
using EeJson;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace NUnitTestProject1
{
    public class JsonBaseElement_Type_Test
    {
        [Test]
        public void Check_ElementType_when_number()
        {
            JToken jtoken = new JValue(123);
            var baseElement = new JsonBaseElement(jtoken);

            Assert.AreEqual(ElementType.Value, baseElement.Type);
        }

        [Test]
        public void Check_ElementType_when_string()
        {
            JToken jtoken = new JValue("text");
            var baseElement = new JsonBaseElement(jtoken);

            Assert.AreEqual(ElementType.Value, baseElement.Type);
        }

        [Test]
        public void Check_ElementType_when_boolean()
        {
            JToken jtoken = new JValue(true);
            var baseElement = new JsonBaseElement(jtoken);

            Assert.AreEqual(ElementType.Value, baseElement.Type);
        }

        [Test]
        public void Check_ElementType_when_array()
        {
            var arr = new[] { 1, 2, 3 };
            var arrJsonText = JsonConvert.SerializeObject(arr);
            var baseElement = new JsonBaseElement(arrJsonText);

            Assert.AreEqual(ElementType.Array, baseElement.Type);
        }

        [Test]
        public void Check_ElementType_when_empty_array()
        {
            var arr = new int[] { };
            var arrJsonText = JsonConvert.SerializeObject(arr);
            var baseElement = new JsonBaseElement(arrJsonText);

            Assert.AreEqual(ElementType.Table, baseElement.Type);
        }

        [Test]
        public void Check_ElementType_when_object()
        {
            var obj = new
            {
                Name = "jkw",
                Items = new[] { "A", "B", "C" },
            };
            var objJsonText = JsonConvert.SerializeObject(obj);
            var baseElement = new JsonBaseElement(objJsonText);

            Assert.AreEqual(ElementType.Object, baseElement.Type);
        }

        [Test]
        public void Check_ElementType_when_table()
        {
            var table = new[]
            {
                new
                {
                    Name = "jkw",
                    Items = new[] { "A", "B", "C" },
                },
                new
                {
                    Name = "abc",
                    Items = new[] { "A", "B", "C", "D" },
                },
            };

            var tableJsonText = JsonConvert.SerializeObject(table);
            var baseElement = new JsonBaseElement(tableJsonText);

            Assert.AreEqual(ElementType.Table, baseElement.Type);
        }
    }
}
