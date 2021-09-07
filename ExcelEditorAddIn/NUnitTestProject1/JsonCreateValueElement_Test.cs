using EeCommon;
using EeJson;
using NUnit.Framework;
using System.Collections.Generic;
using System.Text;

namespace NUnitTestProject1
{
    public class JsonCreateValueElement_Test
    {
        IElement _e = new JsonBaseElement("1");

        [Test]
        public void CreateValueElement_int()
        {
            var v = _e.CreateValueElement(123, 123);

            Assert.AreEqual(ValueType.Integer, v.ValueType);
        }

        [Test]
        public void CreateValueElement_int2()
        {
            var v = _e.CreateValueElement(123, 123, ValueType.String);

            Assert.AreEqual(ValueType.String, v.ValueType);
        }

        [Test]
        public void CreateValueElement_throw_RequireNumberException()
        {
            try
            {
                var v = _e.CreateValueElement("text", "text", ValueType.Integer);

                Assert.Fail($"{nameof(RequireNumberException)} 예외가 발생하지 않았습니다.");
            }
            catch (RequireNumberException)
            {
                Assert.Pass($"{nameof(RequireNumberException)} 예외가 정상적으로 발생했습니다.");
            }
        }

        [Test]
        public void CreateValueElement_float()
        {
            var v = _e.CreateValueElement(123.2, 123.2);

            Assert.AreEqual(ValueType.Float, v.ValueType);
        }

        [Test]
        public void CreateValueElement_float2()
        {
            var v = _e.CreateValueElement(123.0, 123.0);

            Assert.AreEqual(ValueType.Integer, v.ValueType);
        }
    }
}
