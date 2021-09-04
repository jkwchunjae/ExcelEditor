using EeCommon;
using Newtonsoft.Json.Linq;

namespace EeJson
{
    public class JsonValueElement : JsonBaseElement, IValueElement
    {
        public ValueType ValueType => JValue.Type.ToValueType();

        private JValue JValue;
        public JsonValueElement(JValue value)
            : base(value)
        {
            JValue = value;
        }
        public JsonValueElement(JsonBaseElement baseElement)
            : base(baseElement.Token)
        {
            JValue = (JValue)baseElement.Token;
        }

        public void UpdateValue(object value, object value2)
        {
            if (value == null)
            {
                JValue.Value = null;
            }
            else
            {
                JValue jValue = JsonExtensions.CreateJValue(value, value2);
                JValue.Value = jValue.Value;
            }
        }
    }
}
