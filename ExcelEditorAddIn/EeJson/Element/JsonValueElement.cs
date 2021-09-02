using EeCommon;
using Newtonsoft.Json.Linq;

namespace EeJson
{
    public class JsonValueElement : JsonBaseElement, IValueElement
    {
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
    }
}
