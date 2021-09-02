using EeCommon;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EeJson
{
    public class JsonObjectElement : JsonBaseElement, IObjectElement
    {
        public List<string> Keys => throw new NotImplementedException();

        private JObject JObject;

        public JsonObjectElement(JObject obj)
            : base(obj)
        {
            JObject = obj;
        }
        public JsonObjectElement(JsonBaseElement baseElement)
            : base(baseElement.Token)
        {
            JObject = (JObject)baseElement.Token;
        }
    }
}
