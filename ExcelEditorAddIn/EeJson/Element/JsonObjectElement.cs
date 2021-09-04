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
        private JObject JObject;
        public IReadOnlyDictionary<string, IElement> Properties
            => JObject.Properties()
                .Select(x => new { x.Name, Value = x.Value.ToJsonElement(), })
                .ToDictionary(x => x.Name, x => (IElement)x.Value);
        public IEnumerable<string> Keys
            => Properties.Select(x => x.Key).ToList();

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

        public void Add(string key, IElement value)
        {
            var jsonBaseElement = (JsonBaseElement)value;
            JObject.Add(key, jsonBaseElement.Token);
        }

        public void Remove(string key)
        {
            JObject.Remove(key);
        }
    }
}
