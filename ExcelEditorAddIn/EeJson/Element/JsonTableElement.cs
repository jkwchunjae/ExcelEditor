using EeCommon;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JkwExtensions;

namespace EeJson
{
    public class JsonTableElement : JsonBaseElement, ITableElement
    {
        public int Length => JArray.Count;
        public bool Any => Length != 0;
        public bool Empty => Length == 0;
        public IEnumerable<IObjectElement> Elements
            => JArray.Select(x => new JsonObjectElement((JObject)x));
        public IEnumerable<(string PropertyName, ElementType ElementType)> Properties
            => Elements.SelectMany(x => x.Properties
                .Select(property => (property.Key, property.Value.Type)))
                .Distinct();

        private JArray JArray;
        public JsonTableElement(JArray array)
            : base(array)
        {
            JArray = array;
        }
        public JsonTableElement(JsonBaseElement baseElement)
            : base(baseElement.Token)
        {
            JArray = (JArray)baseElement.Token;
        }

        public void Add(IObjectElement objectElement)
        {
            var jsonObjectElement = (JsonObjectElement)objectElement;

            JArray.Add(jsonObjectElement.Token);
        }

        public void AddAt(int index, IObjectElement objectElement)
        {
            var jsonBaseElement = (JsonObjectElement)objectElement;
            if (index < 0)
            {
                throw new IndexOutOfRangeException();
            }
            else if (index == 0)
            {
                JArray.AddFirst(jsonBaseElement.Token);
            }
            else if (index >= JArray.Count)
            {
                Add(objectElement);
            }
            else
            {
                var item = JArray.ElementAt(index);
                item.AddBeforeSelf(jsonBaseElement.Token);
            }
        }

        public void RemoveAt(int index)
        {
            JArray.RemoveAt(index);
        }
    }
}
