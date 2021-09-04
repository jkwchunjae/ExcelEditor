using EeCommon;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EeJson
{
    public class JsonArrayElement : JsonBaseElement, IArrayElement
    {
        public int Length => JArray.Count;
        public bool Any => Length != 0;
        public bool Empty => Length == 0;

        public IEnumerable<IValueElement> Elements
            => JArray.Select(token => new JsonValueElement(new JsonBaseElement(token)));

        private JArray JArray;

        public JsonArrayElement(JArray array)
            : base(array)
        {
            JArray = array;
        }
        public JsonArrayElement(JsonBaseElement baseElement)
            : base(baseElement.Token)
        {
            JArray = (JArray)baseElement.Token;
        }

        public void Add(IElement element)
        {
            var jsonBaseElement = (JsonBaseElement)element;

            JArray.Add(jsonBaseElement.Token);
        }

        public void AddAt(int index, IElement element)
        {
            var jsonBaseElement = (JsonBaseElement)element;
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
                Add(element);
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
