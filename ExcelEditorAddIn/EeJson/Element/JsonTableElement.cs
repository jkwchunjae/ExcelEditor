using EeCommon;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EeJson
{
    public class JsonTableElement : JsonBaseElement, ITableElement
    {
        public int Length => JArray.Count;
        public bool Any => Length != 0;
        public bool Empty => Length == 0;
        public List<string> Keys { get; private set; }
        public object[,] Values { get; private set; }

        private JArray JArray;
        public JsonTableElement(JArray array)
            : base(array)
        {
            JArray = array;
            Keys = MakeKeys();
            Values = MakeValues();
        }
        public JsonTableElement(JsonBaseElement baseElement)
            : base(baseElement.Token)
        {
            JArray = (JArray)baseElement.Token;
            Keys = MakeKeys();
            Values = MakeValues();
        }

        private List<string> MakeKeys()
        {
            var keys = JArray
                .SelectMany(x => ((JObject)x).Properties().Select(e => e.Name))
                .Distinct()
                .ToList();
            return keys;
        }
        private object[,] MakeValues()
        {
            var values = new object[Length, Keys.Count];
            for (var row = 0; row < Length; row++)
            {
                var current = (JObject)JArray[row];
                var obj = current.Properties()
                    .ToDictionary(x => x.Name, x => x.Value);
                for (var column = 0; column < Keys.Count; column++)
                {
                    var key = Keys[column];
                    if (obj.TryGetValue(key, out var element))
                    {
                        values[row, column] = element.ToExcelValue();
                    }
                    else
                    {
                        values[row, column] = null;
                    }
                }
            }
            return values;
        }
    }
}
