using EeCommon;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace EeJson
{
    public class JsonTableDocument : JsonBaseDocument, ITableDocument
    {
        public int Length => Doc?.RootElement.GetArrayLength() ?? 0;
        public bool Any => Length != 0;
        public bool Empty => Length == 0;
        public List<string> Keys { get; private set; }
        public object[,] Values { get; private set; }

        public JsonTableDocument(string jsonText)
            : base(jsonText)
        {
            Keys = MakeKeys(Doc);
            Values = MakeValues(Doc);
        }
        public JsonTableDocument(JsonBaseDocument baseDocument)
            : base(baseDocument.Doc)
        {
            Keys = MakeKeys(Doc);
            Values = MakeValues(Doc);
        }

        private List<string> MakeKeys(JsonDocument doc)
        {
            var keys = doc.RootElement.EnumerateArray()
                .SelectMany(x => x.EnumerateObject().Select(e => e.Name))
                .Distinct()
                .ToList();

            return keys;
        }

        private object[,] MakeValues(JsonDocument doc)
        {
            var values = new object[Length, Keys.Count];
            for (var row = 0; row < Length; row++)
            {
                var current = doc.RootElement[row];
                var obj = current.EnumerateObject()
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
