using Newtonsoft.Json;
using System.Collections.Generic;

namespace ExcelEditorAddIn
{
    public class ColumnSetting
    {
        public class OrderWidth
        {
            [JsonProperty("name")]
            public string Name { get; set; }
            [JsonProperty("width")]
            public double? Width { get; set; }
            [JsonIgnore]
            public int Order { get; set; }
        }

        [JsonProperty("path")]
        public string Path { get; set; }
        [JsonProperty("orderwidth")]
        public List<OrderWidth> Info { get; set; }

        public OrderWidth GetOrderData(string columnName)
        {
            if (Info == null)
                return null;

            for (var order = 0; order < Info.Count; order++)
            {
                if (Info[order].Name == columnName)
                {
                    Info[order].Order = order;
                    return Info[order];
                }
            }

            return null;
        }
    }
}
