using Newtonsoft.Json;
using System.Collections.Generic;
using System.Linq;

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
        public List<OrderWidth> OrderWidthList { get; set; }

        public OrderWidth GetOrderData(string columnName)
        {
            if (OrderWidthList == null)
                return null;

            for (var order = 0; order < OrderWidthList.Count; order++)
            {
                if (OrderWidthList[order].Name == columnName)
                {
                    OrderWidthList[order].Order = order;
                    return OrderWidthList[order];
                }
            }

            return null;
        }

        public override bool Equals(object obj)
        {
            if (obj is ColumnSetting setting)
            {
                if (Path != setting.Path)
                    return false;

                var owList1 = OrderWidthList ?? new List<OrderWidth>();
                var owList2 = setting.OrderWidthList ?? new List<OrderWidth>();

                if (owList1.Count != owList2.Count)
                    return false;

                for (var i = 0; i < owList1.Count; i++)
                {
                    var ow1 = owList1[i];
                    var ow2 = owList2[i];
                    if (ow1.Name != ow2.Name)
                        return false;
                    if (ow1.Width.HasValue != ow2.Width.HasValue)
                        return false;
                    if (ow1.Width.HasValue && ow2.Width.HasValue)
                        if (ow1.Width - ow2.Width > double.Epsilon) // ow1.Width != ow2.Width
                            return false;
                }

                return true;
            }
            else
            {
                return false;
            }
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public static bool operator ==(ColumnSetting obj1, ColumnSetting obj2)
        {
            if (ReferenceEquals(obj1, obj2))
            {
                return true;
            }
            if (ReferenceEquals(obj1, null))
            {
                return false;
            }
            if (ReferenceEquals(obj2, null))
            {
                return false;
            }

            return obj1.Equals(obj2);
        }
        public static bool operator !=(ColumnSetting obj1, ColumnSetting obj2)
        {
            return !(obj1 == obj2);

        }
    }
}
