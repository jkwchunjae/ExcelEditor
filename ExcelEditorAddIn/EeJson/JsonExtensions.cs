using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace EeJson
{
    public static class JsonExtensions
    {
        private static JsonValueKind[] _primitiveTypes = new[]
        {
            JsonValueKind.Null,
            JsonValueKind.True,
            JsonValueKind.String,
            JsonValueKind.False,
            JsonValueKind.Number,
        };

        public static bool IsPrimitiveType(this JsonValueKind jsonValueKind)
        {
            return _primitiveTypes.Contains(jsonValueKind);
        }

        public static object ToExcelValue(this JsonElement jsonElement)
        {
            switch (jsonElement.ValueKind)
            {
                case JsonValueKind.Null:
                    return null;
                case JsonValueKind.True:
                    return true;
                case JsonValueKind.False:
                    return false;
                case JsonValueKind.Number:
                    return jsonElement.GetDouble();
                case JsonValueKind.String:
                    return jsonElement.GetString();
                case JsonValueKind.Array:
                    return "[array]";
                case JsonValueKind.Object:
                    return "{object}";
                default:
                    return null;
            }
        }
    }
}
