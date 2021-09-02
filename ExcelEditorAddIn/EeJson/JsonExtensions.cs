using EeCommon;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EeJson
{
    public static class JsonExtensions
    {
        private static JTokenType[] _primitiveTypes = new[]
        {
            JTokenType.String,
            JTokenType.Integer,
            JTokenType.Boolean,
            JTokenType.Date,
            JTokenType.Float,
            JTokenType.Guid,
            JTokenType.Uri,
            JTokenType.TimeSpan,
            JTokenType.Null,
        };

        public static bool IsPrimitiveType(this JTokenType tokenType)
        {
            return _primitiveTypes.Contains(tokenType);
        }

        public static object ToExcelValue(this JToken token)
        {
            switch (token.Type)
            {
                case JTokenType.Null:
                    return null;
                case JTokenType.Boolean:
                    return token.Value<bool>();
                case JTokenType.Integer:
                    return token.Value<int>();
                case JTokenType.Float:
                    return token.Value<double>();
                case JTokenType.String:
                    return token.Value<string>();
                case JTokenType.Date:
                    return token.Value<DateTime>();
                case JTokenType.Guid:
                    return token.Value<string>();
                case JTokenType.Array:
                    return "[array]";
                case JTokenType.Object:
                    return "{object}";
                default:
                    return token.Value<string>();
            }
        }

        public static JsonBaseElement ToJsonElement(this JToken token)
        {
            var baseElement = new JsonBaseElement(token);

            switch (baseElement.Type)
            {
                case ElementType.Table:
                    return new JsonTableElement(baseElement);
                case ElementType.Array:
                    return new JsonArrayElement(baseElement);
                case ElementType.Object:
                    return new JsonObjectElement(baseElement);
                default:
                    return new JsonValueElement(baseElement);
            }
        }
    }
}
