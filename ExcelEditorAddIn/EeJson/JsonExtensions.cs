using EeCommon;
using JkwExtensions;
using Newtonsoft.Json.Linq;
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

        public static ValueType ToValueType(this JTokenType tokenType)
        {
            switch (tokenType)
            {
                case JTokenType.Null:
                    return ValueType.Null;
                case JTokenType.Boolean:
                    return ValueType.Boolean;
                case JTokenType.Integer:
                    return ValueType.Integer;
                case JTokenType.Float:
                    return ValueType.Float;
                case JTokenType.String:
                    return ValueType.String;
                case JTokenType.Date:
                    return ValueType.DateTime;
                case JTokenType.TimeSpan:
                    return ValueType.DateTime;
                default:
                    return ValueType.String;
            }
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
                    return token.Value<System.DateTime>();
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

        public static JValue CreateJValue(object value, object value2)
        {
            if (value == null)
            {
                return new JValue((string)null);
            }
            else
            {
                string valueText = value.ToString();
                if (long.TryParse(valueText, out var longValue))
                {
                    return new JValue(longValue);
                }
                else if (double.TryParse(valueText, out var doubleValue))
                {
                    return new JValue(doubleValue);
                }
                else if (bool.TryParse(valueText, out var boolValue))
                {
                    return new JValue(boolValue);
                }
                else
                {
                    return new JValue(valueText);
                }
            }
        }

        public static JValue CreateJValue(object value, object value2, ValueType valueType)
        {
            if (value == null)
            {
                return new JValue((string)null);
            }
            else
            {
                string valueText = value.ToString();
                switch (valueType)
                {
                    case ValueType.Integer:
                    case ValueType.Float:
                        {
                            if (long.TryParse(valueText, out var longValue))
                            {
                                return new JValue(longValue);
                            }
                            else if (double.TryParse(valueText, out var doubleValue))
                            {
                                return new JValue(doubleValue);
                            }
                            throw new RequireNumberException(value, value2);
                        }
                    case ValueType.Boolean:
                        {
                            if (bool.TryParse(valueText, out var boolValue))
                            {
                                return new JValue(boolValue);
                            }
                            throw new RequireBooleanException(value, value2);
                        }
                }
                return new JValue(valueText);
            }
        }
    }
}
