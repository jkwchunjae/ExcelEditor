﻿using EeCommon;
using JkwExtensions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EeJson
{
    public class JsonBaseElement : IElement
    {
        public JToken Token { get; private set; }
        public ElementType Type => GetElementType();

        public JsonBaseElement(string jsonText)
        {
            Token = JToken.Parse(jsonText);
        }
        public JsonBaseElement(JToken token)
        {
            Token = token;
        }

        private ElementType GetElementType()
        {
            if (Token.Type == JTokenType.Object)
            {
                return ElementType.Object;
            }
            else if (Token.Type == JTokenType.Array)
            {
                var array = (JArray)Token;
                if (array.Empty())
                {
                    return ElementType.Table;
                }
                else
                {
                    var first = array.First();
                    return first.Type == JTokenType.Object
                        ? ElementType.Table : ElementType.Array;
                }
            }
            return ElementType.Value;
        }

        public string GetString()
        {
            var jsonText = JsonConvert.SerializeObject(Token, Formatting.Indented);
            return jsonText;
        }
    }
}
