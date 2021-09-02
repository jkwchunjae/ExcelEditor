using EeCommon;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace EeJson
{
    public class JsonObjectDocument : JsonBaseDocument, IObjectDocument
    {
        public List<string> Keys => throw new NotImplementedException();

        public JsonObjectDocument(string jsonText)
            : base(jsonText)
        {
        }
        public JsonObjectDocument(JsonBaseDocument baseDocument)
            : base(baseDocument.Doc)
        {
        }
    }
}
