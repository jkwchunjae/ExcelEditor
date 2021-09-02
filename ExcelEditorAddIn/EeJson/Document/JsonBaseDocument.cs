using EeCommon;
using System.Linq;
using System.Text.Json;

namespace EeJson
{
    public class JsonBaseDocument : IDocument
    {
        public JsonDocument Doc { get; private set; }

        public DocumentType Type => GetDocumentType();

        public JsonBaseDocument(string jsonText)
        {
            Doc = JsonDocument.Parse(jsonText, new JsonDocumentOptions
            {
                AllowTrailingCommas = true,
                CommentHandling = JsonCommentHandling.Skip,
            });
        }
        public JsonBaseDocument(JsonDocument doc)
        {
            Doc = doc;
        }

        private DocumentType GetDocumentType()
        {
            if (Doc.RootElement.ValueKind == JsonValueKind.Object)
            {
                return DocumentType.Object;
            }
            else if (Doc.RootElement.ValueKind == JsonValueKind.Array)
            {
                var length = Doc.RootElement.GetArrayLength();
                if (length == 0)
                {
                    return DocumentType.Table;
                }
                else
                {
                    var firstElement = Doc.RootElement.EnumerateArray().First();
                    if (firstElement.ValueKind == JsonValueKind.Object)
                    {
                        return DocumentType.Table;
                    }
                    else
                    {
                        return DocumentType.Array;
                    }
                }
            }
            return DocumentType.Value;
        }

        public string GetString()
        {
            var jsonText = JsonSerializer.Serialize(Doc, new JsonSerializerOptions
            {
                WriteIndented = true,
                NumberHandling = System.Text.Json.Serialization.JsonNumberHandling.AllowNamedFloatingPointLiterals,
            });
            return jsonText;
        }
    }
}
