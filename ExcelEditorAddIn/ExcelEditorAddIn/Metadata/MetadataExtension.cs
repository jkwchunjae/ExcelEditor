using JkwExtensions;
using System.Linq;

namespace ExcelEditorAddIn
{
    public static class MetadataExtension
    {
        private static bool MatchPath(string patternPath, string realPath)
        {
            return true;
        }

        public static ColumnSetting GetColumnSetting(this Metadata metadata, string path)
        {
            if (metadata?.Columns?.Empty() ?? true)
            {
                return null;
            }

            var columnSetting = metadata.Columns.FirstOrDefault(x => MatchPath(path, x.Path));

            return columnSetting;
        }
    }
}
