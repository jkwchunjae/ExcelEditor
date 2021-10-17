using JkwExtensions;
using System.Collections.Generic;
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

        public static void SetColumnSetting(this Metadata metadata, ColumnSetting columnSetting)
        {
            if (metadata.Columns == null)
            {
                metadata.Columns = new List<ColumnSetting> { columnSetting };
            }
            else
            {
                var index = metadata.Columns.FindIndex(x => x.Path == columnSetting.Path);
                if (index == -1)
                {
                    metadata.Columns.Add(columnSetting);
                }
                else
                {
                    metadata.Columns[index] = columnSetting;
                }
            }
        }
    }
}
