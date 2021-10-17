using JkwExtensions;
using System.Collections.Generic;
using System.Linq;

namespace ExcelEditorAddIn
{
    public static class MetadataExtension
    {
        private static bool MatchPath(string normalizedPath, string realPath)
        {
            return normalizedPath == realPath;
        }

        public static ColumnSetting GetColumnSetting(this Metadata metadata, string path)
        {
            if (metadata.Columns?.Empty() ?? true)
            {
                return null;
            }

            var columnSetting = metadata.Columns.FirstOrDefault(x => MatchPath(x.Path, path));

            return columnSetting;
        }

        public static void SetColumnSetting(this Metadata metadata, ColumnSetting columnSetting)
        {
            metadata.Dirty = true;
            if (metadata.Columns == null)
            {
                metadata.Columns = new List<ColumnSetting> { columnSetting };
            }
            else
            {
                // 순서를 바꾸지 않게 하기 위해서 이렇게 처리함.
                var index = metadata.Columns.FindIndex(x => MatchPath(x.Path, columnSetting.Path));
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
