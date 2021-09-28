using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEditorAddIn
{
    public static class PathOf
    {
        private static string ProjectName { get; } = "ExcelEditor";

        public static string LocalRootDirectory
            => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ProjectName);

        public static string TemporaryFilePath(string fileName)
            => Path.Combine(LocalRootDirectory, $"{fileName}.xlsx");

        public static string MetadataFilePath(string filePath)
        {
            var ext = Path.GetExtension(filePath);
            var filePathWithoutExt = filePath.Substring(0, filePath.Length - ext.Length);

            if (ext == ".json")
            {
                var metadataPath = $@"{filePathWithoutExt}.jmeta";
                return metadataPath;
            }

            return $@"{filePathWithoutExt}.eemeta";
        }
    }
}
