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
    }
}
