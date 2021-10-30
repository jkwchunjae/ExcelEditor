using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEditorAddIn
{
    public class RecentItem
    {
        public string FilePath { get; set; }
        public DateTime OpenedTime { get; set; }
    }

    public class Recents
    {
        public static event EventHandler ItemUpdated;
        public static List<RecentItem> Items
        {
            get
            {
                if (File.Exists(PathOf.RecentsPath()))
                {
                    try
                    {
                        var jsonText = File.ReadAllText(PathOf.RecentsPath());
                        var recents = JsonConvert.DeserializeObject<List<RecentItem>>(jsonText);
                        return recents
                            .OrderByDescending(x => x.OpenedTime)
                            .Take(20)
                            .ToList();
                    }
                    catch
                    {
                    }
                }

                return new List<RecentItem>();
            }

            private set
            {
                var jsonText = JsonConvert.SerializeObject(value, Formatting.Indented);
                File.WriteAllText(PathOf.RecentsPath(), jsonText, Encoding.UTF8);
            }
        }

        public static void Update(string filePath)
        {
            var items = Items;
            if (items.Any(x => x.FilePath == filePath))
            {
                var item = items.Find(x => x.FilePath == filePath);
                item.OpenedTime = DateTime.Now;
            }
            else
            {
                items.Add(new RecentItem
                {
                    FilePath = filePath,
                    OpenedTime = DateTime.Now,
                });
            }

            Items = items;

            ItemUpdated?.Invoke(null, null);
        }
    }
}
