using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEditorAddIn
{
    public class FavoriteItem
    {
        public string FilePath { get; set; }
        public DateTime AddTime { get; set; }
        public string Nickname { get; set; }
    }

    public class Favorites
    {
        public static event EventHandler<List<FavoriteItem>> ItemUpdated;

        public static List<FavoriteItem> Items
        {
            get
            {
                if (File.Exists(PathOf.FavoritesPath()))
                {
                    try
                    {
                        var jsonText = File.ReadAllText(PathOf.FavoritesPath());
                        var recents = JsonConvert.DeserializeObject<List<FavoriteItem>>(jsonText);
                        return recents
                            .OrderByDescending(x => x.AddTime)
                            .Take(20)
                            .ToList();
                    }
                    catch
                    {
                    }
                }

                return new List<FavoriteItem>();
            }

            private set
            {
                var jsonText = JsonConvert.SerializeObject(value, Formatting.Indented);
                File.WriteAllText(PathOf.FavoritesPath(), jsonText, Encoding.UTF8);
            }
        }

        public static void Update(string filePath)
        {
            var items = Items;
            if (items.Any(x => x.FilePath == filePath))
            {
                var item = items.Find(x => x.FilePath == filePath);
                item.AddTime = DateTime.Now;
            }
            else
            {
                items.Add(new FavoriteItem
                {
                    FilePath = filePath,
                    AddTime = DateTime.Now,
                    Nickname = Path.GetFileName(filePath),
                });
            }

            Items = items;

            ItemUpdated?.Invoke(null, items);
        }
    }
}
