// ──────────────────────────────────────────────────────────────────────────────
//  FutureTaxSortConfig.cs  —  Persisted user-defined FutureTax preview row order
// ──────────────────────────────────────────────────────────────────────────────
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace GifiMapper
{
    public class FutureTaxSortConfig
    {
        public Dictionary<string, int> Order { get; set; } = new();

        private static string ConfigPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "GifiMapper", "futuretax_sort_order.json");

        public static FutureTaxSortConfig Load()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    var json = File.ReadAllText(ConfigPath);
                    var config = JsonSerializer.Deserialize<FutureTaxSortConfig>(json);
                    if (config != null) return config;
                }
            }
            catch { /* ignore — fall back to defaults */ }

            return new FutureTaxSortConfig();
        }

        public static void Save(FutureTaxSortConfig config)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);
                var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigPath, json);
            }
            catch { /* ignore — remembering sort order is a convenience, not critical */ }
        }
    }
}
