// ──────────────────────────────────────────────────────────────────────────────
//  AppSettings.cs  —  Persisted user settings (GIFI Mapper)
// ──────────────────────────────────────────────────────────────────────────────
using System;
using System.IO;
using System.Text.Json;

namespace GifiMapper
{
    public class AppSettings
    {
        public string? GifiMappingPath { get; set; }

        private static string SettingsPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "GifiMapper", "settings.json");

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var json = File.ReadAllText(SettingsPath);
                    var settings = JsonSerializer.Deserialize<AppSettings>(json);
                    if (settings != null) return settings;
                }
            }
            catch { /* ignore — fall back to defaults */ }

            return new AppSettings();
        }

        public static void Save(AppSettings settings)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsPath, json);
            }
            catch { /* ignore — remembering the path is a convenience, not critical */ }
        }
    }
}
