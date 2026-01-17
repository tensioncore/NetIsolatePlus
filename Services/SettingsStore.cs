﻿using System.IO;
using System.Text.Json;

namespace NetIsolatePlus.Services
{
    public class SettingsStore
    {
        private readonly string _dir;
        private readonly string _file;

        public SettingsStore(string vendor = "Tensioncore Administration Services", string product = "NetIsolatePlus")
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _dir = Path.Combine(appData, vendor, product);
            _file = Path.Combine(_dir, "settings.json");
            Directory.CreateDirectory(_dir);

            TryMigrateLegacy(appData);
        }

        private void TryMigrateLegacy(string appData)
        {
            var legacyDir = Path.Combine(appData, "NetIsolate");
            var legacyFile = Path.Combine(legacyDir, "settings.json");

            if (File.Exists(legacyFile) && !File.Exists(_file))
            {
                Directory.CreateDirectory(_dir);
                try { File.Copy(legacyFile, _file, overwrite: true); } catch { /* best effort */ }
            }
        }

        public T Load<T>(string key, T @default)
        {
            try
            {
                if (!File.Exists(_file)) return @default;

                using var doc = JsonDocument.Parse(File.ReadAllText(_file));
                if (doc.RootElement.TryGetProperty(key, out var el))
                    return JsonSerializer.Deserialize<T>(el.GetRawText()) ?? @default;
            }
            catch { }
            return @default;
        }

        public void Save<T>(string key, T value)
        {
            try
            {
                Directory.CreateDirectory(_dir);

                JsonElement root = default;

                if (File.Exists(_file))
                {
                    try
                    {
                        root = JsonSerializer.Deserialize<JsonElement>(File.ReadAllText(_file));
                    }
                    catch
                    {
                        // If existing JSON is corrupted/invalid, treat as empty and still save the new key.
                        root = default;
                    }
                }

                using var ms = new MemoryStream();
                using var writer = new Utf8JsonWriter(ms, new JsonWriterOptions { Indented = true });

                writer.WriteStartObject();

                if (root.ValueKind == JsonValueKind.Object)
                {
                    foreach (var prop in root.EnumerateObject())
                    {
                        if (prop.NameEquals(key)) continue;
                        prop.WriteTo(writer);
                    }
                }

                writer.WritePropertyName(key);
                JsonSerializer.Serialize(writer, value);
                writer.WriteEndObject();
                writer.Flush();

                // Atomic-ish write: write temp then move over settings.json
                var tmp = Path.Combine(_dir, "settings.json.tmp");
                File.WriteAllBytes(tmp, ms.ToArray());

                try
                {
                    File.Move(tmp, _file, overwrite: true);
                }
                catch
                {
                    // Fallback if Move(overwrite) fails for any reason
                    File.Copy(tmp, _file, overwrite: true);
                    try { File.Delete(tmp); } catch { }
                }
            }
            catch { }
        }
    }
}
