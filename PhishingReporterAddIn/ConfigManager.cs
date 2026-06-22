using System;
using System.IO;
using System.Web.Script.Serialization;

namespace PhishingReporterAddIn
{
    public class PhishingConfig
    {
        public string ApiUrl { get; set; }
        public int TimeoutSeconds { get; set; }
    }

    public static class ConfigManager
    {
        private static readonly string ConfigDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "PhishingReporter"
        );

        private static readonly string ConfigFilePath = Path.Combine(ConfigDir, "config.json");

        public static PhishingConfig Load()
        {
            try
            {
                if (!Directory.Exists(ConfigDir))
                {
                    Directory.CreateDirectory(ConfigDir);
                }

                if (!File.Exists(ConfigFilePath))
                {
                    var defaultConfig = new PhishingConfig
                    {
                        ApiUrl = "https://webhook.site/fba8fef5-2971-48a3-9422-22bb26bba0dd",
                        TimeoutSeconds = 30
                    };
                    Save(defaultConfig);
                    return defaultConfig;
                }

                string json = File.ReadAllText(ConfigFilePath);
                var serializer = new JavaScriptSerializer();
                var config = serializer.Deserialize<PhishingConfig>(json);
                if (config == null)
                {
                    throw new Exception("Failed to deserialize configuration file.");
                }
                return config;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error loading config: " + ex.Message);
                // Return default fallback configuration
                return new PhishingConfig
                {
                    ApiUrl = "https://webhook.site/fba8fef5-2971-48a3-9422-22bb26bba0dd",
                    TimeoutSeconds = 30
                };
            }
        }

        public static void Save(PhishingConfig config)
        {
            try
            {
                if (!Directory.Exists(ConfigDir))
                {
                    Directory.CreateDirectory(ConfigDir);
                }

                var serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(config);
                File.WriteAllText(ConfigFilePath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error saving config: " + ex.Message);
            }
        }
    }
}
