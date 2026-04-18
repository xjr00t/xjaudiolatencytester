using System;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace XJAudioLatencyTester
{
    [Serializable]
    public class AppSettings
    {
        public int[] SelectedOutputDeviceIds { get; set; } = new int[0];
        public int SelectedInputDeviceId { get; set; } = -1;
        public float DetectionThreshold { get; set; } = 0.1f;
        public bool BluetoothWarmupEnabled { get; set; } = true;
        public int TestDurationMs { get; set; } = 5000;

        private static readonly string ConfigFilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "AppSettings.xml");

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var serializer = new XmlSerializer(typeof(AppSettings));
                    using (var fs = new FileStream(ConfigFilePath, FileMode.Open))
                    {
                        var settings = serializer.Deserialize(fs) as AppSettings ?? new AppSettings();
                        if (settings.SelectedOutputDeviceIds == null)
                            settings.SelectedOutputDeviceIds = new int[0];
                        return settings;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading settings: {ex.Message}");
            }
            return new AppSettings();
        }

        public void Save()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(AppSettings));
                using (var fs = new FileStream(ConfigFilePath, FileMode.Create))
                {
                    serializer.Serialize(fs, this);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving settings: {ex.Message}");
            }
        }

        public static void ResetToDefaults()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                    File.Delete(ConfigFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error resetting settings: {ex.Message}");
            }
        }

        public bool HasChanged(AppSettings other)
        {
            var myIds = SelectedOutputDeviceIds ?? new int[0];
            var otherIds = other.SelectedOutputDeviceIds ?? new int[0];
            return SelectedInputDeviceId != other.SelectedInputDeviceId ||
                   !myIds.SequenceEqual(otherIds) ||
                   Math.Abs(DetectionThreshold - other.DetectionThreshold) > 0.001f ||
                   BluetoothWarmupEnabled != other.BluetoothWarmupEnabled ||
                   TestDurationMs != other.TestDurationMs;
        }
    }
}
