using NAudio.CoreAudioApi;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace XJAudioLatencyTester
{
    public partial class Form1 : Form
    {
        private ComboBox inputDevicesCombo;
        private Button testButton;
        private Label statusLabel;
        private Panel outputDevicesPanel;
        private ProgressBar vuMeter;
        private Label vuLabel;
        private TrackBar thresholdTrackBar;
        private Label thresholdLabel;
        private Label thresholdValueLabel;
        private Button calibrateButton;
        private TextBox resultsTextBox;
        private TextBox relativeResultsTextBox;
        private TextBox logTextBox;
        private Button clearLogButton;
        private Button saveDataButton;
        private Button resetSettingsButton;
        private CheckBox bluetoothWarmupCheckBox;
        private List<DevicePanel> devicePanels = new List<DevicePanel>();
        private WaveInEvent waveIn;
        private bool isMonitoring = false;
        private float currentVolume = 0;
        private float threshold = 0.1f;  // amplitude in [0..1], default ≈ -20 dB
        private byte[] lastRecordedData;
        private string testStartTime;
        private Dictionary<string, byte[]> deviceRecordedData = new Dictionary<string, byte[]>();
        private AppSettings appSettings;
        private AppSettings settingsAtStartup;  // snapshot used to detect changes on exit

        public Form1()
        {
            appSettings = AppSettings.Load();
            settingsAtStartup = AppSettings.Load();

            InitializeComponent();
            this.AutoScroll = true;
            this.MinimumSize = new Size(700, 600);

            using (var stream = typeof(Form1).Assembly.GetManifestResourceStream("XJAudioLatencyTester.Properties.app.ico"))
                if (stream != null) this.Icon = new System.Drawing.Icon(stream);

            this.FormClosing += Form1_FormClosing;

            UpdateThresholdDisplay();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            var currentSettings = new AppSettings
            {
                SelectedInputDeviceId = inputDevicesCombo.SelectedIndex,
                SelectedOutputDeviceIds = devicePanels.Select(p => p.SelectedDeviceId).ToArray(),
                DetectionThreshold = threshold,
                BluetoothWarmupEnabled = bluetoothWarmupCheckBox.Checked,
                TestDurationMs = appSettings.TestDurationMs
            };

            if (currentSettings.HasChanged(settingsAtStartup))
            {
                var result = MessageBox.Show(
                    "Settings have been changed. Do you want to save them?",
                    "Save Settings?",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                    SaveSettings();
                else if (result == DialogResult.Cancel)
                    e.Cancel = true;
            }
        }

        private void InitializeComponent()
        {
            this.Text = "XJ Audio Latency Tester";
            this.Size = new Size(700, 1070);
            this.StartPosition = FormStartPosition.CenterScreen;

            var titleLabel = new Label
            {
                Text = "XJ Audio Latency Tester",
                Font = new Font("Arial", 12, FontStyle.Bold),
                Location = new Point(20, 20),
                Size = new Size(500, 25),
                ForeColor = Color.DarkBlue
            };

            var outputSectionLabel = new Label
            {
                Text = "Output Devices:",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(20, 60),
                Size = new Size(200, 20)
            };

            outputDevicesPanel = new Panel
            {
                Location = new Point(20, 85),
                Size = new Size(600, 150),
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle
            };

            var inputSectionLabel = new Label
            {
                Text = "Input Device (Microphone):",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(20, 250),
                Size = new Size(200, 20)
            };

            inputDevicesCombo = new ComboBox
            {
                Location = new Point(20, 275),
                Size = new Size(400, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            inputDevicesCombo.SelectedIndexChanged += InputDevicesCombo_SelectedIndexChanged;

            var vuMeterLabel = new Label
            {
                Text = "Input Level:",
                Location = new Point(20, 320),
                Size = new Size(100, 20)
            };

            vuMeter = new ProgressBar
            {
                Location = new Point(28, 345),
                Size = new Size(384, 25),
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                ForeColor = Color.LimeGreen
            };

            vuLabel = new Label
            {
                Text = "-∞ dB",
                Location = new Point(418, 345),
                Size = new Size(80, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            var thresholdSectionLabel = new Label
            {
                Text = "Detection Threshold (dB):",
                Location = new Point(20, 385),
                Size = new Size(180, 20)
            };

            // Slider maps 0..100 to -60..+40 dB via: dB = value - 60
            thresholdTrackBar = new TrackBar
            {
                Location = new Point(20, 410),
                Size = new Size(400, 45),
                Minimum = 0,
                Maximum = 100,
                Value = 50,
                TickFrequency = 10,
                SmallChange = 1,
                LargeChange = 10
            };
            thresholdTrackBar.ValueChanged += ThresholdTrackBar_ValueChanged;

            thresholdValueLabel = new Label
            {
                Text = "-20 dB",
                Location = new Point(430, 415),
                Size = new Size(60, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            thresholdLabel = new Label
            {
                Text = "Left = lower threshold / more sensitive (noisy), Right = higher threshold / less sensitive",
                Location = new Point(20, 455),
                Size = new Size(400, 40),
                ForeColor = Color.Gray,
                Font = new Font("Arial", 8)
            };

            calibrateButton = new Button
            {
                Text = "Calibrate Threshold",
                Location = new Point(450, 490),
                Size = new Size(150, 25),
                BackColor = Color.LightYellow
            };
            calibrateButton.Click += CalibrateButton_Click;

            resetSettingsButton = new Button
            {
                Text = "Reset Settings",
                Location = new Point(450, 525),
                Size = new Size(150, 25),
                BackColor = Color.LightCoral
            };
            resetSettingsButton.Click += ResetSettingsButton_Click;

            bluetoothWarmupCheckBox = new CheckBox
            {
                Text = "Bluetooth Warm-up (1s init)",
                Location = new Point(20, 535),
                Size = new Size(200, 20),
                Checked = true
            };

            testButton = new Button
            {
                Text = "Run Latency Analysis",
                Location = new Point(20, 490),
                Size = new Size(200, 40),
                BackColor = Color.LightBlue,
                Font = new Font("Arial", 10, FontStyle.Bold)
            };
            testButton.Click += TestButton_Click;

            statusLabel = new Label
            {
                Text = "Select devices and click Run",
                Location = new Point(230, 500),
                Size = new Size(300, 20),
                ForeColor = Color.Gray
            };

            var resultsLabel = new Label
            {
                Text = "Test Results:",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(20, 570),
                Size = new Size(200, 20)
            };

            resultsTextBox = new TextBox
            {
                Location = new Point(20, 595),
                Size = new Size(600, 100),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9)
            };

            var relativeResultsLabel = new Label
            {
                Text = "Relative Latency:",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(20, 710),
                Size = new Size(200, 20)
            };

            relativeResultsTextBox = new TextBox
            {
                Location = new Point(20, 735),
                Size = new Size(600, 100),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9)
            };

            var logLabel = new Label
            {
                Text = "Diagnostic Log:",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(20, 850),
                Size = new Size(200, 20)
            };

            clearLogButton = new Button
            {
                Text = "Clear Log",
                Location = new Point(500, 850),
                Size = new Size(70, 25),
                BackColor = Color.LightGray
            };
            clearLogButton.Click += ClearLogButton_Click;

            saveDataButton = new Button
            {
                Text = "Save Data",
                Location = new Point(580, 850),
                Size = new Size(70, 25),
                BackColor = Color.LightYellow,
                Enabled = false
            };
            saveDataButton.Click += SaveDataButton_Click;

            logTextBox = new TextBox
            {
                Location = new Point(20, 875),
                Size = new Size(600, 150),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Font = new Font("Consolas", 8),
                WordWrap = false
            };

            this.Controls.AddRange(new Control[]
            {
                titleLabel,
                outputSectionLabel,
                outputDevicesPanel,
                inputSectionLabel,
                inputDevicesCombo,
                vuMeterLabel,
                vuMeter,
                vuLabel,
                thresholdSectionLabel,
                thresholdTrackBar,
                thresholdValueLabel,
                thresholdLabel,
                testButton,
                bluetoothWarmupCheckBox,
                calibrateButton,
                resetSettingsButton,
                statusLabel,
                resultsLabel,
                resultsTextBox,
                relativeResultsLabel,
                relativeResultsTextBox,
                logLabel,
                clearLogButton,
                saveDataButton,
                logTextBox
            });

            AddDevicePanel();
            LoadAudioDevices();
            ApplySettings();
            UpdateThresholdDisplay();

            AddLog("Application started");
        }

        private void AddLog(string message)
        {
            if (logTextBox.InvokeRequired)
            {
                logTextBox.Invoke(new Action<string>(AddLog), message);
                return;
            }

            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            logTextBox.AppendText($"[{timestamp}] {message}\r\n");
            logTextBox.SelectionStart = logTextBox.Text.Length;
            logTextBox.ScrollToCaret();
        }

        private void CalibrateButton_Click(object sender, EventArgs e)
        {
            if (inputDevicesCombo.SelectedIndex == -1)
            {
                MessageBox.Show("Please select an input device first.");
                return;
            }

            try
            {
                AddLog("Starting threshold calibration...");
                var calibrator = new ThresholdCalibrator();
                float calibratedThreshold = calibrator.Calibrate(inputDevicesCombo.SelectedIndex);

                // Convert amplitude to slider position: dB = 20·log10(amp), slider = dB + 60
                float calibratedDB = (float)(20 * Math.Log10(calibratedThreshold));
                int sliderValue = Math.Max(0, Math.Min(100, (int)(calibratedDB + 60)));
                thresholdTrackBar.Value = sliderValue;

                AddLog($"Calibrated: amplitude={calibratedThreshold:F6}, dB={calibratedDB:F1}, slider={sliderValue}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Calibration failed: {ex.Message}");
                AddLog($"Calibration ERROR: {ex.Message}");
            }
        }

        private void ClearLogButton_Click(object sender, EventArgs e)
        {
            logTextBox.Clear();
            AddLog("Log cleared");
        }

        private void ResetSettingsButton_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(
                "Are you sure you want to reset all settings to default values?",
                "Reset Settings",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.Yes)
            {
                AppSettings.ResetToDefaults();
                appSettings = new AppSettings();
                threshold = appSettings.DetectionThreshold;
                thresholdTrackBar.Value = 50;
                bluetoothWarmupCheckBox.Checked = true;

                if (inputDevicesCombo.Items.Count > 0)
                    inputDevicesCombo.SelectedIndex = 0;

                AddLog("Settings reset to default values");
                MessageBox.Show("Settings have been reset to defaults.", "Reset Complete");
            }
        }

        private void SaveSettings()
        {
            appSettings.SelectedInputDeviceId = inputDevicesCombo.SelectedIndex;
            appSettings.SelectedOutputDeviceIds = devicePanels.Select(p => p.SelectedDeviceId).ToArray();
            appSettings.DetectionThreshold = threshold;
            appSettings.BluetoothWarmupEnabled = bluetoothWarmupCheckBox.Checked;
            appSettings.Save();
            AddLog("Settings saved");
        }

        private void ApplySettings()
        {
            var savedOutputIds = appSettings.SelectedOutputDeviceIds ?? new int[0];
            if (savedOutputIds.Length > 0)
            {
                while (devicePanels.Count > savedOutputIds.Length)
                    RemoveDevicePanel(devicePanels.Count - 1);
                while (devicePanels.Count < savedOutputIds.Length)
                    AddDevicePanel();
                for (int i = 0; i < savedOutputIds.Length; i++)
                    devicePanels[i].SetSelectedDevice(savedOutputIds[i]);
            }

            if (appSettings.SelectedInputDeviceId >= 0 && appSettings.SelectedInputDeviceId < inputDevicesCombo.Items.Count)
                inputDevicesCombo.SelectedIndex = appSettings.SelectedInputDeviceId;

            threshold = appSettings.DetectionThreshold;
            float dB = (float)(20 * Math.Log10(threshold));
            int sliderValue = Math.Max(0, Math.Min(100, (int)(dB + 60)));
            thresholdTrackBar.Value = sliderValue;

            bluetoothWarmupCheckBox.Checked = appSettings.BluetoothWarmupEnabled;
        }

        private void StartNewTestSession()
        {
            testStartTime = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            deviceRecordedData.Clear();
        }

        private void SaveDeviceData(byte[] data, string deviceName, int deviceIndex, bool isMultipleDevices)
        {
            if (data == null || data.Length == 0) return;

            string safeDeviceName = CleanFileName(deviceName);
            string fileName = $"{testStartTime}_{deviceIndex:00}_{safeDeviceName}.wav";
            deviceRecordedData[fileName] = data;
        }

        private string CleanFileName(string fileName)
        {
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                fileName = fileName.Replace(c, '_');
            // Colons are valid on some FS but appear in WASAPI device names
            fileName = fileName.Replace(':', '_');
            return fileName;
        }

        private void UpdateSaveButtonSafe(bool enable)
        {
            if (saveDataButton.InvokeRequired)
            {
                saveDataButton.Invoke(new Action<bool>(UpdateSaveButtonSafe), enable);
                return;
            }
            saveDataButton.Enabled = enable;
        }

        private void SaveDataButton_Click(object sender, EventArgs e)
        {
            if (deviceRecordedData.Count == 0)
            {
                MessageBox.Show("No recorded data available to save.");
                return;
            }

            try
            {
                if (deviceRecordedData.Count == 1)
                {
                    var item = deviceRecordedData.First();
                    using (var saveDialog = new SaveFileDialog())
                    {
                        saveDialog.FileName = Path.ChangeExtension(item.Key, ".wav");
                        saveDialog.Filter = "WAV files (*.wav)|*.wav";
                        saveDialog.Title = "Save Recorded Data";

                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            SaveAsWav(saveDialog.FileName, item.Value);
                            AddLog($"Recorded data saved to: {saveDialog.FileName}");
                        }
                    }
                }
                else
                {
                    using (var folderDialog = new FolderBrowserDialog())
                    {
                        folderDialog.Description = "Select folder to save recorded data for all devices";
                        folderDialog.ShowNewFolderButton = true;

                        if (folderDialog.ShowDialog() == DialogResult.OK)
                        {
                            foreach (var item in deviceRecordedData)
                            {
                                string filePath = Path.Combine(folderDialog.SelectedPath, item.Key);
                                SaveAsWav(filePath, item.Value);
                                AddLog($"Saved: {item.Key}");
                            }
                            AddLog($"All recorded data saved to: {folderDialog.SelectedPath}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving data: {ex.Message}");
                AddLog($"ERROR saving data: {ex.Message}");
            }
        }

        private void SaveAsWav(string filename, byte[] data)
        {
            using (var writer = new WaveFileWriter(filename, new WaveFormat(44100, 16, 1)))
                writer.Write(data, 0, data.Length);
        }

        private void AddDevicePanel()
        {
            var devicePanel = new DevicePanel(devicePanels.Count, devicePanels.Count == 0);
            devicePanel.Location = new Point(0, devicePanels.Count * 35);
            devicePanel.RemoveRequested += (index) => RemoveDevicePanel(index);
            devicePanel.AddRequested += () => AddDevicePanel();

            outputDevicesPanel.Controls.Add(devicePanel);
            devicePanels.Add(devicePanel);
            devicePanel.LoadOutputDevices();

            outputDevicesPanel.Height = Math.Min(150, devicePanels.Count * 35 + 10);
        }

        private void RemoveDevicePanel(int index)
        {
            if (devicePanels.Count <= 1) return;

            var panelToRemove = devicePanels[index];
            outputDevicesPanel.Controls.Remove(panelToRemove);
            devicePanels.RemoveAt(index);

            for (int i = 0; i < devicePanels.Count; i++)
            {
                devicePanels[i].Index = i;
                devicePanels[i].Location = new Point(0, i * 35);
                devicePanels[i].IsFirstPanel = (i == 0);
            }

            outputDevicesPanel.Height = Math.Min(150, devicePanels.Count * 35 + 10);
        }

        private void LoadAudioDevices()
        {
            try
            {
                statusLabel.Text = "Loading audio devices...";
                AddLog("Loading audio devices...");

                foreach (var panel in devicePanels)
                    panel.LoadOutputDevices();

                inputDevicesCombo.Items.Clear();
                for (int i = 0; i < WaveIn.DeviceCount; i++)
                {
                    var capabilities = WaveIn.GetCapabilities(i);
                    inputDevicesCombo.Items.Add($"{i}: {capabilities.ProductName}");
                }

                if (inputDevicesCombo.Items.Count > 0)
                    inputDevicesCombo.SelectedIndex = 0;

                statusLabel.Text = "Ready to test latency";
                AddLog($"Loaded {WaveOut.DeviceCount} output devices and {WaveIn.DeviceCount} input devices");
            }
            catch (Exception ex)
            {
                statusLabel.Text = $"Error loading devices: {ex.Message}";
                AddLog($"ERROR loading devices: {ex.Message}");
            }
        }

        private void InputDevicesCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            AddLog($"Input device changed to: {inputDevicesCombo.SelectedItem}");
            RestartVUMonitoring();
        }

        private void ThresholdTrackBar_ValueChanged(object sender, EventArgs e)
        {
            // Slider 0..100 → dB = value - 60 → amplitude = 10^(dB/20)
            float dB = thresholdTrackBar.Value - 60;
            threshold = (float)Math.Pow(10, dB / 20.0);

            UpdateThresholdDisplay();
            AddLog($"Threshold changed to {thresholdTrackBar.Value} ({dB:F1} dB, amplitude: {threshold:F4})");
        }

        private void UpdateThresholdDisplay()
        {
            float dB = thresholdTrackBar.Value - 60;
            thresholdValueLabel.Text = $"{dB:F0} dB";
        }

        private void StartVUMonitoring()
        {
            try
            {
                if (inputDevicesCombo.SelectedIndex == -1)
                    return;

                waveIn = new WaveInEvent
                {
                    DeviceNumber = inputDevicesCombo.SelectedIndex,
                    WaveFormat = new WaveFormat(44100, 16, 1),
                    BufferMilliseconds = 50
                };
                waveIn.DataAvailable += OnDataAvailable;
                waveIn.RecordingStopped += OnRecordingStopped;

                waveIn.StartRecording();
                isMonitoring = true;
                AddLog("VU monitoring started");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to start VU monitoring: {ex.Message}");
                statusLabel.Text = $"VU Meter Error: {ex.Message}";
                AddLog($"ERROR starting VU monitoring: {ex.Message}");
            }
        }

        private void RestartVUMonitoring()
        {
            if (isMonitoring)
            {
                waveIn?.StopRecording();
                waveIn?.Dispose();
                isMonitoring = false;
                AddLog("VU monitoring stopped");
            }

            if (inputDevicesCombo.SelectedIndex != -1)
                StartVUMonitoring();
        }

        private void OnDataAvailable(object sender, WaveInEventArgs e)
        {
            float max = 0;
            for (int index = 0; index < e.BytesRecorded; index += 2)
            {
                short sample = (short)((e.Buffer[index + 1] << 8) | e.Buffer[index]);
                float sample32 = Math.Abs(sample / 32768f);
                if (sample32 > max) max = sample32;
            }

            currentVolume = max;

            if (vuMeter.InvokeRequired)
                vuMeter.Invoke(new Action(() => UpdateVUMeter(currentVolume)));
            else
                UpdateVUMeter(currentVolume);
        }

        private void UpdateVUMeter(float volume)
        {
            double dB = volume > 0 ? 20 * Math.Log10(volume) : -100;
            int meterValue = (int)Math.Max(0, Math.Min(100, (dB + 60) * 100 / 60));
            vuMeter.Value = meterValue;

            if (dB > -3) vuMeter.ForeColor = Color.Red;
            else if (dB > -10) vuMeter.ForeColor = Color.Orange;
            else vuMeter.ForeColor = Color.LimeGreen;

            vuLabel.Text = dB > -60 ? $"{dB:F1} dB" : "-∞ dB";
        }

        private void OnRecordingStopped(object sender, StoppedEventArgs e)
        {
            isMonitoring = false;
        }

        private async void TestButton_Click(object sender, EventArgs e)
        {
            if (devicePanels.Count == 0 || inputDevicesCombo.SelectedIndex == -1)
            {
                MessageBox.Show("Please select at least one output device and one input device.");
                return;
            }

            StartNewTestSession();

            var devicesToTest = new List<DeviceTestInfo>();
            foreach (var panel in devicePanels)
            {
                string wasapiId = panel.GetSelectedDeviceId();
                if (wasapiId != null)
                {
                    devicesToTest.Add(new DeviceTestInfo
                    {
                        DeviceId = panel.SelectedDeviceId,
                        DeviceName = panel.SelectedDeviceName,
                        WasapiDeviceId = wasapiId
                    });
                }
            }

            if (devicesToTest.Count == 0)
            {
                MessageBox.Show("Please select at least one output device.");
                return;
            }

            int inputDeviceId = inputDevicesCombo.SelectedIndex;

            testButton.Enabled = false;
            saveDataButton.Enabled = false;
            ClearResults();
            float currentThresholdDB = (float)(20 * Math.Log10(threshold));
            AddLog($"=== Starting latency test for {devicesToTest.Count} device(s) ===");
            AddLog($"Threshold: {currentThresholdDB:F1} dB (amplitude: {threshold:F6}), Input device: {inputDeviceId}");
            statusLabel.Text = $"Testing {devicesToTest.Count} device(s)...";

            try
            {
                var results = await TestDevicesLatency(devicesToTest, inputDeviceId);
                DisplayResults(results);

                if (results.Count > 1)
                    DisplayRelativeResults(results);

                statusLabel.Text = "Test completed";
                AddLog("=== Test completed ===");
            }
            catch (Exception ex)
            {
                statusLabel.Text = $"Test failed: {ex.Message}";
                AddLog($"=== Test failed: {ex.Message} ===");
            }
            finally
            {
                testButton.Enabled = true;
            }
        }

        private System.Threading.Tasks.Task<List<DeviceTestResult>> TestDevicesLatency(
            List<DeviceTestInfo> devices, int inputDeviceId)
        {
            return System.Threading.Tasks.Task.Run(() =>
            {
                var results = new List<DeviceTestResult>();
                bool enableBluetoothWarmup = bluetoothWarmupCheckBox.Checked;
                var latencyMeasurer = new ImprovedLatencyMeasurer(threshold, AddLog, enableBluetoothWarmup);

                for (int i = 0; i < devices.Count; i++)
                {
                    var device = devices[i];
                    UpdateStatusSafe($"Testing device {i + 1} of {devices.Count}: {device.DeviceName}");
                    AddLog($"Testing device {i + 1}: {device.DeviceName} (ID: {device.DeviceId})");

                    try
                    {
                        var (latency, recordedData) = latencyMeasurer.MeasureLatency(
                            device.WasapiDeviceId, device.DeviceId, inputDeviceId);

                        var result = new DeviceTestResult
                        {
                            DeviceName = device.DeviceName,
                            LatencyMs = latency,
                            Success = true
                        };
                        results.Add(result);
                        UpdateResultsSafe(results);
                        AddLog($"Device {device.DeviceName} - Latency: {latency:F1} ms");

                        SaveDeviceData(recordedData, device.DeviceName, i + 1, devices.Count > 1);
                        UpdateSaveButtonSafe(true);
                    }
                    catch (Exception ex)
                    {
                        var result = new DeviceTestResult
                        {
                            DeviceName = device.DeviceName,
                            ErrorMessage = ex.Message,
                            Success = false
                        };
                        results.Add(result);
                        UpdateResultsSafe(results);
                        AddLog($"Device {device.DeviceName} - ERROR: {ex.Message}");
                    }
                }

                return results;
            });
        }

        private void UpdateResultsSafe(List<DeviceTestResult> results)
        {
            if (resultsTextBox.InvokeRequired)
            {
                resultsTextBox.Invoke(new Action<List<DeviceTestResult>>(UpdateResultsSafe), results);
                return;
            }

            var sb = new System.Text.StringBuilder();
            foreach (var result in results)
            {
                sb.AppendLine(result.Success
                    ? $"{result.DeviceName} - {result.LatencyMs:F1} ms"
                    : $"{result.DeviceName} - Error: {result.ErrorMessage}");
            }
            resultsTextBox.Text = sb.ToString();
        }

        private void UpdateStatusSafe(string status)
        {
            if (statusLabel.InvokeRequired)
                statusLabel.Invoke(new Action(() => statusLabel.Text = status));
            else
                statusLabel.Text = status;
        }

        private void ClearResults()
        {
            resultsTextBox.Text = string.Empty;
            relativeResultsTextBox.Text = string.Empty;
        }

        private void DisplayResults(List<DeviceTestResult> results)
        {
            if (resultsTextBox.InvokeRequired)
                resultsTextBox.Invoke(new Action(() => DisplayResultsInternal(results)));
            else
                DisplayResultsInternal(results);
        }

        private void DisplayResultsInternal(List<DeviceTestResult> results)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var result in results)
            {
                sb.AppendLine(result.Success
                    ? $"{result.DeviceName} - {result.LatencyMs:F1} ms"
                    : $"{result.DeviceName} - Error: {result.ErrorMessage}");
            }
            resultsTextBox.Text = sb.ToString();
        }

        private void DisplayRelativeResults(List<DeviceTestResult> results)
        {
            if (relativeResultsTextBox.InvokeRequired)
                relativeResultsTextBox.Invoke(new Action(() => DisplayRelativeResultsInternal(results)));
            else
                DisplayRelativeResultsInternal(results);
        }

        private void DisplayRelativeResultsInternal(List<DeviceTestResult> results)
        {
            var successfulResults = results.Where(r => r.Success).ToList();
            if (successfulResults.Count < 2)
            {
                relativeResultsTextBox.Text = "Not enough successful measurements for relative comparison";
                return;
            }

            // Sort slowest-first; relative value = delay to add in Voicemeeter to sync the device
            var sortedResults = successfulResults.OrderByDescending(r => r.LatencyMs).ToList();
            double maxLatency = sortedResults[0].LatencyMs;

            var sb = new System.Text.StringBuilder();
            foreach (var result in sortedResults)
                sb.AppendLine($"{result.DeviceName} - {maxLatency - result.LatencyMs:F1} ms");

            relativeResultsTextBox.Text = sb.ToString();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (waveIn != null)
            {
                waveIn.StopRecording();
                waveIn.Dispose();
            }
            base.OnFormClosed(e);
        }
    }

    public class DeviceTestInfo
    {
        public int DeviceId { get; set; }
        public string DeviceName { get; set; }
        public string WasapiDeviceId { get; set; }
    }

    public class DeviceTestResult
    {
        public string DeviceName { get; set; }
        public double LatencyMs { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
    }
}
