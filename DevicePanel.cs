using NAudio.CoreAudioApi;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace XJAudioLatencyTester
{
    public class DevicePanel : Panel
    {
        private ComboBox deviceCombo;
        private Button addButton;
        private Button removeButton;
        private MMDeviceCollection wasapiDevices;

        public int Index { get; set; }
        public bool IsFirstPanel { get; set; }
        public event Action<int> RemoveRequested;
        public event Action AddRequested;

        public int SelectedDeviceId => deviceCombo.SelectedIndex;
        public string SelectedDeviceName => deviceCombo.SelectedItem?.ToString() ?? string.Empty;

        public string GetSelectedDeviceId()
        {
            int idx = deviceCombo.SelectedIndex;
            return (idx >= 0 && wasapiDevices != null && idx < wasapiDevices.Count)
                ? wasapiDevices[idx].ID : null;
        }

        public DevicePanel(int index, bool isFirstPanel)
        {
            Index = index;
            IsFirstPanel = isFirstPanel;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Size = new Size(580, 30);

            deviceCombo = new ComboBox
            {
                Location = new Point(0, 5),
                Size = new Size(450, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            addButton = new Button
            {
                Text = "+",
                Location = new Point(455, 5),
                Size = new Size(25, 25),
                BackColor = Color.LightGreen
            };
            addButton.Click += (s, e) => AddRequested?.Invoke();

            removeButton = new Button
            {
                Text = "-",
                Location = new Point(485, 5),
                Size = new Size(25, 25),
                BackColor = Color.LightCoral
            };
            removeButton.Click += (s, e) => RemoveRequested?.Invoke(Index);
            removeButton.Visible = !IsFirstPanel;

            this.Controls.AddRange(new Control[] { deviceCombo, addButton, removeButton });
        }

        public void LoadOutputDevices()
        {
            deviceCombo.Items.Clear();
            var enumerator = new MMDeviceEnumerator();
            wasapiDevices = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);
            foreach (var device in wasapiDevices)
                deviceCombo.Items.Add(device.FriendlyName);
            if (deviceCombo.Items.Count > 0)
                deviceCombo.SelectedIndex = 0;
        }

        public void SetSelectedDevice(int index)
        {
            if (index >= 0 && index < deviceCombo.Items.Count)
                deviceCombo.SelectedIndex = index;
        }
    }
}
