using HidSharp;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Reflection;


namespace TraySense
{
    public partial class Form1 : Form
    {
        private NotifyIcon _trayIcon;
        private ContextMenuStrip _trayMenu;
        private HidDevice _dualSenseDevice;
        bool debugflag = false;
        //Intervals for checking the controller in milliseconds
        int refreshtime_searchingforcontroler = 5000; //(5s)
        int refreshtime_between_checks = 120000; //(2min)
        string CurrentIcon = "default";
        int batterywarning;



        int refreshtime;
        public Form1()
        {
            // Debug: List all embedded resources in exe
            /*
            
            string[] resourceNames = Assembly.GetExecutingAssembly().GetManifestResourceNames();
            foreach (string resourceName in resourceNames)
            {
                Debug.WriteLine(resourceName);
            }
            */
            refreshtime = refreshtime_searchingforcontroler;
            InitializeComponent();
            InitializeTrayIcon();
            this.FormClosing += Form1_FormClosing;
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            ApplySystemTheme();
            SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;
            LogToRichTextBox("Application started.", Color.GreenYellow);
            StartControllerCheck();

        }

        // Required to set darkmode on windows 10 title bar
        [DllImport("DwmApi")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, int[] attrValue, int attrSize);



        // Set correct title bar mode on start
        protected override void OnHandleCreated(EventArgs e)
        {
            if (GetIsAppDarkTheme())
            {
                //set dark mode for title bar
                if (DwmSetWindowAttribute(Handle, 19, new[] { 1 }, 4) != 0)
                    DwmSetWindowAttribute(Handle, 20, new[] { 1 }, 4);
            }
            else
            {
                //set light mode for title bar
                if (DwmSetWindowAttribute(Handle, 19, new[] { 0 }, 4) != 0)
                    DwmSetWindowAttribute(Handle, 20, new[] { 0 }, 4);
            }
        }

        // Detect Windows theme change
        private void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.General)
                ApplySystemTheme();
            LogToRichTextBox("User Preference Changed Detected", Color.SandyBrown);
        }

        private void ApplySystemTheme()
        {
            bool isSystemDarkTheme = GetIsSystemDarkTheme();
            bool isAppDarkTheme = GetIsAppDarkTheme();
            ApplyTheme(isAppDarkTheme ? Theme.Dark : Theme.Light);

            _trayIcon.Icon = LoadEmbeddedIcon("TraySense.icons." + (isSystemDarkTheme ? "Darkmode" : "Lightmode") + "." + CurrentIcon + ".ico");


            if (!isAppDarkTheme)
            {
                // Set light mode for title bar
                if (DwmSetWindowAttribute(Handle, 19, new[] { 0 }, 4) != 0)
                    DwmSetWindowAttribute(Handle, 20, new[] { 0 }, 4);
            }
            else
            {
                // Set dark mode for title bar
                if (DwmSetWindowAttribute(Handle, 19, new[] { 1 }, 4) != 0)
                    DwmSetWindowAttribute(Handle, 20, new[] { 1 }, 4);
            }

            ApplyContextMenuTheme();
        }

        // App dark theme applies to windows
        private bool GetIsAppDarkTheme()
        {
            int themeValue = (int)Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1);
            return themeValue == 0; // Dark theme = 0, Light theme = 1
        }

        // System dark theme applies to tray icon
        private bool GetIsSystemDarkTheme()
        {
            int themeValue = (int)Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme", 1);
            return themeValue == 0; // Dark theme = 0, Light theme = 1
        }


        private void ApplyTheme(Theme theme)
        {
            this.BackColor = theme == Theme.Dark ? Color.FromArgb(32, 32, 32) : Color.White;
            this.ForeColor = theme == Theme.Dark ? Color.White : Color.Black;
            ApplyThemeToControls(this, theme);
        }

        private void ApplyThemeToControls(Control parent, Theme theme)
        {
            foreach (Control control in parent.Controls)
            {
                control.BackColor = theme == Theme.Dark ? Color.FromArgb(32, 32, 32) : Color.White;
                control.ForeColor = theme == Theme.Dark ? Color.White : Color.Black;

                if (control.HasChildren)
                {
                    ApplyThemeToControls(control, theme);
                }
            }
        }
        // Hide to tray on debug form close
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                MinimizeToTray();
            }
        }

        private void MinimizeToTray()
        {
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Hide();
            _trayIcon.BalloonTipTitle = "TraySense";
            _trayIcon.BalloonTipText = "The application is still running in the system tray.";
            _trayIcon.ShowBalloonTip(1000);
        }

        private ToolStripLabel _batteryInfoLabel; // Declare a ToolStripLabel to show battery info

        private void InitializeTrayIcon()
        {
            // Create a context menu for the tray icon
            _trayMenu = new ContextMenuStrip();

            // Add context menu items
            // Add a non-clickable label at the top of the context menu to show battery info
            _batteryInfoLabel = new ToolStripLabel("Looking for Controller...");
            _batteryInfoLabel.Enabled = false; // Make it non-clickable
            _trayMenu.Items.Add(_batteryInfoLabel);
            // Add a separator to separate the label from other menu items
            _trayMenu.Items.Add(new ToolStripSeparator());
            var exitItem = new ToolStripMenuItem("Exit", null, (s, e) => ExitApplication());

            // Set the Image property of each item to null to remove the extra space
            exitItem.Image = null;

            _trayMenu.Items.Add(exitItem);

            // Create a debug menu item, to add by key combination
            var Debug_Menu = new ToolStripMenuItem("Debug Menu", null, (s, e) => ShowForm());
            Debug_Menu.Image = null;

            // Apply initial theme to the context menu
            ApplyContextMenuTheme();

            // Create the tray icon with the dynamically selected icon
            _trayIcon = new NotifyIcon
            {
                Icon = LoadEmbeddedIcon("TraySense.icons." + (GetIsSystemDarkTheme() ? "Darkmode" : "Lightmode") + ".default.ico"), // Load icon based on the theme
                ContextMenuStrip = _trayMenu,
                Text = "TraySense",
                Visible = true
            };

            // Event handler to detect CTRL+SHIFT on right click
            _trayIcon.MouseClick += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    if (Control.ModifierKeys == (Keys.Control | Keys.Shift))
                    {
                        // Add debug item to the context menu dynamically
                        if (!_trayMenu.Items.Contains(Debug_Menu))
                        {
                            // Insert at index 0
                            _trayMenu.Items.Insert(0, Debug_Menu);
                        }
                    }
                    else
                    {
                        // Ensure debug item is removed if Shift key is not pressed
                        if (_trayMenu.Items.Contains(Debug_Menu))
                        {
                            _trayMenu.Items.Remove(Debug_Menu);
                        }
                    }
                }
            };
        }

        private void ApplyContextMenuTheme()
        {
            //bool isDarkTheme = GetIsDarkTheme(); //dark mode on context menu looks weird
            bool isDarkTheme = false;

            Color menuBackColor = isDarkTheme ? Color.FromArgb(32, 32, 32) : Color.White;
            Color menuForeColor = isDarkTheme ? Color.White : Color.Black;
            Color itemBackColor = isDarkTheme ? Color.FromArgb(45, 45, 45) : Color.White;
            Color itemForeColor = isDarkTheme ? Color.White : Color.Black;

            // Apply colors to the context menu
            _trayMenu.BackColor = menuBackColor;
            _trayMenu.ForeColor = menuForeColor;

            // Apply colors to each item in the menu
            foreach (ToolStripItem item in _trayMenu.Items)
            {
                item.BackColor = itemBackColor;
                item.ForeColor = itemForeColor;
            }
        }

        private void ShowForm()
        {
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            this.Show();
            debugflag = true;
        }

        private void ExitApplication()
        {
            SystemEvents.UserPreferenceChanged -= SystemEvents_UserPreferenceChanged;
            _trayIcon.Visible = false;
            Application.Exit();
        }

        private async void StartControllerCheck()
        {
            while (true)
            {
                await CheckHIDRawData();
                await Task.Delay(refreshtime);
            }
        }

        private async Task CheckHIDRawData()
        {
            try
            {
                // Get the list of current HID devices
                var deviceList = DeviceList.Local;
                // Look for first Dualsense (1356 3302)
                _dualSenseDevice = deviceList.GetHidDevices(1356, 3302).FirstOrDefault();

                if (_dualSenseDevice == null)
                {
                    LogToRichTextBox("DualSense controller not found.", Color.Red);
                    CurrentIcon = "default";
                    _trayIcon.Icon = LoadEmbeddedIcon("TraySense.icons." + (GetIsSystemDarkTheme() ? "Darkmode" : "Lightmode") + "." + CurrentIcon + ".ico");
                    _batteryInfoLabel.Text = "Looking for Controller...";
                    _trayIcon.Text = "TraySense - Looking for Controller...";
                    refreshtime = refreshtime_searchingforcontroler;
                    return;
                }
                else { refreshtime = refreshtime_between_checks; };

                LogToRichTextBox($"Found DualSense controller: {_dualSenseDevice.GetProductName()}", Color.Green);

                // Open a stream to communicate with the controller
                using (var stream = _dualSenseDevice.Open())
                {
                    await Task.Delay(500); // Add a small delay before reading data
                    byte[] inputBuffer = new byte[_dualSenseDevice.GetMaxInputReportLength()];
                    int bytesRead = stream.Read(inputBuffer, 0, inputBuffer.Length);

                    // If data is received from the controller, process it
                    if (bytesRead > 0)
                    {
                        // Convert the raw data to a hex string for display
                        string hexData = string.Join(" ", inputBuffer.Take(bytesRead).Select(b => b.ToString("X2")));
                        LogToRichTextBox($"Raw Data (Length {bytesRead}): {hexData}", Color.Cyan);
                        ProcessControllerData(inputBuffer, bytesRead);
                    }
                    else
                    {
                        CurrentIcon = "default";
                        _trayIcon.Icon = new Icon(Path.Combine(Application.StartupPath, "Icons", GetIsSystemDarkTheme() ? "Darkmode" : "Lightmode", (CurrentIcon + ".ico")));
                        LogToRichTextBox("No data received from controller.", Color.Red);
                    }
                }
            }
            catch (Exception ex)
            {
                CurrentIcon = "default";
                _trayIcon.Icon = LoadEmbeddedIcon("TraySense.icons." + (GetIsSystemDarkTheme() ? "Darkmode" : "Lightmode") + "." + CurrentIcon + ".ico");
                LogToRichTextBox($"Error checking HID raw data: {ex.Message}", Color.Red);
            }
        }

        // Tested on DS firmware A-0520
        private void ProcessControllerData(byte[] inputBuffer, int bytesRead)
        {
            // Check length of data read , USB is 64 and Bluetooth is 78
            if (bytesRead == 64)
            {
                LogToRichTextBox("Controller is in USB mode.", Color.Cyan);
                this.Text = $"USB Mode - Data Length: {bytesRead}";
                // In USB mode battery info is at 53 and charging info is at 54
                // In USB mode charging seems to be 8 if charging and 0 when not
                DisplayBatteryInfo(inputBuffer[53], inputBuffer[54]);
            }
            // Bluetooth has length of 78 and 3 modes:
            else if (bytesRead == 78)
            {
                // 0x31 header - full readout mode includes trackpad, impulse triggers, battery level and charging info, everything !
                if (inputBuffer[0] == 0x31)
                {
                    LogToRichTextBox("Controller is in Bluetooth mode.", Color.Cyan);
                    this.Text = $"Bluetooth Mode - Data Length: {bytesRead}";
                    // In Bluetooth mode 0x31 battery level is at 54 and charging info is at 55
                    // In Bluetooth charging seems to be indicated by 16 if charging and 0 when not
                    DisplayBatteryInfo(inputBuffer[54], inputBuffer[55]);
                }
                // 0x01 header may indicate two bluetooth modes
                else if (inputBuffer[0] == 0x01)
                {
                    // Check for many 0's in the data (ignoring the first few bytes)
                    int zeroCount = 0;
                    for (int i = 1; i < inputBuffer.Length; i++)  // Start checking from byte 1 (after the 0x01 byte)
                    {
                        if (inputBuffer[i] == 0x00)
                        {
                            zeroCount++;
                        }
                    }
                    // 0x01 may send only a few bytes that are indicating button presses and nothing else, most importantly no battery or charging info
                    // If most of data is empty then controller is in "Minimal Bluetooth"
                    // This mode is usually only seen when bluetooth module or pc was restarted and "fresh connection" is made
                    if (zeroCount > 70)
                    {
                        LogToRichTextBox("Controller is in Minimal Bluetooth mode", Color.Pink);
                        CurrentIcon = "unknown";
                        _trayIcon.Icon = LoadEmbeddedIcon("TraySense.icons." + (GetIsSystemDarkTheme() ? "Darkmode" : "Lightmode") + "." + CurrentIcon + ".ico");
                        // Read "Magic Packet" to "wake" controller to 0x31 (Full Bluetooth mode)
                        WaketofullBT();
                    }
                    else
                    {
                        // 0x01 when not mostly empty indicates "Basic Bluetooth", it sends button presses and some more info along with battery info but no charging status
                        // For reliability and charing info I still switch it to Full Bluetooth mode
                        LogToRichTextBox("Controller is in Basic Bluetooth mode.", Color.Cyan);
                        // If charging info is not provided use -111 to indicate N/A
                        DisplayBatteryInfo(inputBuffer[54], -111);
                        // Read "Magic Packet" to "wake" controller to 0x31 (Full Bluetooth mode)
                        WaketofullBT();
                    }
                }
                else
                {
                    LogToRichTextBox($"Error: Unknown Bluetooth mode. First byte: 0x{inputBuffer[0]:X2}", Color.Red);
                }
            }
            else
            {
                LogToRichTextBox($"Error: Invalid mode detected. Data Length: {bytesRead}", Color.Red);
            }
        }

        private void DisplayBatteryInfo(int battery0, int battery1)
        {
            // Last 4 bits mark battery level
            // Battery level is on a scale 0(empty) - 8(full)
            int batterynumber0to8 = (battery0 & 0x0F);
            // Sometimes battery can report status 9 when full
            // Make it 8
            if( batterynumber0to8 == 9)
            {
                batterynumber0to8 = 8;
            }
            int batteryLevelPercent = batterynumber0to8 * 100 / 8;
            // Currently charging is 16 when bluetooth , 8 when usb and 0 when not charging
            // Old projects online only show one value for charging "8" "1000" no mentions of "16" "10000" New DS firmware changed it ?
            bool isCharging = battery1 > 0;

            LogToRichTextBox("Battery: " + batterynumber0to8 + " which equals " + batteryLevelPercent + "%", Color.Yellow);
            LogToRichTextBox("Charging status: " + isCharging + " | raw data: " + battery1, Color.Yellow);

            if (isCharging)
            {
                CurrentIcon = "charging";
                _trayIcon.Icon = LoadEmbeddedIcon("TraySense.icons." + (GetIsSystemDarkTheme() ? "Darkmode" : "Lightmode") + "." + CurrentIcon + ".ico");
                batterywarning = 9;
            }
            else
            {
                if (batterynumber0to8 >= 1)
                {
                    batterywarning = 9;
                }
                else
                {
                    if (batterywarning != batterynumber0to8)
                    {

                        //if (batterynumber0to8 == 1)
                       // {
                       //     _trayIcon.BalloonTipText = "DS#1 battery is at " + batteryLevelPercent + "%";
                       // }
                       //else
                       //{
                            _trayIcon.BalloonTipText = "DS#1 battery is running low !";
                       //}
                        _trayIcon.BalloonTipTitle = "TraySense";
                        _trayIcon.ShowBalloonTip(1000);
                        batterywarning = batterynumber0to8;
                    }
                }
                CurrentIcon = batterynumber0to8.ToString();
                _trayIcon.Icon = LoadEmbeddedIcon("TraySense.icons." + (GetIsSystemDarkTheme() ? "Darkmode" : "Lightmode") + "." + CurrentIcon + ".ico");
            }
            string batteryStatusText = $"{batteryLevelPercent}% {(isCharging ? "(Charging)" : "(Not Charging)")}";

            // If in Minimal bluetooth mode charging info is not provided
            if (battery1 == -111)
            {
                batteryStatusText = $"{batteryLevelPercent}% - Charging: N/A";
            }

            // Update the Tray menu label with the battery info
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateTrayLabel(batteryStatusText)));
            }
            else
            {
                UpdateTrayLabel(batteryStatusText);
            }
        }

        private void UpdateTrayLabel(string batteryStatusText)
        {
            // Update the label text with the current battery status
            _batteryInfoLabel.Text = "DS #1: " + batteryStatusText;

            _trayIcon.Text = "TraySense - DS#1: " + batteryStatusText;
        }
        // If debug menu is open display all debug info there , if it wasn't open then don't collect logs to optimize resources
        private void LogToRichTextBox(string message, Color color)
        {
            if (debugflag)
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(() => LogToRichTextBox(message, color)));
                }
                else
                {
                    const int maxLines = 500;
                    if (richTextBox1.Lines.Length > maxLines)
                        richTextBox1.Lines = richTextBox1.Lines.Skip(richTextBox1.Lines.Length - maxLines).ToArray();

                    richTextBox1.SelectionColor = color;
                    richTextBox1.AppendText(message + Environment.NewLine);
                    richTextBox1.ScrollToCaret();
                }
            }
        }

        // Dualsense sometimes (usually when bluetooth was restarted) sends only button presses and no other info , to trigger it into full info mode it is required to read 0x05 report from it.
        // AFAIK when controller recieves request to read from it 0x05 report it switches to full bluetooth functionality
        void WaketofullBT()
        {
            try
            {
                if (_dualSenseDevice.TryOpen(out HidStream hidStream))
                {
                    using (hidStream)
                    {
                        byte[] buffer = new byte[_dualSenseDevice.GetMaxFeatureReportLength()];
                        buffer[0] = 0x05;  // Report ID

                        hidStream.GetFeature(buffer);

                        LogToRichTextBox("Succesfully sent 0x05 to Controller", Color.LightGreen);
                    }
                }
                else
                {
                    LogToRichTextBox("Couldn't open device", Color.Red);
                }
            }
            catch (Exception ex)
            {
                LogToRichTextBox($"Error: {ex.Message}", Color.Red);
            }
            refreshtime = refreshtime_searchingforcontroler;
        }


        // Enum for themes
        private enum Theme
        {
            Dark,
            Light
        }

        // Retrieve .Ico 's embedded in exe
        private Icon LoadEmbeddedIcon(string resourceName)
        {
            // Get the assembly containing the embedded resource
            Assembly assembly = Assembly.GetExecutingAssembly();

            // Locate the resource by its fully qualified name
            using Stream resourceStream = assembly.GetManifestResourceStream(resourceName);

            if (resourceStream == null)
            {
                throw new FileNotFoundException("Resource not found: " + resourceName);
            }

            // Return the icon
            return new Icon(resourceStream);
        }
    }
}
