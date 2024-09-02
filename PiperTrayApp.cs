using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;
using NAudio.Wave;
using System.Threading.Tasks;
using NAudio.CoreAudioApi;
using System.Runtime.InteropServices;

namespace PiperTrayClassic
{
    public class PiperTrayApp : Form
    {
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const int HOTKEY_ID = 9000;
        private const uint MOD_ALT = 0x0001;
        private const uint VK_Q = 0x51;

        private NotifyIcon trayIcon;
        private ContextMenuStrip contextMenu;
        private ToolStripMenuItem toggleMonitoringMenuItem;
        private ToolStripMenuItem stopPlaybackMenuItem;
        private ToolStripMenuItem exitMenuItem;
        private System.Windows.Forms.Timer clipboardTimer;
        private string lastClipboardContent = "";
        private bool isMonitoring = true;
        private string piperPath;
        private string logFilePath;
        private bool isProcessing = false;
        private WaveOutEvent currentWaveOut;
        private bool isLoggingEnabled = false;

        public PiperTrayApp()
        {
            InitializeComponent();
            InitializeClipboardMonitoring();
            SetPiperPath();
            SetLogFilePath();
            var (model, speed, logging) = ReadSettings();
            isLoggingEnabled = logging;
            CheckVoiceFiles();
            Log("Application started");
            LogAudioDevices();
            LogDefaultAudioDevice();
            StartMonitoring();

            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_ALT, VK_Q);
        }

        private void InitializeComponent()
        {
            string iconPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "icon.ico");
            trayIcon = new NotifyIcon()
            {
                Icon = new Icon(iconPath),
                ContextMenuStrip = new ContextMenuStrip(),
                Visible = true,
                Text = "Piper Tray"
            };

            toggleMonitoringMenuItem = new ToolStripMenuItem("Monitoring", null, ToggleMonitoring)
            {
                Checked = true
            };
            stopPlaybackMenuItem = new ToolStripMenuItem("Stop Speech", null, StopPlayback);
            exitMenuItem = new ToolStripMenuItem("Exit", null, Exit);

            trayIcon.ContextMenuStrip.Items.Add(toggleMonitoringMenuItem);
            trayIcon.ContextMenuStrip.Items.Add(stopPlaybackMenuItem);
            trayIcon.ContextMenuStrip.Items.Add(exitMenuItem);

            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
            this.WindowState = FormWindowState.Minimized;
        }

        private void CheckVoiceFiles()
        {
            string appDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string[] onnxFiles = Directory.GetFiles(appDirectory, "*.onnx");

            if (onnxFiles.Length == 0)
            {
                MessageBox.Show("No voice files (.onnx) detected in the application directory. The application may not function correctly.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312 && m.WParam.ToInt32() == HOTKEY_ID)
            {
                StopPlayback(this, EventArgs.Empty);
            }
            base.WndProc(ref m);
        }

        private void InitializeClipboardMonitoring()
        {
            clipboardTimer = new System.Windows.Forms.Timer();
            clipboardTimer.Interval = 1000; // Check every second
            clipboardTimer.Tick += ClipboardTimer_Tick;
        }

        private void SetPiperPath()
        {
            string assemblyLocation = Assembly.GetExecutingAssembly().Location;
            string assemblyDirectory = Path.GetDirectoryName(assemblyLocation);
            piperPath = Path.Combine(assemblyDirectory, "piper.exe");

            if (!File.Exists(piperPath))
            {
                Log($"Error: piper.exe not found at {piperPath}");
                MessageBox.Show("piper.exe not found in the application directory.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            else
            {
                Log($"Piper executable found at {piperPath}");
            }
        }

        private void SetLogFilePath()
        {
            string assemblyLocation = Assembly.GetExecutingAssembly().Location;
            string assemblyDirectory = Path.GetDirectoryName(assemblyLocation);
            logFilePath = Path.Combine(assemblyDirectory, "system.log");
            Log($"Log file initialized at {logFilePath}");
        }

        private void Log(string message)
        {
            if (!isLoggingEnabled) return;

            try
            {
                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
                File.AppendAllText(logFilePath, logMessage + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing to log file: {ex.Message}");
            }
        }

        private void LogAudioDevices()
        {
            Log("Available audio devices:");
            for (int i = 0; i < WaveOut.DeviceCount; i++)
            {
                var capabilities = WaveOut.GetCapabilities(i);
                Log($"Device {i}: {capabilities.ProductName}");
            }
        }

        private void LogDefaultAudioDevice()
        {
            try
            {
                var defaultDevice = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                Log($"Default audio device: {defaultDevice.FriendlyName}");
            }
            catch (Exception ex)
            {
                Log($"Error getting default audio device: {ex.Message}");
            }
        }

        private void StartMonitoring()
        {
            isMonitoring = true;
            toggleMonitoringMenuItem.Checked = true;
            toggleMonitoringMenuItem.Text = "Monitoring";
            clipboardTimer.Start();
            Log("Clipboard monitoring started");
            trayIcon.ShowBalloonTip(3000, "Piper Tray", "Clipboard monitoring started", ToolTipIcon.Info);
        }

        private void StopMonitoring()
        {
            isMonitoring = false;
            toggleMonitoringMenuItem.Checked = false;
            toggleMonitoringMenuItem.Text = "Monitoring";
            clipboardTimer.Stop();
            Log("Clipboard monitoring stopped");
            trayIcon.ShowBalloonTip(3000, "Piper Tray", "Clipboard monitoring stopped", ToolTipIcon.Info);
        }

        private void ToggleMonitoring(object sender, EventArgs e)
        {
            if (isMonitoring)
            {
                StopMonitoring();
            }
            else
            {
                StartMonitoring();
            }
        }

        private void StopPlayback(object sender, EventArgs e)
        {
            if (currentWaveOut != null && currentWaveOut.PlaybackState == PlaybackState.Playing)
            {
                currentWaveOut.Stop();
                Log("Audio playback stopped by user");
            }
        }

        private async void ClipboardTimer_Tick(object sender, EventArgs e)
        {
            if (isProcessing)
            {
                return; // Skip this tick if we're still processing the previous text
            }

            if (Clipboard.ContainsText())
            {
                string clipboardContent = Clipboard.GetText();
                if (clipboardContent != lastClipboardContent)
                {
                    lastClipboardContent = clipboardContent;
                    Log($"New clipboard content detected: {clipboardContent.Substring(0, Math.Min(50, clipboardContent.Length))}...");
                    isProcessing = true;
                    await ConvertAndPlayTextToSpeechAsync(clipboardContent);
                    isProcessing = false;
                }
            }
        }

        private (string model, float speed, bool logging) ReadSettings()
        {
            string settingsPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "settings.conf");
            string model = "en_US-libritts_r-medium.onnx";
            float speed = 1.0f;
            bool logging = false;

            if (File.Exists(settingsPath))
            {
                foreach (string line in File.ReadAllLines(settingsPath))
                {
                    if (line.StartsWith("Model="))
                    {
                        model = line.Substring(6).Trim();
                    }
                    else if (line.StartsWith("Speed="))
                    {
                        if (float.TryParse(line.Substring(6).Trim(), out float parsedSpeed))
                        {
                            speed = parsedSpeed;
                        }
                    }
                    else if (line.StartsWith("Logging="))
                    {
                        logging = line.Substring(8).Trim().Equals("true", StringComparison.OrdinalIgnoreCase);
                    }
                }
            }
            else
            {
                Log("settings.conf not found. Using default values.");
            }

            return (model, speed, logging);
        }

        private (HashSet<string> ignoreWords, HashSet<string> bannedWords, Dictionary<string, string> replaceWords) LoadDictionaries()
        {
            var ignoreWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var bannedWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var replaceWords = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            string baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            if (File.Exists(Path.Combine(baseDir, "ignore.dict")))
                ignoreWords = new HashSet<string>(File.ReadAllLines(Path.Combine(baseDir, "ignore.dict")), StringComparer.OrdinalIgnoreCase);

            if (File.Exists(Path.Combine(baseDir, "banned.dict")))
                bannedWords = new HashSet<string>(File.ReadAllLines(Path.Combine(baseDir, "banned.dict")), StringComparer.OrdinalIgnoreCase);

            if (File.Exists(Path.Combine(baseDir, "replace.dict")))
            {
                foreach (var line in File.ReadAllLines(Path.Combine(baseDir, "replace.dict")))
                {
                    var parts = line.Split(new[] { '=' }, 2);
                    if (parts.Length == 2)
                        replaceWords[parts[0].Trim()] = parts[1].Trim();
                }
            }

            return (ignoreWords, bannedWords, replaceWords);
        }

        private async Task ConvertAndPlayTextToSpeechAsync(string text)
        {
            var (ignoreWords, bannedWords, replaceWords) = LoadDictionaries();

            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var processedLines = new List<string>();

            foreach (var line in lines)
            {
                if (bannedWords.Any(word => line.IndexOf(word, StringComparison.OrdinalIgnoreCase) >= 0))
                    continue;

                var words = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                var processedWords = new List<string>();

                foreach (var word in words)
                {
                    if (ignoreWords.Contains(word))
                        continue;

                    var processedWord = word;
                    foreach (var replace in replaceWords)
                    {
                        if (processedWord.IndexOf(replace.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                            processedWord = processedWord.Replace(replace.Key, replace.Value, StringComparison.OrdinalIgnoreCase);
                    }
                    processedWords.Add(processedWord);
                }

                processedLines.Add(string.Join(" ", processedWords));
            }

            var processedText = string.Join("\n", processedLines);

            var (model, speed, _) = ReadSettings();

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = piperPath,
                Arguments = $"--model {model} --output-raw --length-scale {speed}",
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(piperPath)
            };

            try
            {
                using (Process process = new Process())
                {
                    process.StartInfo = psi;
                    process.Start();
                    Log($"Piper process started with arguments: {psi.Arguments}");

                    process.ErrorDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                        {
                            Log($"Piper Error: {e.Data}");
                        }
                    };
                    process.BeginErrorReadLine();

                    await process.StandardInput.WriteLineAsync(processedText);
                    process.StandardInput.Close();

                    using (var memoryStream = new MemoryStream())
                    {
                        await process.StandardOutput.BaseStream.CopyToAsync(memoryStream);
                        memoryStream.Position = 0;

                        if (memoryStream.Length == 0)
                        {
                            Log("Error: No audio data received from Piper");
                            return;
                        }

                        Log($"Received {memoryStream.Length} bytes of audio data");

                        await PlayAudioWithWaveOutEvent(memoryStream);
                    }

                    if (!process.HasExited)
                    {
                        process.Kill();
                        Log("Piper process terminated after audio playback");
                    }

                    await Task.Run(() => process.WaitForExit());
                    Log($"Piper process exited with code: {process.ExitCode}");
                }
            }
            catch (Exception ex)
            {
                Log($"Error in ConvertAndPlayTextToSpeech: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task PlayAudioWithWaveOutEvent(MemoryStream audioStream)
        {
            audioStream.Position = 0;
            using (var rawStream = new RawSourceWaveStream(audioStream, new WaveFormat(22050, 16, 1)))
            using (currentWaveOut = new WaveOutEvent())
            {
                currentWaveOut.Init(rawStream);
                Log("WaveOutEvent initialized");

                currentWaveOut.Play();
                Log("WaveOutEvent playback started");

                var playbackStopwatch = Stopwatch.StartNew();
                while (currentWaveOut.PlaybackState == PlaybackState.Playing)
                {
                    await Task.Delay(100);
                }
                playbackStopwatch.Stop();

                Log($"WaveOutEvent playback completed. Duration: {playbackStopwatch.ElapsedMilliseconds}ms");
            }
            currentWaveOut = null;
        }

        private void Exit(object sender, EventArgs e)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            clipboardTimer.Stop();
            trayIcon.Visible = false;
            Log("Application exiting");
            Application.Exit();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                UnregisterHotKey(this.Handle, HOTKEY_ID);
                trayIcon?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
