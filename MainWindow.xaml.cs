using Microsoft.Win32;
using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace AT_baseline_verifier
{
    public partial class MainWindow : Window
    {
        private string selectedFilePath;

        public MainWindow()
        {
            InitializeComponent();
        }

        private string EnsureUserConfigExists()
        {
            string appDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "AT_baseline_verifier"
            );
            Directory.CreateDirectory(appDataFolder);

            string userConfigPath = Path.Combine(appDataFolder, "config.json");
            string defaultConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");

            if (!File.Exists(userConfigPath))
            {
                if (File.Exists(defaultConfigPath))
                {
                    File.Copy(defaultConfigPath, userConfigPath);
                }
                else
                {
                    throw new FileNotFoundException($"Default config.json not found at {defaultConfigPath}");
                }
            }

            return userConfigPath;
        }

        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                selectedFilePath = openFileDialog.FileName;
                SelectedFileLabel.Text = selectedFilePath;

                try
                {
                    string userConfigPath = EnsureUserConfigExists();

                    var json = JObject.Parse(File.ReadAllText(userConfigPath));
                    json["excel_path"] = selectedFilePath;
                    File.WriteAllText(userConfigPath, json.ToString());

                    StatusText.Text = $"Excel path updated to: {json["excel_path"]}";
                }
                catch (Exception ex)
                {
                    StatusText.Text = $"Error updating config: {ex.Message}";
                }
            }
        }

        private async void RunScript_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(selectedFilePath) || string.IsNullOrWhiteSpace(STDNameInput.Text))
            {
                StatusText.Text = "Please select a file and enter STD name.";
                return;
            }

            try
            {
                string userConfigPath = EnsureUserConfigExists();

                var json = JObject.Parse(File.ReadAllText(userConfigPath));
                json["std_name"] = STDNameInput.Text;
                File.WriteAllText(userConfigPath, json.ToString());

                string pythonExePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test_open_azure_vsts.exe");

                if (!File.Exists(pythonExePath))
                {
                    StatusText.Text = $"Python EXE not found at {pythonExePath}";
                    return;
                }

                StatusText.Text = "Running verification...";

                var psi = new ProcessStartInfo
                {
                    FileName = pythonExePath,
                    Arguments = $"\"{selectedFilePath}\" \"{STDNameInput.Text}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                var outputBuilder = new System.Text.StringBuilder();
                var errorBuilder = new System.Text.StringBuilder();

                await Task.Run(() =>
                {
                    using (var process = new Process())
                    {
                        process.StartInfo = psi;

                        process.OutputDataReceived += (s, ea) =>
                        {
                            if (!string.IsNullOrEmpty(ea.Data))
                                outputBuilder.AppendLine(ea.Data);
                        };

                        process.ErrorDataReceived += (s, ea) =>
                        {
                            if (!string.IsNullOrEmpty(ea.Data))
                                errorBuilder.AppendLine(ea.Data);
                        };

                        process.Start();
                        process.BeginOutputReadLine();
                        process.BeginErrorReadLine();
                        process.WaitForExit();
                    }
                });

                // ✅ Now update UI safely (only here)
                if (errorBuilder.Length > 0)
                {
                    StatusText.Text = $"Error:\n{errorBuilder}";
                }
                else
                {
                    StatusText.Text = $"Script executed successfully:\n{outputBuilder}";
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Execution failed: {ex.Message}";
            }
        }



        private void STDNameInput_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            STDPlaceholder.Visibility = string.IsNullOrEmpty(STDNameInput.Text)
                ? Visibility.Visible
                : Visibility.Hidden;
        }

        // Drag & Drop Handlers
        private void Border_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && (files[0].EndsWith(".xls", StringComparison.OrdinalIgnoreCase) || files[0].EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
                {
                    e.Effects = DragDropEffects.Copy;
                }
                else
                {
                    e.Effects = DragDropEffects.None;
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void Border_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && (files[0].EndsWith(".xls", StringComparison.OrdinalIgnoreCase) || files[0].EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
                {
                    selectedFilePath = files[0];
                    SelectedFileLabel.Text = selectedFilePath;
                    StatusText.Text = "File dropped successfully.";
                }
                else
                {
                    StatusText.Text = "Only Excel files (.xls, .xlsx) are supported.";
                }
            }
        }
        private void Window_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && (files[0].EndsWith(".xls") || files[0].EndsWith(".xlsx")))
                {
                    e.Effects = DragDropEffects.Copy;
                }
                else
                {
                    e.Effects = DragDropEffects.None;
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length > 0 && (files[0].EndsWith(".xls") || files[0].EndsWith(".xlsx")))
            {
                selectedFilePath = files[0];
                SelectedFileLabel.Text = selectedFilePath;
                StatusText.Text = "File selected via drag & drop.";
            }
            else
            {
                StatusText.Text = "Invalid file format. Please drop an Excel file.";
            }
        }
    }
}
}
