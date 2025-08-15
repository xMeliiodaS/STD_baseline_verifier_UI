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
                StatusText.Text = "File selected successfully.";
            }
        }

        private void RunScript_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(selectedFilePath) || string.IsNullOrWhiteSpace(STDNameInput.Text))
            {
                StatusText.Text = "Please select a file and enter STD name.";
                return;
            }

            try
            {
                // --- 1. Prepare per-user AppData folder for JSON ---
                string appDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AT_baseline_verifier"
                );
                Directory.CreateDirectory(appDataFolder);

                string userConfigPath = Path.Combine(appDataFolder, "config.json");
                string defaultConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Scripts", "config.json");

                // --- 2. Copy default JSON if missing ---
                if (!File.Exists(userConfigPath))
                {
                    if (File.Exists(defaultConfigPath))
                        File.Copy(defaultConfigPath, userConfigPath);
                    else
                    {
                        StatusText.Text = $"Default config.json not found at {defaultConfigPath}";
                        return;
                    }
                }

                // --- 3. Update std_name ---
                var json = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(userConfigPath));
                json["std_name"] = STDNameInput.Text;
                File.WriteAllText(userConfigPath, json.ToString());
                StatusText.Text = $"STD name updated to: {json["std_name"]}";

                // --- 4. Run Python EXE ---
                string exePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Scripts", "open_azure_vsts_test.exe");
                if (!File.Exists(exePath))
                {
                    StatusText.Text = $"Python EXE not found at {exePath}";
                    return;
                }

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = $"\"{selectedFilePath}\" \"{STDNameInput.Text}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string errors = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    StatusText.Text = !string.IsNullOrEmpty(errors) ? $"Error: {errors}" : $"Script executed successfully.\n{output}";
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
