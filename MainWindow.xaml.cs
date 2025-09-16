using Microsoft.Win32;
using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using static System.Runtime.InteropServices.JavaScript.JSType;
using IOPath = System.IO.Path;

namespace AT_baseline_verifier
{
    public partial class MainWindow : Window
    {
        private string selectedFilePath;
        private Storyboard automationSpinnerStoryboard;
        private Storyboard violationSpinnerStoryboard;
        private readonly string logFilePath = IOPath.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.log");

        private readonly string configJsonFileName = "config.json";
        private readonly string configSTDNameKey = "std_name";
        private readonly string configIterationPathKey = "iteration_path";
        private readonly string configCurrentVVKey = "current_version";


        public MainWindow()
        {
            InitializeComponent();
        }


        // Ensures the user-specific configuration file exists by copying from the default if missing.
        private string EnsureUserConfigExists()
        {
            string appDataFolder = IOPath.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "AT_baseline_verifier"
            );

            Directory.CreateDirectory(appDataFolder);

            string userConfigPath = IOPath.Combine(appDataFolder, configJsonFileName);
            string defaultConfigPath = IOPath.Combine(AppDomain.CurrentDomain.BaseDirectory, configJsonFileName);

            try
            {
                if (!File.Exists(userConfigPath))
                {
                    if (File.Exists(defaultConfigPath))
                    {
                        File.Copy(defaultConfigPath, userConfigPath);
                    }
                    else
                    {
                        throw new FileNotFoundException($"Default {configJsonFileName} not found at {defaultConfigPath}");
                    }
                }

                return userConfigPath;
            }
            catch (Exception ex)
            {
                throw new IOException($"Failed to ensure user config exists: {ex.Message}", ex);
            }
        }


        // Executes an external process with the given executable and arguments, handling logging, validation, and UI state.
        private async Task RunExternalProcess(string exeName, string successMessage, bool useConfig = true)
        {
            File.AppendAllText(logFilePath, "======================================================");

            if (string.IsNullOrWhiteSpace(selectedFilePath))
            {
                SetResultStatus("Please Select an Excel File.", true);
                return;
            }

            if (useConfig)
            {
                if (string.IsNullOrWhiteSpace(STDNameInput.Text) ||
                    string.IsNullOrWhiteSpace(IterationPathInput.Text) ||
                    string.IsNullOrWhiteSpace(VVVersionInput.Text))
                {
                    SetResultStatus("Please Fill all fields.", true);
                    return;
                }
            }

            try
            {
                if (useConfig)
                {
                    TrimInputs();
                    ReplaceCommasWithDots();

                    string userConfigPath = EnsureUserConfigExists();
                    var json = JObject.Parse(File.ReadAllText(userConfigPath));
                    json[configSTDNameKey] = STDNameInput.Text;
                    json[configIterationPathKey] = IterationPathInput.Text;
                    json[configCurrentVVKey] = VVVersionInput.Text;
                    File.WriteAllText(userConfigPath, json.ToString());
                }

                string exePath = IOPath.Combine(AppDomain.CurrentDomain.BaseDirectory, exeName);
                if (!File.Exists(exePath))
                {
                    SetResultStatus($"Error: {exeName} not found.", true);
                    LogError($"{exeName} not found at {exePath}");
                    return;
                }

                SetActionButtonsEnabled(false);

                SetResultStatus("Running...", false);

                var psi = new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = $"\"{selectedFilePath}\" \"{STDNameInput.Text}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                var outputBuilder = new StringBuilder();
                var errorBuilder = new StringBuilder();

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

                if (errorBuilder.Length > 0)
                {
                    LogError(errorBuilder.ToString());
                    SetResultStatus("Execution failed. See log for details.", true);
                }
                else
                {
                    SetResultStatus(successMessage, false);
                }
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                SetResultStatus("Execution failed. See log for details.", true);
            }
            finally
            {
                SetActionButtonsEnabled(true);
                File.AppendAllText(logFilePath, "======================================================");
            }
        }


        // Runs the automation validation executable and updates the UI accordingly.
        private async void RunAutomation_Click(object sender, RoutedEventArgs e)
        {
            StartButtonSpinner(RunAutomationButton, RunIcon, RunButtonText, ButtonSpinner, ButtonSpinnerRotate, ref automationSpinnerStoryboard);
            await RunExternalProcess("test_bugs_std_validation.exe", "Automation completed successfully!", true);
            StopButtonSpinner(RunAutomationButton, RunIcon, RunButtonText, ButtonSpinner, ref automationSpinnerStoryboard);
        }


        // Runs the violation check executable and updates the UI accordingly.
        private async void RunViolationCheck_Click(object sender, RoutedEventArgs e)
        {
            StartButtonSpinner(RunViolationButton, ViolationIcon, RunViolationButtonText, ViolationButtonSpinner, ViolationButtonSpinnerRotate, ref violationSpinnerStoryboard);
            await RunExternalProcess("test_excel_violations.exe", "Violation check completed successfully!", false);
            StopButtonSpinner(RunViolationButton, ViolationIcon, RunViolationButtonText, ViolationButtonSpinner, ref violationSpinnerStoryboard);
        }


        // Toggles placeholder visibility for STD Name input field based on text changes.
        private void STDNameInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            STDPlaceholder.Visibility = string.IsNullOrEmpty(STDNameInput.Text)
                ? Visibility.Visible
                : Visibility.Hidden;
        }


        // Toggles placeholder visibility for Iteration Path input field based on text changes.
        private void IterationPathInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            IterationPathPlaceholder.Visibility = string.IsNullOrEmpty(IterationPathInput.Text)
                ? Visibility.Visible
                : Visibility.Hidden;
        }


        // Toggles placeholder visibility for Current Version input field based on text changes.
        private void VVVersionInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            VVVersionPlaceholder.Visibility = string.IsNullOrEmpty(VVVersionInput.Text)
                ? Visibility.Visible
                : Visibility.Hidden;
        }

        // Sets the selected Excel file path, updates config.json, and refreshes the status label.
        private void HandleExcelFileSelection(string filePath)
        {
            selectedFilePath = filePath;

            // Get the base file name without extension for the input
            string fileNameWithoutExt = IOPath.GetFileNameWithoutExtension(selectedFilePath);

            // Update label with actual file name (including correct extension)
            SelectedFileLabel.Text = IOPath.GetFileName(selectedFilePath);

            try
            {
                // Ensure user config exists and update Excel path
                string userConfigPath = EnsureUserConfigExists();
                var json = JObject.Parse(File.ReadAllText(userConfigPath));
                json["excel_path"] = selectedFilePath;
                File.WriteAllText(userConfigPath, json.ToString());

                // Set STD Name input if empty, focus it, and select all text
                if (string.IsNullOrWhiteSpace(STDNameInput.Text))
                {
                    STDNameInput.Text = fileNameWithoutExt;
                }
                STDNameInput.Focus();
                Dispatcher.BeginInvoke(new Action(() => STDNameInput.SelectAll()));

                // Update status
                SetResultStatus($"Excel path updated to: {json["excel_path"]}", false);
            }
            catch (Exception ex)
            {
                SetResultStatus($"Error updating config: {ex.Message}", true);
            }
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
                HandleExcelFileSelection(openFileDialog.FileName);
            }
        }


        // Handles drag-over events to validate whether the dragged file is an Excel file.
        private void Window_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && (files[0].EndsWith(".xls") || files[0].EndsWith(".xlsx")))
                {
                    e.Effects = DragDropEffects.Copy;
                }
                else e.Effects = DragDropEffects.None;
            }
            else e.Effects = DragDropEffects.None;

            e.Handled = true;
        }


        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && (files[0].EndsWith(".xls") || files[0].EndsWith(".xlsx")))
                {
                    HandleExcelFileSelection(files[0]);
                }
                else
                {
                    SetResultStatus("Invalid file format. Please drop an Excel file.", false);
                }
            }
        }


        // Starts a spinner animation for the specified button and updates UI state to "running".
        private void StartButtonSpinner(Button targetButton, TextBlock icon, TextBlock buttonText, Ellipse spinner, RotateTransform spinnerRotate, ref Storyboard storyboard)
        {
            spinner.Visibility = Visibility.Visible;
            icon.Visibility = Visibility.Collapsed;

            storyboard = new Storyboard();
            var anim = new DoubleAnimation
            {
                From = 0,
                To = 360,
                Duration = new Duration(TimeSpan.FromSeconds(1)),
                RepeatBehavior = RepeatBehavior.Forever
            };
            Storyboard.SetTarget(anim, spinnerRotate);
            Storyboard.SetTargetProperty(anim, new PropertyPath("Angle"));
            storyboard.Children.Add(anim);
            storyboard.Begin();

            targetButton.IsEnabled = false;
            buttonText.Text = "Running…";
            StatusText.Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255));
        }


        // Stops the spinner animation for the specified button and restores default button/UI state.
        private void StopButtonSpinner(Button targetButton, TextBlock icon, TextBlock buttonText, Ellipse spinner, ref Storyboard storyboard)
        {
            storyboard?.Stop();
            spinner.Visibility = Visibility.Collapsed;
            icon.Visibility = Visibility.Visible;

            if (targetButton == RunAutomationButton)
                buttonText.Text = "Check STD Bugs in VSTS";
            else if (targetButton == RunViolationButton)
                buttonText.Text = "Validate STD Rules";

            targetButton.IsEnabled = true;
        }


        // Enables or disables the main action buttons.
        private void SetActionButtonsEnabled(bool isEnabled)
        {
            RunAutomationButton.IsEnabled = isEnabled;
            RunViolationButton.IsEnabled = isEnabled;
            SelectSTDButton.IsEnabled = isEnabled;

            STDNameInput.IsEnabled = isEnabled;
            IterationPathInput.IsEnabled = isEnabled;
            VVVersionInput.IsEnabled = isEnabled;
        }


        // Updates the status text message with error or success formatting.
        private void SetResultStatus(string message, bool isError = false)
        {
            StatusText.Text = message;
            StatusText.FontWeight = FontWeights.Bold;
            StatusText.Foreground = new SolidColorBrush(isError ? Color.FromRgb(255, 102, 102) : Color.FromRgb(255, 255, 255));
        }


        // Trims whitespace.
        private void TrimInputs()
        {
            // Trim spaces
            STDNameInput.Text = STDNameInput.Text.Trim();
            IterationPathInput.Text = IterationPathInput.Text.Trim();
            VVVersionInput.Text = VVVersionInput.Text.Trim();
        }


        // Replaces commas with dots
        private void ReplaceCommasWithDots()
        {
            // Replace commas with dots only for Current Version field
            VVVersionInput.Text = VVVersionInput.Text.Replace(',', '.');
        }


        // Logs error messages with timestamps to the application log file.
        private void LogError(string error)
        {
            File.AppendAllText(logFilePath, $"\n[ERROR {DateTime.Now}] {error}\n");
        }
    }
}
