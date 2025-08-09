using Microsoft.Win32;
using System;
using System.Diagnostics;
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
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "python",
                    Arguments = $"your_script.py \"{selectedFilePath}\" \"{STDNameInput.Text}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string errors = process.StandardError.ReadToEnd();

                    if (!string.IsNullOrEmpty(errors))
                    {
                        StatusText.Text = $"Error: {errors}";
                    }
                    else
                    {
                        StatusText.Text = "Script executed successfully.";
                    }
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
