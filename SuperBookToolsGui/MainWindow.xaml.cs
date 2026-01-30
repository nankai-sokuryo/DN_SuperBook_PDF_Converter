using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Interop;

using IPA.Cores.Basic;
using IPA.Cores.Helper.Basic;
using static IPA.Cores.Globals.Basic;

using IPA.Cores.Codes;
using IPA.Cores.Helper.Codes;
using static IPA.Cores.Globals.Codes;

using SuperBookTools;

namespace SuperBookToolsGui
{
    public partial class MainWindow : Window
    {
        private CancellationTokenSource? _cts;
        private bool _isRunning;

    public MainWindow()
    {
        InitializeComponent();
        this.Closing += MainWindow_Closing;
        Log("SuperBookTools GUI started.");
        Log($"Application root: {Env.AppRootDir}");
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        // Cancel any running operation
        _cts?.Cancel();
        
        // Shutdown the application properly
        Application.Current.Shutdown();
    }

    private void Log(string message)
    {
        Dispatcher.BeginInvoke(() =>
        {
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
            txtLog.ScrollToEnd();
        });
    }

    private void UpdateProgress(int current, int total, string status)
    {
        Dispatcher.BeginInvoke(() =>
        {
            if (total > 0)
            {
                progressBar.Maximum = total;
                progressBar.Value = current;
                txtProgress.Text = $"{current} / {total} - {status}";
            }
            else
            {
                txtProgress.Text = status;
            }
        });
    }

    private void SetRunning(bool running)
    {
        Dispatcher.BeginInvoke(() =>
        {
            _isRunning = running;
            btnStart.IsEnabled = !running;
            btnCancel.IsEnabled = running;
            txtSourceDir.IsEnabled = !running;
            txtDestDir.IsEnabled = !running;
            chkOcr.IsEnabled = !running;
            
            if (!running)
            {
                progressBar.Value = 0;
            }
        });
    }

    private void BtnBrowseSource_Click(object sender, RoutedEventArgs e)
    {
        string? folder = ShowFolderBrowserDialog("Select Source Directory containing PDF files");
        if (folder != null)
        {
            txtSourceDir.Text = folder;
            
            // Auto-fill destination if empty
            if (string.IsNullOrWhiteSpace(txtDestDir.Text))
            {
                txtDestDir.Text = folder + "_converted";
            }
        }
    }

    private void BtnBrowseDest_Click(object sender, RoutedEventArgs e)
    {
        string? folder = ShowFolderBrowserDialog("Select Output Directory");
        if (folder != null)
        {
            txtDestDir.Text = folder;
        }
    }

    private string? ShowFolderBrowserDialog(string description)
    {
        // Use Shell32 COM dialog via interop
        var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = description,
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true
        };

        var hwnd = new WindowInteropHelper(this).Handle;
        var result = dialog.ShowDialog(new Win32Window(hwnd));
        
        return result == System.Windows.Forms.DialogResult.OK ? dialog.SelectedPath : null;
    }

    // Helper class for WinForms interop
    private class Win32Window : System.Windows.Forms.IWin32Window
    {
        public IntPtr Handle { get; }
        public Win32Window(IntPtr handle) => Handle = handle;
    }

    private async void BtnStart_Click(object sender, RoutedEventArgs e)
    {
        string srcDir = txtSourceDir.Text.Trim();
        string dstDir = txtDestDir.Text.Trim();
        bool performOcr = chkOcr.IsChecked == true;

        // Validation
        if (string.IsNullOrWhiteSpace(srcDir))
        {
            MessageBox.Show("Please select a source directory.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(dstDir))
        {
            MessageBox.Show("Please select an output directory.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!Directory.Exists(srcDir))
        {
            MessageBox.Show("Source directory does not exist.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (srcDir.Equals(dstDir, StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show("Source and output directories must be different.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _cts = new CancellationTokenSource();
        SetRunning(true);

        try
        {
            await RunConversionAsync(srcDir, dstDir, performOcr, _cts.Token);
        }
        catch (OperationCanceledException)
        {
            Log("Operation cancelled by user.");
            MessageBox.Show("Operation cancelled.", "Cancelled", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            Log($"ERROR: {ex.Message}");
            MessageBox.Show($"An error occurred:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetRunning(false);
            _cts?.Dispose();
            _cts = null;
        }
    }

    private async Task RunConversionAsync(string srcDir, string dstDir, bool performOcr, CancellationToken ct)
    {
        srcDir = PP.RemoveLastSeparatorChar(await Lfs.NormalizePathAsync(srcDir, normalizeRelativePathIfSupported: true, cancel: ct));
        dstDir = PP.RemoveLastSeparatorChar(await Lfs.NormalizePathAsync(dstDir, normalizeRelativePathIfSupported: true, cancel: ct));

        Log($"Source: {srcDir}");
        Log($"Output: {dstDir}");
        Log($"OCR: {(performOcr ? "Enabled" : "Disabled")}");

        await Lfs.CreateDirectoryAsync(dstDir, cancel: ct);

        var srcFiles = (await Lfs.EnumDirectoryAsync(srcDir, true, cancel: ct))
            .Where(x => x.IsFile && !x.Name.StartsWith("_") && x.Name._IsExtensionMatch(".pdf"))
            .OrderBy(x => x.FullPath, StrCmpi)
            .ToList();

        int numTotal = srcFiles.Count;
        int numOk = 0;
        int numError = 0;
        int numSkip = 0;

        Log($"Found {numTotal} PDF files.");

        if (numTotal == 0)
        {
            Log("No PDF files found in source directory.");
            MessageBox.Show("No PDF files found in the source directory.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var options = new SuperPerformPdfOptions();
        var errorFiles = new List<string>();

        for (int i = 0; i < srcFiles.Count; i++)
        {
            ct.ThrowIfCancellationRequested();

            var src = srcFiles[i];
            string relativePath = PP.GetRelativeFileName(src.FullPath, srcDir);
            string dstPath = PP.Combine(dstDir, relativePath);

            UpdateProgress(i + 1, numTotal, Path.GetFileName(src.FullPath));
            Log($"[{i + 1}/{numTotal}] Processing: {src.Name}");

            try
            {
                bool result = await SuperPdfUtil.PerformPdfAsync(src.FullPath, dstPath, options, useOkFile: true, cancel: ct);
                
                if (!result)
                {
                    numSkip++;
                    Log($"  -> Skipped (already exists or unchanged)");
                }
                else
                {
                    numOk++;
                    Log($"  -> OK");
                }
            }
            catch (Exception ex)
            {
                numError++;
                errorFiles.Add(src.FullPath);
                Log($"  -> ERROR: {ex.Message}");
            }
        }

        // Perform OCR if enabled
        if (performOcr && numOk > 0)
        {
            ct.ThrowIfCancellationRequested();
            
            Log("");
            Log("Starting Japanese OCR processing (YomiToku AI)...");
            UpdateProgress(0, 0, "Running OCR...");

            try
            {
                await SuperBookExternalTools.YomiToku.PerformOcrDirAsync(dstDir, PP.Combine(dstDir, SuperBookExternalTools.Post_OCR_Dir), SuperBookExternalTools.Post_OCR_Dir, ct);
                Log("OCR processing completed.");
            }
            catch (Exception ex)
            {
                Log($"OCR ERROR: {ex.Message}");
            }
        }

        // Summary
        Log("");
        Log("========== CONVERSION COMPLETE ==========");
        Log($"Total:   {numTotal}");
        Log($"Success: {numOk}");
        Log($"Skipped: {numSkip}");
        Log($"Errors:  {numError}");

        if (errorFiles.Count > 0)
        {
            Log("");
            Log("Failed files:");
            foreach (var f in errorFiles)
            {
                Log($"  - {f}");
            }
        }

        UpdateProgress(numTotal, numTotal, "Complete");

        string message = $"Conversion complete!\n\nTotal: {numTotal}\nSuccess: {numOk}\nSkipped: {numSkip}\nErrors: {numError}";
        MessageBox.Show(message, "Complete", MessageBoxButton.OK, 
            numError > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        if (_cts != null && !_cts.IsCancellationRequested)
        {
            Log("Cancellation requested...");
            _cts.Cancel();
        }
    }

    private void BtnClearLog_Click(object sender, RoutedEventArgs e)
    {
        txtLog.Clear();
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        if (_isRunning)
        {
            var result = MessageBox.Show(
                "Conversion is in progress. Are you sure you want to exit?",
                "Confirm Exit",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.No)
            {
                e.Cancel = true;
                return;
            }

            _cts?.Cancel();
        }

        base.OnClosing(e);
    }
}
}
