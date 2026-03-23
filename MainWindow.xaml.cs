using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Input;
using Mammoth;
using Microsoft.Win32;

namespace WordViewer;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    // ── State ────────────────────────────────────────────────────────────────

    private string? _currentFilePath;
    private string? _tempHtmlFile;
    private double  _zoom = 1.0;
    private bool    _documentLoaded;
    private bool    _loading;
    private string  _statusText   = "Ready";
    private string  _warningText  = string.Empty;

    // ── Constructor ──────────────────────────────────────────────────────────

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;
        BuildCommands();
        Loaded += OnLoaded;
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        await WebView.EnsureCoreWebView2Async();

        WebView.AllowDrop = true;

        // Hook NavigationCompleted so we can apply zoom reliably after each load
        WebView.CoreWebView2.NavigationCompleted += OnNavigationCompleted;

        // Open a file passed via command-line argument
        var args = Environment.GetCommandLineArgs();
        if (args.Length > 1 && File.Exists(args[1]))
            await OpenFileAsync(args[1]);
    }

    private async void OnNavigationCompleted(object? sender,
        Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e)
    {
        if (_documentLoaded)
            await ApplyZoomAsync(_zoom);
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        DeleteTempFile();
    }

    // ── Commands ─────────────────────────────────────────────────────────────

    public ICommand OpenCommand           { get; private set; } = null!;
    public ICommand ReloadCommand         { get; private set; } = null!;
    public ICommand CloseDocumentCommand  { get; private set; } = null!;
    public ICommand ZoomInCommand         { get; private set; } = null!;
    public ICommand ZoomOutCommand        { get; private set; } = null!;
    public ICommand ZoomResetCommand      { get; private set; } = null!;

    private void BuildCommands()
    {
        OpenCommand          = new RelayCommand(_ => PickFile());
        ReloadCommand        = new RelayCommand(_ => _ = ReloadAsync(),       _ => _currentFilePath != null);
        CloseDocumentCommand = new RelayCommand(_ => CloseDocument(),         _ => _documentLoaded);
        ZoomInCommand        = new RelayCommand(_ => _ = ChangeZoomAsync(+0.1), _ => _documentLoaded);
        ZoomOutCommand       = new RelayCommand(_ => _ = ChangeZoomAsync(-0.1), _ => _documentLoaded);
        ZoomResetCommand     = new RelayCommand(_ => _ = ChangeZoomAsync(0),    _ => _documentLoaded);
    }

    // ── File picking ─────────────────────────────────────────────────────────

    private void PickFile()
    {
        var dlg = new OpenFileDialog
        {
            Title  = "Open Word Document",
            Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*",
        };

        if (_currentFilePath != null)
            dlg.InitialDirectory = Path.GetDirectoryName(_currentFilePath);

        if (dlg.ShowDialog() == true)
            _ = OpenFileAsync(dlg.FileName);
    }

    // ── Drag & drop ──────────────────────────────────────────────────────────

    private void OnDragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void OnDrop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
            _ = OpenFileAsync(files[0]);
    }

    // ── Core open / reload / close ───────────────────────────────────────────

    private async Task OpenFileAsync(string path)
    {
        if (!File.Exists(path))
        {
            MessageBox.Show($"File not found:\n{path}", "Word Viewer", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _currentFilePath = path;
        Title = $"{Path.GetFileName(path)} — Word Viewer";
        await RenderDocumentAsync();
    }

    private async Task ReloadAsync()
    {
        if (_currentFilePath != null)
            await RenderDocumentAsync();
    }

    private void CloseDocument()
    {
        _currentFilePath = null;
        _documentLoaded  = false;
        Title            = "Word Viewer";
        StatusText       = "Ready";
        WarningText      = string.Empty;
        WebView.CoreWebView2.Navigate("about:blank");
        DeleteTempFile();
        OnPropertyChanged(nameof(WelcomeVisibility));
        OnPropertyChanged(nameof(DocumentVisibility));
        CommandManager.InvalidateRequerySuggested();
    }

    private async Task RenderDocumentAsync()
    {
        Loading = true;
        WarningText = string.Empty;
        StatusText = $"Loading {Path.GetFileName(_currentFilePath)}…";

        try
        {
            var (html, warnings) = await Task.Run(() => ConvertDocument(_currentFilePath!));

            var page = BuildHtmlPage(html);
            await NavigateToHtmlAsync(page);

            _documentLoaded = true;
            StatusText  = Path.GetFileName(_currentFilePath)!;
            WarningText = warnings;

            OnPropertyChanged(nameof(WelcomeVisibility));
            OnPropertyChanged(nameof(DocumentVisibility));
            CommandManager.InvalidateRequerySuggested();
        }
        catch (Exception ex)
        {
            StatusText = "Failed to open document.";
            var msg = BuildExceptionMessage(ex);
            MessageBox.Show(msg, "Word Viewer — Conversion Error",
                            MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            Loading = false;
        }
    }

    // ── Mammoth conversion (runs on thread pool) ─────────────────────────────

    private static (string html, string warnings) ConvertDocument(string path)
    {
        using var stream  = DocxPreprocessor.Preprocess(path);
        var converter     = new DocumentConverter();
        var result        = converter.ConvertToHtml(stream);
        var warnings      = result.Warnings.Count > 0
            ? $"{result.Warnings.Count} conversion warning(s)"
            : string.Empty;
        return (result.Value, warnings);
    }

    // ── Navigation helpers ────────────────────────────────────────────────────

    /// <summary>
    /// NavigateToString has a ~2 MB limit (fails on large documents with embedded images).
    /// Writing to a temp file and using Navigate() has no such limit.
    /// </summary>
    private async Task NavigateToHtmlAsync(string html)
    {
        // Reuse a single temp file for the lifetime of the app
        _tempHtmlFile ??= Path.Combine(Path.GetTempPath(), $"WordViewer_{Environment.ProcessId}.html");
        await File.WriteAllTextAsync(_tempHtmlFile, html, System.Text.Encoding.UTF8);
        WebView.CoreWebView2.Navigate("file:///" + _tempHtmlFile.Replace('\\', '/'));
    }

    private void DeleteTempFile()
    {
        if (_tempHtmlFile == null) return;
        try { File.Delete(_tempHtmlFile); } catch { }
        _tempHtmlFile = null;
    }

    // ── HTML page builder ─────────────────────────────────────────────────────

    private static string BuildHtmlPage(string body)
    {
        return
            """
            <!DOCTYPE html>
            <html lang="en">
            <head>
            <meta charset="UTF-8" />
            <meta name="viewport" content="width=device-width, initial-scale=1" />
            <style>

            *, *::before, *::after { box-sizing: border-box; }

            html, body {
                margin: 0;
                padding: 0;
                background: #404040;
                font-family: 'Calibri', 'Segoe UI', Arial, sans-serif;
                font-size: 11pt;
                color: #111;
            }

            /* Scrollbar */
            ::-webkit-scrollbar              { width: 10px; height: 10px; }
            ::-webkit-scrollbar-track        { background: #404040; }
            ::-webkit-scrollbar-thumb        { background: #777; border-radius: 5px; }
            ::-webkit-scrollbar-thumb:hover  { background: #999; }

            /* Paper */
            .document-page {
                width: 210mm;
                min-height: 297mm;
                margin: 24px auto;
                background: #fff;
                padding: 25.4mm 25.4mm 25.4mm 25.4mm; /* 1-inch margins */
                box-shadow: 0 4px 20px rgba(0,0,0,.45);
                border-radius: 2px;
            }

            /* Typography */
            p  { margin: 0 0 .6em; line-height: 1.5; }
            h1 { font-size: 20pt; margin: .8em 0 .4em; }
            h2 { font-size: 16pt; margin: .7em 0 .35em; }
            h3 { font-size: 13pt; margin: .6em 0 .3em; }
            h4, h5, h6 { font-size: 11pt; margin: .5em 0 .25em; }

            /* Tables */
            table {
                border-collapse: collapse;
                width: 100%;
                margin: .8em 0;
                font-size: 10pt;
            }
            th, td {
                border: 1px solid #c0c0c0;
                padding: 5px 8px;
                vertical-align: top;
            }
            th {
                background: #f0f0f0;
                font-weight: 600;
            }
            tr:nth-child(even) td { background: #fafafa; }

            /* Lists */
            ul, ol { margin: 0 0 .6em 1.5em; padding: 0; }
            li     { margin-bottom: .2em; line-height: 1.5; }

            /* Images */
            img { max-width: 100%; height: auto; display: block; margin: .6em 0; }

            /* Inline code / pre */
            code { font-family: Consolas, monospace; font-size: 9pt; background: #f4f4f4; padding: 1px 4px; border-radius: 3px; }
            pre  { background: #f4f4f4; padding: 10px 14px; border-radius: 4px; overflow-x: auto; font-size: 9pt; line-height: 1.4; margin: .6em 0; }

            /* Hyperlinks */
            a       { color: #0078d4; text-decoration: none; }
            a:hover { text-decoration: underline; }

            /* Blockquote / callout */
            blockquote {
                margin: .6em 0 .6em 1.2em;
                padding: .4em .8em;
                border-left: 3px solid #0078d4;
                color: #444;
                font-style: italic;
            }

            /* HR */
            hr { border: none; border-top: 1px solid #ddd; margin: 1.2em 0; }

            /* Strong / em */
            strong { font-weight: 700; }
            em     { font-style: italic; }

            /* Underline (Mammoth preserves as <u>) */
            u { text-decoration: underline; }

            </style>
            </head>
            <body>
            <div class="document-page">
            """ +
            body +
            """

            </div>
            </body>
            </html>
            """;
    }

    // ── Zoom ─────────────────────────────────────────────────────────────────

    private async Task ChangeZoomAsync(double delta)
    {
        if (delta == 0)
            _zoom = 1.0;
        else
            _zoom = Math.Clamp(_zoom + delta, 0.25, 4.0);

        OnPropertyChanged(nameof(ZoomText));

        if (_documentLoaded)
            await ApplyZoomAsync(_zoom);
    }

    private async Task ApplyZoomAsync(double zoom)
    {
        try
        {
            await WebView.ExecuteScriptAsync(
                $"document.body.style.zoom='{zoom.ToString(System.Globalization.CultureInfo.InvariantCulture)}';");
        }
        catch { /* WebView may not be ready */ }
    }

    // ── Bindable properties ───────────────────────────────────────────────────

    public string ZoomText => $"{(int)(_zoom * 100)}%";

    public Visibility WelcomeVisibility  => _documentLoaded ? Visibility.Collapsed : Visibility.Visible;
    public Visibility DocumentVisibility => _documentLoaded ? Visibility.Visible   : Visibility.Collapsed;
    public Visibility LoadingVisibility  => _loading        ? Visibility.Visible   : Visibility.Collapsed;
    public Visibility WarningVisibility  =>
        string.IsNullOrEmpty(_warningText) ? Visibility.Collapsed : Visibility.Visible;

    public string StatusText
    {
        get => _statusText;
        set { _statusText = value; OnPropertyChanged(); }
    }

    public string WarningText
    {
        get => _warningText;
        set { _warningText = value; OnPropertyChanged(); OnPropertyChanged(nameof(WarningVisibility)); }
    }

    public bool Loading
    {
        get => _loading;
        set { _loading = value; OnPropertyChanged(); OnPropertyChanged(nameof(LoadingVisibility)); }
    }

    // ── Error formatting ──────────────────────────────────────────────────────

    private static string BuildExceptionMessage(Exception ex)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("Could not open the document.");
        sb.AppendLine();

        var current = ex;
        while (current != null)
        {
            sb.AppendLine($"[{current.GetType().Name}]  {current.Message}");
            if (current.StackTrace != null)
            {
                foreach (var line in current.StackTrace.Split('\n').Take(6))
                    sb.AppendLine("  " + line.TrimEnd());
            }
            if (current.InnerException != null)
                sb.AppendLine("  --- inner exception ---");
            current = current.InnerException;
        }

        // Also write to a log file for easy sharing
        try
        {
            var logPath = Path.Combine(Path.GetTempPath(), "WordViewer_error.txt");
            File.WriteAllText(logPath, sb.ToString());
            sb.AppendLine();
            sb.AppendLine($"Full details saved to: {logPath}");
        }
        catch { /* best effort */ }

        return sb.ToString().TrimEnd();
    }

    // ── INotifyPropertyChanged ────────────────────────────────────────────────

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

// ── Simple relay command ───────────────────────────────────────────────────────

internal sealed class RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null) : ICommand
{
    public bool CanExecute(object? parameter) => canExecute?.Invoke(parameter) ?? true;
    public void Execute(object? parameter)    => execute(parameter);
    public event EventHandler? CanExecuteChanged
    {
        add    => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }
}
