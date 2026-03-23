# Word Viewer

A lightweight WPF application for reading Microsoft Word documents (`.docx`) on machines that don't have Word installed.

Documents are converted to HTML in-memory using [Mammoth](https://github.com/mwilliamson/dotnet-mammoth) and rendered in an embedded [WebView2](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) control. No temporary files are written to disk and the original document is never modified.

![Screenshot placeholder](docs/screenshot.png)

---

## Features

- Open `.docx` files via toolbar, `Ctrl+O`, or drag & drop
- Renders text, headings, tables, lists, images, hyperlinks, and inline formatting
- Zoom in/out (`Ctrl+`+`/`-`/`0`) from 25% to 400%
- Reload from disk (`F5`) to pick up external changes
- Pass a file path as a command-line argument to open it directly
- Dark VS-style chrome; document displayed as white A4 paper on a grey background

## Requirements

| Requirement | Version |
|---|---|
| .NET | 8.0 (Windows) |
| WebView2 Runtime | Evergreen (pre-installed on Windows 11; [download](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) for Windows 10) |

## Getting Started

```bash
# Clone / download the repository, then:
cd WordViewer
dotnet run
```

To open a specific file on launch:

```bash
dotnet run -- "C:\path\to\document.docx"
```

Or after publishing:

```bash
WordViewer.exe "C:\path\to\document.docx"
```

## Building a Release

```bash
dotnet publish -c Release -r win-x64 --self-contained false -o publish\
```

The `publish\` folder contains a self-contained `WordViewer.exe` that can be xcopy-deployed.

For a fully self-contained single-file executable:

```bash
dotnet publish -c Release -r win-x64 --self-contained true \
  -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true \
  -o publish\
```

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl+O` | Open document |
| `Ctrl+W` | Close document |
| `F5` | Reload current document |
| `Ctrl++` | Zoom in |
| `Ctrl+-` | Zoom out |
| `Ctrl+0` | Reset zoom to 100% |

## Limitations

Mammoth is designed to produce clean semantic HTML rather than a pixel-perfect replica of the Word layout. The following are not currently rendered:

- Headers and footers
- Page numbers
- Text boxes and floating shapes
- SmartArt and charts
- Complex mail-merge fields
- Precise font sizes and paragraph spacing

For most reading purposes (reviewing reports, feedback documents, specifications) the output is accurate and easy to read.

## Dependencies

| Package | Purpose |
|---|---|
| [Mammoth 1.3.1](https://www.nuget.org/packages/Mammoth) | DOCX → HTML conversion |
| [Microsoft.Web.WebView2](https://www.nuget.org/packages/Microsoft.Web.WebView2) | Chromium-based HTML renderer |
