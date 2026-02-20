using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;

namespace RCColumnCalculator;

public partial class MainWindow : Window
{
    // index.html 所在路徑（放於 wwwroot 子目錄）
    private readonly string _htmlPath;

    public MainWindow()
    {
        InitializeComponent();

        _htmlPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "wwwroot", "index.html");

        Loaded += MainWindow_Loaded;
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        await InitWebViewAsync();
    }

    // ─────────────────────────────────────────────────────────
    // 初始化 WebView2
    // ─────────────────────────────────────────────────────────
    private async Task InitWebViewAsync()
    {
        // UserData 資料夾：存放 WebView2 快取
        var userDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "RCColumnCalculator", "WebView2Data");

        var env = await CoreWebView2Environment.CreateAsync(
            browserExecutableFolder: null,
            userDataFolder: userDataFolder);

        await webView.EnsureCoreWebView2Async(env);

        // 關閉右鍵選單（視需求可移除）
        webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
        // 關閉開發者工具（發行版時啟用此行）
        // webView.CoreWebView2.Settings.AreDevToolsEnabled = false;

        // 頁面載入完成後注入橋接腳本
        webView.CoreWebView2.NavigationCompleted += OnNavigationCompleted;

        // 接收來自 JS 的訊息
        webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

        // 載入 index.html
        webView.CoreWebView2.Navigate(new Uri(_htmlPath).AbsoluteUri);
    }

    // ─────────────────────────────────────────────────────────
    // 頁面載入完成：注入 C#↔JS 橋接腳本
    // ─────────────────────────────────────────────────────────
    private async void OnNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        if (!e.IsSuccess) return;

        // 覆寫 exportRCColumnCSV 與 importRCColumnCSV，改由 C# 處理檔案 I/O
        const string bridgeScript = """
            (function() {
                if (!window.chrome?.webview) return;

                // ── 匯出 CSV ──────────────────────────────────────────
                window.exportRCColumnCSV = async function() {
                    try {
                        const csv = window._buildCSVContent();
                        if (!csv) { alert('無法產生 CSV 內容'); return; }
                        window.chrome.webview.postMessage(JSON.stringify({
                            type: 'saveCSV',
                            content: csv
                        }));
                    } catch(err) {
                        alert('CSV 匯出失敗：' + err.message);
                    }
                };

                // ── 匯入 CSV ──────────────────────────────────────────
                window.importRCColumnCSV = function() {
                    window.chrome.webview.postMessage(JSON.stringify({
                        type: 'openCSV'
                    }));
                };

                // ── 接收 C# 回傳的匯入內容 ────────────────────────────
                window.chrome.webview.addEventListener('message', function(e) {
                    const msg = JSON.parse(e.data);
                    if (msg.type === 'csvContent' && msg.content) {
                        window._processImportedCSV(msg.content);
                    }
                });

                console.log('[WPF Bridge] 已啟用 C# 檔案橋接');
            })();
            """;

        await webView.CoreWebView2.ExecuteScriptAsync(bridgeScript);
    }

    // ─────────────────────────────────────────────────────────
    // 接收來自 JS 的訊息
    // ─────────────────────────────────────────────────────────
    private async void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        try
        {
            var msg = JsonDocument.Parse(e.WebMessageAsJson);
            var type = msg.RootElement.GetProperty("type").GetString();

            switch (type)
            {
                case "saveCSV":
                    await HandleSaveCSV(msg.RootElement);
                    break;

                case "openCSV":
                    await HandleOpenCSV();
                    break;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"橋接訊息處理失敗：{ex.Message}", "錯誤",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ─────────────────────────────────────────────────────────
    // 儲存 CSV
    // ─────────────────────────────────────────────────────────
    private async Task HandleSaveCSV(JsonElement root)
    {
        var content = root.GetProperty("content").GetString() ?? string.Empty;

        var dlg = new SaveFileDialog
        {
            Title = "匯出 CSV",
            Filter = "CSV 檔案 (*.csv)|*.csv|所有檔案 (*.*)|*.*",
            FileName = "rc_column_data.csv",
            DefaultExt = ".csv"
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            await File.WriteAllTextAsync(dlg.FileName, content, new UTF8Encoding(false));

            // 驗證檔案大小
            var info = new FileInfo(dlg.FileName);
            var expectedBytes = Encoding.UTF8.GetByteCount(content);
            if (info.Length != expectedBytes)
            {
                MessageBox.Show("CSV 匯出警告：檔案可能未正確寫入。", "警告",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"CSV 匯出失敗：檔案可能被其他程式（如 Excel）鎖定，請關閉後重試。\n({ex.Message})",
                "匯出失敗", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ─────────────────────────────────────────────────────────
    // 開啟 CSV
    // ─────────────────────────────────────────────────────────
    private async Task HandleOpenCSV()
    {
        var dlg = new OpenFileDialog
        {
            Title = "匯入 CSV",
            Filter = "CSV 檔案 (*.csv)|*.csv|所有檔案 (*.*)|*.*",
            DefaultExt = ".csv"
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            var content = await File.ReadAllTextAsync(dlg.FileName, Encoding.UTF8);

            // 將檔案內容回傳給 JS（跳脫特殊字元）
            var escaped = JsonSerializer.Serialize(content);
            await webView.CoreWebView2.ExecuteScriptAsync(
                $"window.chrome.webview.postMessage(JSON.stringify({{type:'csvContent', content:{escaped}}}))" );
        }
        catch (Exception ex)
        {
            MessageBox.Show($"CSV 匯入失敗：{ex.Message}", "匯入失敗",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
