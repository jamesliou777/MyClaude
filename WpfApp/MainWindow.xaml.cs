using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Windows;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;

namespace RCColumnCalculator;

public partial class MainWindow : Window
{
    private readonly string _htmlPath;
    private HttpListener? _httpListener;
    private bool _webViewReady = false;

    public MainWindow()
    {
        InitializeComponent();
        _htmlPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "wwwroot", "index.html");
        Loaded += MainWindow_Loaded;
        Closed += MainWindow_Closed;
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        await InitWebViewAsync();
    }

    private void MainWindow_Closed(object? sender, EventArgs e)
    {
        _httpListener?.Stop();
    }

    // ─────────────────────────────────────────────────────────
    // 初始化 WebView2
    // ─────────────────────────────────────────────────────────
    private async Task InitWebViewAsync()
    {
        var userDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "RCColumnCalculator", "WebView2Data");

        var env = await CoreWebView2Environment.CreateAsync(
            browserExecutableFolder: null,
            userDataFolder: userDataFolder);

        await webView.EnsureCoreWebView2Async(env);

        webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
        // 開發期間保留開發者工具；發行版可設為 false
        webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

        webView.CoreWebView2.NavigationCompleted += OnNavigationCompleted;
        webView.CoreWebView2.WebMessageReceived  += OnWebMessageReceived;

        webView.CoreWebView2.Navigate(new Uri(_htmlPath).AbsoluteUri);
    }

    // ─────────────────────────────────────────────────────────
    // 注入橋接腳本 + 啟動 HTTP API
    // ─────────────────────────────────────────────────────────
    private async void OnNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        if (!e.IsSuccess) return;

        // 注入 CSV 檔案 I/O 橋接腳本
        const string bridgeScript = """
            (function() {
                if (!window.chrome?.webview) return;

                window.exportRCColumnCSV = async function() {
                    try {
                        const csv = window._buildCSVContent();
                        if (!csv) { alert('無法產生 CSV 內容'); return; }
                        window.chrome.webview.postMessage(JSON.stringify({ type: 'saveCSV', content: csv }));
                    } catch(err) {
                        alert('CSV 匯出失敗：' + err.message);
                    }
                };

                window.importRCColumnCSV = function() {
                    window.chrome.webview.postMessage(JSON.stringify({ type: 'openCSV' }));
                };

                window.chrome.webview.addEventListener('message', function(e) {
                    const msg = JSON.parse(e.data);
                    if (msg.type === 'csvContent' && msg.content)
                        window._processImportedCSV(msg.content);
                });

                console.log('[WPF Bridge] CSV 橋接已啟用');
            })();
            """;

        await webView.CoreWebView2.ExecuteScriptAsync(bridgeScript);

        // WebView 就緒後啟動 HTTP API
        _webViewReady = true;
        StartHttpApi();
    }

    // ─────────────────────────────────────────────────────────
    // 接收 JS 訊息（CSV 檔案操作）
    // ─────────────────────────────────────────────────────────
    private async void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        try
        {
            var msg  = JsonDocument.Parse(e.WebMessageAsJson);
            var type = msg.RootElement.GetProperty("type").GetString();
            switch (type)
            {
                case "saveCSV": await HandleSaveCSV(msg.RootElement); break;
                case "openCSV": await HandleOpenCSV(); break;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"橋接錯誤：{ex.Message}");
        }
    }

    private async Task HandleSaveCSV(JsonElement root)
    {
        var content = root.GetProperty("content").GetString() ?? string.Empty;
        var dlg = new SaveFileDialog
        {
            Title = "匯出 CSV", Filter = "CSV 檔案 (*.csv)|*.csv",
            FileName = "rc_column_data.csv", DefaultExt = ".csv"
        };
        if (dlg.ShowDialog() != true) return;
        try
        {
            await File.WriteAllTextAsync(dlg.FileName, content, new UTF8Encoding(false));
            var info = new FileInfo(dlg.FileName);
            if (info.Length != Encoding.UTF8.GetByteCount(content))
                MessageBox.Show("CSV 匯出警告：檔案可能未正確寫入。", "警告",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"CSV 匯出失敗：{ex.Message}", "錯誤",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task HandleOpenCSV()
    {
        var dlg = new OpenFileDialog
        {
            Title = "匯入 CSV", Filter = "CSV 檔案 (*.csv)|*.csv", DefaultExt = ".csv"
        };
        if (dlg.ShowDialog() != true) return;
        try
        {
            var content = await File.ReadAllTextAsync(dlg.FileName, Encoding.UTF8);
            var escaped = JsonSerializer.Serialize(content);
            await webView.CoreWebView2.ExecuteScriptAsync(
                $"window.chrome.webview.postMessage(JSON.stringify({{type:'csvContent',content:{escaped}}}))" );
        }
        catch (Exception ex)
        {
            MessageBox.Show($"CSV 匯入失敗：{ex.Message}", "錯誤",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ═════════════════════════════════════════════════════════
    //  HTTP API  (localhost:5050)
    // ═════════════════════════════════════════════════════════
    private const int ApiPort = 5050;

    private void StartHttpApi()
    {
        try
        {
            _httpListener = new HttpListener();
            _httpListener.Prefixes.Add($"http://localhost:{ApiPort}/");
            _httpListener.Start();
            _ = Task.Run(ListenLoopAsync);
            Title = $"RC 柱交互作用曲線計算器  [API: http://localhost:{ApiPort}]";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"HTTP API 啟動失敗（Port {ApiPort} 可能已被佔用）：{ex.Message}",
                "API 警告", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async Task ListenLoopAsync()
    {
        while (_httpListener?.IsListening == true)
        {
            try
            {
                var ctx = await _httpListener.GetContextAsync();
                _ = Task.Run(() => HandleRequestAsync(ctx));
            }
            catch { break; }
        }
    }

    private async Task HandleRequestAsync(HttpListenerContext ctx)
    {
        var req  = ctx.Request;
        var resp = ctx.Response;
        resp.ContentType = "application/json; charset=utf-8";
        resp.Headers.Add("Access-Control-Allow-Origin", "*");

        string result;
        try
        {
            var path = req.Url?.AbsolutePath ?? "/";

            if (path == "/api/ping")
            {
                result = $"{{\"ok\":true,\"message\":\"RC Column Calculator API\",\"port\":{ApiPort}}}";
            }
            else if (path == "/api/pmcurve" && req.HttpMethod == "POST")
            {
                string body;
                using (var sr = new StreamReader(req.InputStream, Encoding.UTF8))
                    body = await sr.ReadToEndAsync();

                result = await ExecuteJsCalcAsync(body);
            }
            else
            {
                resp.StatusCode = 404;
                result = "{\"ok\":false,\"error\":\"Unknown endpoint. Use POST /api/pmcurve\"}";
            }
        }
        catch (Exception ex)
        {
            resp.StatusCode = 500;
            result = $"{{\"ok\":false,\"error\":\"{ex.Message.Replace("\"", "'")}\"}}";
        }

        var bytes = Encoding.UTF8.GetBytes(result);
        resp.ContentLength64 = bytes.Length;
        await resp.OutputStream.WriteAsync(bytes);
        resp.OutputStream.Close();
    }

    /// <summary>
    /// 在 UI 執行緒上執行 JS _apiCalc，透過 TaskCompletionSource 回傳結果
    /// </summary>
    private Task<string> ExecuteJsCalcAsync(string inputJson)
    {
        var tcs = new TaskCompletionSource<string>();

        Application.Current.Dispatcher.BeginInvoke(new Action(async () =>
        {
            try
            {
                if (!_webViewReady)
                {
                    tcs.SetResult("{\"ok\":false,\"error\":\"WebView 尚未就緒，請稍候再試\"}");
                    return;
                }
                // JsonSerializer.Serialize 會產生正確轉義的 JS 字串字面值
                var inputLiteral = JsonSerializer.Serialize(inputJson);
                var script       = $"window._apiCalc({inputLiteral})";
                var rawResult    = await webView.CoreWebView2.ExecuteScriptAsync(script);
                // rawResult 為 JSON 編碼的字串（例如 "\"{ ... }\"" ）
                var innerJson    = JsonSerializer.Deserialize<string>(rawResult) ?? "{}";
                tcs.SetResult(innerJson);
            }
            catch (Exception ex)
            {
                tcs.SetResult($"{{\"ok\":false,\"error\":\"{ex.Message.Replace("\"", "'")}\"}}");
            }
        }));

        return tcs.Task;
    }
}
