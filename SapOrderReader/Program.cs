using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text.Json;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

namespace SapOrderReader;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

public sealed class MainForm : Form
{
    private readonly Button runButton = new();
    private readonly Label statusLabel = new();
    private readonly TextBox resultBox = new();

    private const int TableLeft = 0, TableTop = 35, TableWidth = 320, TableHeight = 460;
    private const int OrderLeft = 0, OrderTop = 35, OrderWidth = 850, OrderHeight = 230;
    private const int SelectionMarginX = 16;

    private string ProjectFolder =>
        Directory.GetParent(AppContext.BaseDirectory)!.Parent!.Parent!.Parent!.FullName;

    public MainForm()
    {
        Text = "SAP Order Reader";
        Size = new Size(760, 560);
        StartPosition = FormStartPosition.CenterScreen;
        TopMost = true;
        WindowState = FormWindowState.Minimized;
        ShowInTaskbar = false;

        runButton.Text = "Read SAP Order / Tenant / GroupTag";
        runButton.SetBounds(20, 20, 290, 36);
        runButton.Click += RunButton_Click;

        statusLabel.Text = "Starting...";
        statusLabel.SetBounds(330, 25, 380, 35);
        statusLabel.ForeColor = Color.DarkBlue;

        resultBox.SetBounds(20, 80, 690, 410);
        resultBox.Multiline = true;
        resultBox.ScrollBars = ScrollBars.Both;
        resultBox.ReadOnly = true;
        resultBox.Font = new Font("Consolas", 10);

        Controls.AddRange(new Control[] { runButton, statusLabel, resultBox });

        Shown += async (_, _) =>
        {
            await Task.Delay(250);
            runButton.PerformClick();
        };
    }

    private async void RunButton_Click(object? sender, EventArgs e)
    {
        runButton.Enabled = false;

        try
        {
            Hide();
            await Task.Delay(700);

            // Do not OCR the order screen before PSD.
            // This prevents large/different orders from blocking navigation.
            statusLabel.Text = "Opening PSD...";
            await ActivatePsdAsync();
            await Task.Delay(1200);

            statusLabel.Text = "Finding Process ID ENDUSER...";
            int processY = await FindEnduserYAsync();
            MouseHelper.Click(SelectionMarginX, processY);
            await Task.Delay(350);
            await KeyboardNavigation.TabEnterAsync(2);
            await Task.Delay(1400);

            statusLabel.Text = "Finding Items ENDUSER...";
            int itemY = await FindEnduserYAsync();
            MouseHelper.Click(SelectionMarginX, itemY);
            await Task.Delay(350);
            await KeyboardNavigation.TabEnterAsync(4);
            await Task.Delay(1600);

            statusLabel.Text = "Copying Tenant and GroupTag...";
            Values values = await ReadValuesByKeyboardAsync();

            statusLabel.Text = "Returning to the order...";
            await KeyboardNavigation.TabEnterAsync(3);
            await Task.Delay(1000);
            await KeyboardNavigation.TabEnterAsync(5);
            await Task.Delay(1000);
            await KeyboardNavigation.TabEnterAsync(3);
            await Task.Delay(1400);

            statusLabel.Text = "Reading Order Number and optional Qty...";
            OrderInfo order = await ReadOrderInfoAsync();

            var output = new SapResult
            {
                OrderNumber = order.OrderNumber,
                Qty = order.Qty,
                Tenant = values.Tenant,
                GroupTag = values.GroupTag,
                CreatedUtc = DateTime.UtcNow
            };

            string json = JsonSerializer.Serialize(output,
                new JsonSerializerOptions { WriteIndented = true });

            string projectResult = Path.Combine(ProjectFolder, "SapOrderReader_result.json");
            string runtimeResult = Path.Combine(AppContext.BaseDirectory, "SapOrderReader_result.json");

            File.WriteAllText(projectResult, json);
            if (!string.Equals(projectResult, runtimeResult, StringComparison.OrdinalIgnoreCase))
                File.WriteAllText(runtimeResult, json);

            Close();
        }
        catch (Exception ex)
        {
            Show();
            Activate();
            statusLabel.ForeColor = Color.Red;
            statusLabel.Text = "SAP automation stopped safely.";
            resultBox.Text = ex.ToString();
            runButton.Enabled = true;
        }
    }

    private static async Task ActivatePsdAsync()
    {
        for (int i = 0; i < 4; i++)
        {
            NativeKeyboard.Hotkey(NativeKeyboard.Key.Control, NativeKeyboard.Key.Tab);
            await Task.Delay(120);
        }

        NativeKeyboard.KeyPress(NativeKeyboard.Key.Tab);
        await Task.Delay(180);
        NativeKeyboard.KeyPress(NativeKeyboard.Key.Enter);
    }

    private async Task<OrderInfo> ReadOrderInfoAsync()
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(12);
        Exception? last = null;

        while (DateTime.UtcNow < deadline)
        {
            try
            {
                using Bitmap image = Capture(OrderLeft, OrderTop, OrderWidth, OrderHeight);
                OcrResult ocr = await OcrAsync(image);
                string? order = FindOrderNumber(ocr);

                if (!string.IsNullOrWhiteSpace(order))
                    return new OrderInfo(order, FindQuantity(ocr) ?? "");

                last = new InvalidOperationException("Order number not visible yet.");
            }
            catch (Exception ex) { last = ex; }

            await Task.Delay(500);
        }

        throw new InvalidOperationException("Order number was not found.", last);
    }

    private static string? FindOrderNumber(OcrResult ocr)
    {
        foreach (OcrLine line in ocr.Lines)
        {
            foreach (OcrWord word in line.Words)
            {
                string d = new(word.Text.Where(char.IsDigit).ToArray());
                if (d.Length == 10 && d.StartsWith("7")) return d;
            }

            string all = new(line.Text.Where(char.IsDigit).ToArray());
            for (int i = 0; i + 10 <= all.Length; i++)
            {
                string c = all.Substring(i, 10);
                if (c.StartsWith("7")) return c;
            }
        }
        return null;
    }

    private static string? FindQuantity(OcrResult ocr)
    {
        foreach (OcrLine line in ocr.Lines)
        {
            string n = Normalize(line.Text);
            if (!n.Contains("7047342") || !n.Contains("DKCO000382") ||
                !n.Contains("PC") ||
                (!n.Contains("UPLOAD") && !n.Contains("AUTOPILOT")))
                continue;

            List<string> tokens = line.Words.Select(w => w.Text.Trim())
                .Where(s => s.Length > 0).ToList();

            for (int i = 1; i < tokens.Count; i++)
                if (Normalize(tokens[i]) == "PC" && IsNumeric(tokens[i - 1]))
                    return tokens[i - 1].Replace(",", ".").Trim();
        }
        return null;
    }

    private static bool IsNumeric(string s) =>
        decimal.TryParse(s.Replace(",", "."),
            System.Globalization.NumberStyles.Number,
            System.Globalization.CultureInfo.InvariantCulture, out _);

    private async Task<int> FindEnduserYAsync()
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(15);
        Exception? last = null;

        while (DateTime.UtcNow < deadline)
        {
            try
            {
                using Bitmap image = Capture(TableLeft, TableTop, TableWidth, TableHeight);
                OcrResult ocr = await OcrAsync(image);
                List<DetectedWord> matches = FindWords(ocr, "ENDUSER");

                if (matches.Count == 1)
                    return TableTop + matches[0].Y + matches[0].Height / 2;

                if (matches.Count > 1)
                    throw new InvalidOperationException("More than one ENDUSER was found.");

                last = new InvalidOperationException("ENDUSER not visible yet.");
            }
            catch (InvalidOperationException ex) { last = ex; }

            await Task.Delay(500);
        }

        throw new InvalidOperationException("ENDUSER was not found within 15 seconds.", last);
    }

    private async Task<Values> ReadValuesByKeyboardAsync()
    {
        string diagnostic = Path.Combine(AppContext.BaseDirectory, "SapValuesKeyboard.txt");
        List<string> log = new();

        await Task.Delay(700);

        string tenant = await CopyFocusedValueAsync("Tenant", log);
        NativeKeyboard.KeyPress(NativeKeyboard.Key.Down);
        await Task.Delay(450);
        string group = await CopyFocusedValueAsync("GroupTag", log);

        log.Add($"Tenant length: {tenant.Length}");
        log.Add($"GroupTag length: {group.Length}");
        File.WriteAllLines(diagnostic, log);

        if (string.IsNullOrWhiteSpace(tenant))
            throw new InvalidOperationException("Tenant clipboard value was empty.");

        // GroupTag is intentionally allowed to be blank.
        return new Values(tenant.Trim(), group.Trim());
    }

    private static async Task<string> CopyFocusedValueAsync(string label, List<string> log)
    {
        ClipboardHelper.Clear();
        NativeKeyboard.Hotkey(NativeKeyboard.Key.Control, NativeKeyboard.Key.A);
        await Task.Delay(300);
        NativeKeyboard.Hotkey(NativeKeyboard.Key.Control, NativeKeyboard.Key.C);
        await Task.Delay(500);

        string value = await ClipboardHelper.ReadTextOrEmptyAsync();
        log.Add($"{label}: {(string.IsNullOrWhiteSpace(value) ? "<blank>" : value.Trim())}");
        return value.Trim();
    }

    private static string Normalize(string value) =>
        new(value.ToUpperInvariant().Where(char.IsLetterOrDigit).ToArray());

    private static List<DetectedWord> FindWords(OcrResult ocr, string wanted)
    {
        string target = Normalize(wanted);
        List<DetectedWord> result = new();

        foreach (OcrLine line in ocr.Lines)
        foreach (OcrWord word in line.Words)
        {
            string text = Normalize(word.Text);
            if (text == target || text.Contains(target))
                result.Add(new DetectedWord(word.Text,
                    (int)word.BoundingRect.X, (int)word.BoundingRect.Y,
                    (int)word.BoundingRect.Width, (int)word.BoundingRect.Height));
        }
        return result;
    }

    private static Bitmap Capture(int left, int top, int width, int height)
    {
        Bitmap bitmap = new(width, height, PixelFormat.Format32bppArgb);
        using Graphics graphics = Graphics.FromImage(bitmap);
        graphics.CopyFromScreen(left, top, 0, 0, new Size(width, height));
        return bitmap;
    }

    private static async Task<OcrResult> OcrAsync(Bitmap bitmap)
    {
        using MemoryStream memory = new();
        bitmap.Save(memory, ImageFormat.Png);

        using InMemoryRandomAccessStream stream = new();
        using DataWriter writer = new(stream);
        writer.WriteBytes(memory.ToArray());
        await writer.StoreAsync();
        await writer.FlushAsync();
        writer.DetachStream();
        stream.Seek(0);

        BitmapDecoder decoder = await BitmapDecoder.CreateAsync(stream);
        SoftwareBitmap software = await decoder.GetSoftwareBitmapAsync();

        OcrEngine? engine = OcrEngine.TryCreateFromUserProfileLanguages();
        if (engine is null)
        {
            var langs = OcrEngine.AvailableRecognizerLanguages;
            if (langs.Count == 0) throw new InvalidOperationException("No Windows OCR language is available.");
            engine = OcrEngine.TryCreateFromLanguage(langs[0]);
        }

        if (engine is null) throw new InvalidOperationException("Windows OCR could not be initialized.");
        return await engine.RecognizeAsync(software);
    }

    private sealed record OrderInfo(string OrderNumber, string Qty);
    private sealed record DetectedWord(string Text, int X, int Y, int Width, int Height);
    private sealed record Values(string Tenant, string GroupTag);

    private sealed class SapResult
    {
        public string OrderNumber { get; set; } = "";
        public string Qty { get; set; } = "";
        public string Tenant { get; set; } = "";
        public string GroupTag { get; set; } = "";
        public DateTime CreatedUtc { get; set; }
    }

    private static class ClipboardHelper
    {
        public static void Clear()
        {
            for (int i = 0; i < 5; i++)
            {
                try { Clipboard.Clear(); return; }
                catch { Thread.Sleep(100); }
            }
        }

        public static async Task<string> ReadTextOrEmptyAsync()
        {
            for (int i = 0; i < 8; i++)
            {
                try
                {
                    if (Clipboard.ContainsText()) return Clipboard.GetText().Trim();
                }
                catch { }
                await Task.Delay(120);
            }
            return "";
        }
    }

    private static class NativeKeyboard
    {
        public enum Key : byte { Control = 0x11, Tab = 0x09, Enter = 0x0D, Down = 0x28, A = 0x41, C = 0x43 }
        private const uint KeyUp = 0x0002;

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte key, byte scan, uint flags, UIntPtr extra);

        public static void KeyPress(Key key)
        {
            keybd_event((byte)key, 0, 0, UIntPtr.Zero);
            Thread.Sleep(80);
            keybd_event((byte)key, 0, KeyUp, UIntPtr.Zero);
            Thread.Sleep(120);
        }

        public static void Hotkey(Key modifier, Key key)
        {
            keybd_event((byte)modifier, 0, 0, UIntPtr.Zero);
            Thread.Sleep(100);
            keybd_event((byte)key, 0, 0, UIntPtr.Zero);
            Thread.Sleep(100);
            keybd_event((byte)key, 0, KeyUp, UIntPtr.Zero);
            Thread.Sleep(100);
            keybd_event((byte)modifier, 0, KeyUp, UIntPtr.Zero);
            Thread.Sleep(150);
        }
    }

    private static class MouseHelper
    {
        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);

        private const uint LeftDown = 0x0002, LeftUp = 0x0004;

        public static void Click(int x, int y)
        {
            if (!SetCursorPos(x, y)) throw new InvalidOperationException("Could not move mouse.");
            Thread.Sleep(200);
            mouse_event(LeftDown, 0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(LeftUp, 0, 0, 0, UIntPtr.Zero);
        }
    }
}
