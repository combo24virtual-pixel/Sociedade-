using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Win32;

namespace MonitorEpsonL395;

internal sealed class PrintEntry
{
    public string Id { get; set; } = "";
    public DateTime Date { get; set; }
    public string Printer { get; set; } = "";
    public string Document { get; set; } = "";
    public int Pages { get; set; }
    public int Copies { get; set; } = 1;
    public string Type { get; set; } = "Não identificado";
    public decimal? Value { get; set; }
}

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

internal sealed class MainForm : Form
{
    private readonly string dataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MonitorEpsonL395");
    private readonly List<PrintEntry> entries = new();
    private readonly DataGridView grid = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
    private readonly Label summary = new() { Dock = DockStyle.Top, Height = 65, Font = new Font("Segoe UI", 11f), AutoSize = false };
    private readonly System.Windows.Forms.Timer timer = new() { Interval = 30000 };
    private readonly NotifyIcon tray = new() { Icon = SystemIcons.Application, Text = "Monitor Epson L395", Visible = true };
    private readonly Dictionary<string, DateTime> seen = new();
    private readonly string logPath;
    private readonly string dataPath;
    private bool exitRequested;

    public MainForm()
    {
        Text = "Monitor Epson L395"; Width = 1100; Height = 650; MinimumSize = new Size(800, 450);
        Directory.CreateDirectory(dataFolder);
        dataPath = Path.Combine(dataFolder, "historico.json");
        logPath = Path.Combine(dataFolder, "monitor.log");
        LoadData();
        var menu = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 48, AutoSize = false, WrapContents = false };
        AddButton(menu, "Atualizar", () => { Poll(); RefreshGrid(); });
        AddButton(menu, "Exportar CSV mensal", ExportCsv);
        AddButton(menu, "Abrir pasta de dados", () => Process.Start(new ProcessStartInfo("explorer.exe", dataFolder) { UseShellExecute = true }));
        AddButton(menu, "Iniciar com Windows", EnableStartup);
        AddButton(menu, "Sair", () => { exitRequested = true; Close(); });
        Controls.Add(grid); Controls.Add(summary); Controls.Add(menu);
        grid.Columns.Add("date", "Data e horário"); grid.Columns.Add("document", "Documento"); grid.Columns.Add("printer", "Impressora");
        grid.Columns.Add("pages", "Páginas"); grid.Columns.Add("copies", "Cópias"); grid.Columns.Add("type", "Tipo"); grid.Columns.Add("value", "Valor (R$)");
        tray.DoubleClick += (_, _) => { Show(); WindowState = FormWindowState.Normal; Activate(); };
        var context = new ContextMenuStrip();
        context.Items.Add("Abrir", null, (_, _) => { Show(); WindowState = FormWindowState.Normal; Activate(); });
        context.Items.Add("Sair", null, (_, _) => { exitRequested = true; Close(); });
        tray.ContextMenuStrip = context;
        FormClosing += (_, e) => { if (!exitRequested) { e.Cancel = true; Hide(); } else { tray.Visible = false; timer.Stop(); } };
        timer.Tick += (_, _) => Poll();
        timer.Start(); RefreshGrid(); Poll();
    }

    private static void AddButton(FlowLayoutPanel panel, string label, Action action)
    {
        var b = new Button { Text = label, AutoSize = true, Height = 36, Margin = new Padding(4) };
        b.Click += (_, _) => action(); panel.Controls.Add(b);
    }

    private void LoadData()
    {
        try { if (File.Exists(dataPath)) entries.AddRange(JsonSerializer.Deserialize<List<PrintEntry>>(File.ReadAllText(dataPath)) ?? new()); }
        catch (Exception ex) { Log(ex); MessageBox.Show("Não foi possível carregar o histórico. Os dados anteriores não serão sobrescritos.\n" + ex.Message); }
    }

    private void SaveData()
    {
        try
        {
            string temp = dataPath + ".tmp";
            File.WriteAllText(temp, JsonSerializer.Serialize(entries, new JsonSerializerOptions { WriteIndented = true }), Encoding.UTF8);
            File.Move(temp, dataPath, true);
        }
        catch (Exception ex) { Log(ex); }
    }

    private void Log(Exception ex)
    {
        try { File.AppendAllText(logPath, $"{DateTime.Now:O} {ex}\n"); } catch { }
    }

    private void Poll()
    {
        // Windows print spooler jobs may disappear before the polling interval.
        // Only currently visible jobs are captured. No print job is modified.
        try
        {
            foreach (var printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters.Cast<string>().Where(p => p.Contains("L395", StringComparison.OrdinalIgnoreCase)))
            {
                foreach (var job in SpoolerJobs.GetJobs(printer))
                {
                    string id = printer + ":" + job.JobId + ":" + job.Submitted.ToString("O");
                    if (entries.Any(e => e.Id == id) || seen.ContainsKey(id)) continue;
                    seen[id] = DateTime.Now;
                    int pages = Math.Max(0, job.Pages);
                    // Windows spooler does not reliably expose copy count or photo vs document.
                    entries.Add(new PrintEntry { Id = id, Date = job.Submitted, Printer = printer, Document = job.Document, Pages = pages, Copies = 1, Type = "Não identificado", Value = null });
                    SaveData(); RefreshGrid();
                }
            }
            foreach (var key in seen.Where(kv => kv.Value < DateTime.Now.AddDays(-2)).Select(kv => kv.Key).ToList()) seen.Remove(key);
        }
        catch (Exception ex) { Log(ex); }
    }

    private void RefreshGrid()
    {
        grid.Rows.Clear();
        foreach (var e in entries.OrderByDescending(e => e.Date))
            grid.Rows.Add(e.Date.ToString("dd/MM/yyyy HH:mm:ss"), e.Document, e.Printer, e.Pages == 0 ? "Desconhecido" : e.Pages.ToString(), e.Copies == 1 ? "Não confirmado" : e.Copies.ToString(), e.Type, e.Value?.ToString("N2") ?? "Não calculado");
        var month = entries.Where(e => e.Date.Year == DateTime.Today.Year && e.Date.Month == DateTime.Today.Month).ToList();
        summary.Text = $"Mês atual: {month.Count} trabalhos observados | {month.Sum(e => e.Pages)} páginas informadas | Total confirmado: R$ {month.Sum(e => e.Value ?? 0):N2}\n" +
                       "Tipos e cópias não confirmados pelo Windows não recebem cobrança automática. Monitoramento da fila pode perder trabalhos rápidos.";
    }

    private void ExportCsv()
    {
        using var dlg = new SaveFileDialog { Filter = "Arquivo CSV (*.csv)|*.csv", FileName = $"impressao_{DateTime.Today:yyyy_MM}.csv" };
        if (dlg.ShowDialog() != DialogResult.OK) return;
        try
        {
            static string Q(string s) => "\"" + s.Replace("\"", "\"\"") + "\"";
            var lines = new List<string> { "Data;Documento;Impressora;Paginas;Copias;Tipo;Valor" };
            lines.AddRange(entries.Where(e => e.Date.Year == DateTime.Today.Year && e.Date.Month == DateTime.Today.Month)
                .Select(e => string.Join(";", new[] { Q(e.Date.ToString("dd/MM/yyyy HH:mm:ss")), Q(e.Document), Q(e.Printer), e.Pages.ToString(), Q(e.Copies == 1 ? "Não confirmado" : e.Copies.ToString()), Q(e.Type), e.Value?.ToString("0.00", CultureInfo.GetCultureInfo("pt-BR")) ?? "" })));
            File.WriteAllLines(dlg.FileName, lines, new UTF8Encoding(true));
            MessageBox.Show("Relatório exportado.");
        }
        catch (Exception ex) { Log(ex); MessageBox.Show("Erro ao exportar: " + ex.Message); }
    }

    private void EnableStartup()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);
            key?.SetValue("MonitorEpsonL395", "\"" + Application.ExecutablePath + "\"");
            MessageBox.Show("Inicialização automática ativada para este usuário.");
        }
        catch (Exception ex) { Log(ex); MessageBox.Show("Não foi possível ativar: " + ex.Message); }
    }
}

internal static class SpoolerJobs
{
    // WMI is available on Windows 10; System.Management requires a NuGet package.
    // To keep the app dependency-free, use PowerShell as a read-only Windows component.
    internal record Job(string JobId, string Document, int Pages, DateTime Submitted);
    internal static IEnumerable<Job> GetJobs(string printer)
    {
        // Print queue enumeration through Windows PowerShell Get-PrintJob; no printer settings changed.
        var escaped = printer.Replace("'", "''");
        var command = "$ErrorActionPreference='SilentlyContinue'; Get-PrintJob -PrinterName '" + escaped + "' | Select-Object ID,DocumentName,TotalPages,SubmittedTime | ConvertTo-Json -Compress";
        var bytes = Encoding.Unicode.GetBytes(command);
        var psi = new ProcessStartInfo("powershell.exe", "-NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(bytes))
        { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true };
        using var proc = Process.Start(psi);
        if (proc == null) yield break;
        string output = proc.StandardOutput.ReadToEnd();
        if (!proc.WaitForExit(10000)) { try { proc.Kill(); } catch { } yield break; }
        if (string.IsNullOrWhiteSpace(output)) yield break;
        using var doc = JsonDocument.Parse(output);
        var root = doc.RootElement;
        foreach (var element in root.ValueKind == JsonValueKind.Array ? root.EnumerateArray().ToArray() : new[] { root })
        {
            string id = element.TryGetProperty("ID", out var x) ? x.ToString() : "";
            string name = element.TryGetProperty("DocumentName", out var n) ? n.ToString() : "Documento";
            int pages = element.TryGetProperty("TotalPages", out var p) && p.TryGetInt32(out var pg) ? pg : 0;
            DateTime submitted = element.TryGetProperty("SubmittedTime", out var t) && DateTime.TryParse(t.ToString(), out var dt) ? dt : DateTime.Now;
            yield return new Job(id, name, pages, submitted);
        }
    }
}
