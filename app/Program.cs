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
using System.Threading.Tasks;
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
        using var mutex = new System.Threading.Mutex(true, @"Local\MonitorEpsonL395_InstanciaUnica", out bool firstInstance);
        if (!firstInstance) return;
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

internal sealed class MainForm : Form
{
    private readonly string dataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MonitorEpsonL395");
    private readonly List<PrintEntry> entries = new();
    private readonly DataGridView grid = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
    private readonly Label summary = new() { Dock = DockStyle.Top, Height = 58, Font = new Font("Segoe UI", 11f), AutoSize = false };
    private readonly System.Windows.Forms.Timer timer = new() { Interval = 5000 };
    private readonly NotifyIcon tray = new() { Icon = SystemIcons.Application, Text = "Monitor Epson L395", Visible = true };
    private readonly Dictionary<string, DateTime> seen = new();
    private readonly string logPath;
    private readonly string dataPath;
    private bool polling;
    private bool historyReadable = true;
    private readonly ContextMenuStrip jobMenu = new();
    private long lastEventId;
    private readonly string checkpointPath;
    private string monitorStatus = "Verificando impressoras...";
    private string detectedPrinters = "Nenhuma impressora identificada";

    public MainForm()
    {
        Text = "Monitor Epson L395"; Width = 1100; Height = 650; MinimumSize = new Size(800, 450);
        Directory.CreateDirectory(dataFolder);
        dataPath = Path.Combine(dataFolder, "historico.json");
        logPath = Path.Combine(dataFolder, "monitor.log");
        checkpointPath = Path.Combine(dataFolder, "ultimo_evento.txt");
        try { if (File.Exists(checkpointPath)) long.TryParse(File.ReadAllText(checkpointPath), out lastEventId); }
        catch (Exception ex) { Log(ex); }
        LoadData();
        var menu = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 48, AutoSize = false, WrapContents = false };
        AddButton(menu, "Atualizar", () => { _ = PollAsync(); });
        AddButton(menu, "Exportar CSV mensal", ExportCsv);
        AddButton(menu, "Zerar histórico", ResetHistory);
        AddButton(menu, "Verificar registro", ShowMonitoringHelp);
        AddButton(menu, "Abrir pasta de dados", () => Process.Start(new ProcessStartInfo("explorer.exe", dataFolder) { UseShellExecute = true }));
        Controls.Add(grid); Controls.Add(summary); Controls.Add(menu);
        grid.Columns.Add("date", "Data e horário"); grid.Columns.Add("document", "Documento"); grid.Columns.Add("printer", "Impressora");
        grid.Columns.Add("pages", "Páginas"); grid.Columns.Add("copies", "Cópias"); grid.Columns.Add("type", "Tipo"); grid.Columns.Add("value", "Valor (R$)");
        jobMenu.Items.Add("Papel fotográfico — R$ 6,00", null, (_, _) => SetPaperType("Fotográfico", 6.00m));
        jobMenu.Items.Add("Papel comum — R$ 0,50", null, (_, _) => SetPaperType("Comum", 0.50m));
        jobMenu.Items.Add("Corrigir quantidade de folhas", null, (_, _) => EditSheetCount());
        grid.ContextMenuStrip = jobMenu;
        grid.CellMouseDown += (_, e) =>
        {
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0)
            {
                grid.ClearSelection();
                grid.Rows[e.RowIndex].Selected = true;
                grid.CurrentCell = grid.Rows[e.RowIndex].Cells[0];
            }
        };
        tray.DoubleClick += (_, _) => { Show(); WindowState = FormWindowState.Normal; Activate(); };
        var context = new ContextMenuStrip();
        context.Items.Add("Abrir", null, (_, _) => { Show(); WindowState = FormWindowState.Normal; Activate(); });
        tray.ContextMenuStrip = context;
        FormClosing += (_, e) =>
        {
            // O X apenas oculta a janela. O Windows pode encerrar o processo ao desligar.
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
            else
            {
                timer.Stop();
                tray.Visible = false;
            }
        };
        EnableStartupSilently();
        timer.Tick += (_, _) => { _ = PollAsync(); };
        timer.Start(); RefreshGrid(); _ = PollAsync();
    }

    private static void AddButton(FlowLayoutPanel panel, string label, Action action)
    {
        var b = new Button { Text = label, AutoSize = true, Height = 36, Margin = new Padding(4) };
        b.Click += (_, _) => action(); panel.Controls.Add(b);
    }

    private void LoadData()
    {
        try { if (File.Exists(dataPath)) entries.AddRange(JsonSerializer.Deserialize<List<PrintEntry>>(File.ReadAllText(dataPath)) ?? new()); }
        catch (Exception ex) { historyReadable = false; Log(ex); MessageBox.Show("Não foi possível carregar o histórico. O arquivo será preservado.\n" + ex.Message); }
    }

    private void SaveData()
    {
        if (!historyReadable) return;
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

    private void ResetHistory()
    {
        if (MessageBox.Show("Deseja apagar os registros do monitor? Esta ação não pode ser desfeita.\nO arquivo anterior será guardado como cópia de segurança.", "Zerar histórico", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
        try
        {
            if (File.Exists(dataPath))
            {
                string backup = Path.Combine(dataFolder, "historico_backup_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + ".json");
                File.Copy(dataPath, backup);
            }
            entries.Clear();
            seen.Clear();
            // Preserva o checkpoint para não reimportar impressões antigas.
            SaveData();
            RefreshGrid();
            MessageBox.Show("Histórico zerado. A cópia de segurança foi mantida na pasta de dados.");
        }
        catch (Exception ex) { Log(ex); MessageBox.Show("Não foi possível zerar o histórico: " + ex.Message); }
    }

    private async Task PollAsync()
    {
        if (polling || IsDisposed || !historyReadable) return;
        polling = true;
        try
        {
            long checkpoint = lastEventId;
            var result = await Task.Run(() => PrintEvents.Read(checkpoint));
            if (IsDisposed) return;
            var printers = System.Drawing.Printing.PrinterSettings.InstalledPrinters.Cast<string>().ToList();
            detectedPrinters = printers.Count == 0 ? "Nenhuma impressora instalada" : string.Join(", ", printers);
            if (!result.Enabled)
            {
                monitorStatus = "Registro de impressão desativado — clique em Verificar registro";
            }
            else
            {
                monitorStatus = "Monitorando automaticamente";
                // Na primeira execução, estabelece o ponto inicial para não importar trabalhos antigos.
                if (!File.Exists(checkpointPath))
                {
                    lastEventId = result.HighestId;
                    File.WriteAllText(checkpointPath, lastEventId.ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    bool changed = false;
                    foreach (var item in result.Events.OrderBy(e => e.RecordId))
                    {
                        string id = "print-event:" + item.RecordId;
                        if (entries.Any(e => e.Id == id)) continue;
                        // Eventos 307 indicam conclusão informada pelo spooler, não comprovam saída física.
                        entries.Add(new PrintEntry { Id = id, Date = item.Date, Printer = item.Printer,
                            Document = item.Document, Pages = Math.Max(0, item.Pages), Copies = 1,
                            Type = "Não identificado", Value = null });
                        changed = true;
                    }
                    if (changed) SaveData();
                    if (result.HighestId > lastEventId)
                    {
                        lastEventId = result.HighestId;
                        File.WriteAllText(checkpointPath, lastEventId.ToString(CultureInfo.InvariantCulture));
                    }
                }
            }
            RefreshGrid();
        }
        catch (Exception ex)
        {
            monitorStatus = "Falha no monitoramento (consulte monitor.log)";
            Log(ex); RefreshGrid();
        }
        finally { polling = false; }
    }

    private PrintEntry? SelectedEntry()
    {
        if (grid.SelectedRows.Count == 0) return null;
        return grid.SelectedRows[0].Tag as PrintEntry;
    }

    private void SetPaperType(string type, decimal price)
    {
        var entry = SelectedEntry();
        if (entry == null) return;
        if (entry.Pages <= 0)
        {
            MessageBox.Show("Primeiro informe a quantidade de folhas deste trabalho.");
            return;
        }
        entry.Type = type;
        entry.Value = entry.Pages * price;
        SaveData();
        RefreshGrid();
    }

    private void EditSheetCount()
    {
        var entry = SelectedEntry();
        if (entry == null) return;
        using var dialog = new Form { Text = "Quantidade de folhas", Width = 340, Height = 165,
            FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false };
        var number = new NumericUpDown { Minimum = 1, Maximum = 100000, Value = Math.Max(1, entry.Pages),
            Left = 16, Top = 18, Width = 290 };
        var ok = new Button { Text = "Confirmar", DialogResult = DialogResult.OK, Left = 186, Top = 62, Width = 120 };
        dialog.Controls.Add(number);
        dialog.Controls.Add(ok);
        dialog.AcceptButton = ok;
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        entry.Pages = (int)number.Value;
        entry.Value = entry.Type switch
        {
            "Fotográfico" => entry.Pages * 6.00m,
            "Comum" => entry.Pages * 0.50m,
            _ => null
        };
        SaveData();
        RefreshGrid();
    }

    private void RefreshGrid()
    {
        grid.Rows.Clear();
        foreach (var e in entries.OrderByDescending(e => e.Date))
        {
            int index = grid.Rows.Add(e.Date.ToString("dd/MM/yyyy HH:mm:ss"), e.Document, e.Printer, e.Pages == 0 ? "Desconhecido" : e.Pages.ToString(), e.Copies == 1 ? "Não confirmado" : e.Copies.ToString(), e.Type, e.Value?.ToString("N2") ?? "Pendente");
            grid.Rows[index].Tag = e;
        }
        var month = entries.Where(e => e.Date.Year == DateTime.Today.Year && e.Date.Month == DateTime.Today.Month).ToList();
        summary.Text = $"Impressoras: {detectedPrinters} | {monitorStatus}\n" +
                       $"Mês atual: {month.Count} trabalhos | {month.Sum(e => e.Pages)} páginas informadas | R$ {month.Sum(e => e.Value ?? 0):N2} calculados | {month.Count(e => e.Value == null)} pendentes";
    }

    private void ShowMonitoringHelp()
    {
        MessageBox.Show(this,
            "O Windows precisa registrar os trabalhos de impressão.\n\n" +
            "Para ativar: abra o Visualizador de Eventos > Logs de Aplicativos e Serviços > Microsoft > Windows > PrintService > Operacional. " +
            "Clique com o botão direito em Operacional e escolha Habilitar Log.\n\n" +
            "Depois imprima uma página de teste. O monitor identifica trabalhos concluídos pelo Windows; " +
            "isso não comprova a saída física da folha.\n\n" +
            "O X apenas esconde a janela; o monitor permanece perto do relógio enquanto sua sessão do Windows estiver ativa.",
            "Verificar monitoramento", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

    private void EnableStartupSilently()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true)
                ?? throw new IOException("Não foi possível acessar a inicialização do Windows.");
            key.SetValue("MonitorEpsonL395", "\"" + Application.ExecutablePath + "\"");
        }
        catch (Exception ex)
        {
            Log(ex);
            monitorStatus = "Não foi possível configurar o início automático: " + ex.Message;
        }
    }

}

internal static class PrintEvents
{
    internal record Item(long RecordId, DateTime Date, string Printer, string Document, int Pages);
    internal record Result(bool Enabled, long HighestId, List<Item> Events);

    internal static Result Read(long afterId)
    {
        // Evento 307: documento impresso segundo o serviço de impressão do Windows.
        // Não altera a fila nem as configurações da impressora.
        const string script = @"
$ErrorActionPreference = 'Stop'
$log = 'Microsoft-Windows-PrintService/Operational'
$info = Get-WinEvent -ListLog $log -ErrorAction Stop
if (-not $info.IsEnabled) { [Console]::Out.WriteLine('{""enabled"":false}'); exit 0 }
$all = @(Get-WinEvent -FilterHashtable @{LogName=$log; Id=307} -MaxEvents 500 -ErrorAction SilentlyContinue)
$max = 0
if ($all.Count -gt 0) { $max = [long]$all[0].RecordId }
$events = @($all | Where-Object { [long]$_.RecordId -gt $after } | ForEach-Object {
  $v = @($_.Properties | ForEach-Object { [string]$_.Value })
  [pscustomobject]@{
    id = [long]$_.RecordId
    date = $_.TimeCreated.ToString('o')
    document = if ($v.Count -gt 1) { $v[1] } else { 'Documento' }
    printer = if ($v.Count -gt 4) { $v[4] } else { 'Impressora desconhecida' }
    pages = if ($v.Count -gt 7) { $v[7] } else { '0' }
  }
})
[pscustomobject]@{ enabled=$true; max=$max; events=@($events) } | ConvertTo-Json -Depth 5 -Compress
";
        var command = "$after=" + afterId.ToString(CultureInfo.InvariantCulture) + ";" + script;
        var psi = new ProcessStartInfo("powershell.exe", "-NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(command)))
        { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true };
        using var process = Process.Start(psi) ?? throw new IOException("Não foi possível iniciar o PowerShell.");
        string output = process.StandardOutput.ReadToEnd();
        string errors = process.StandardError.ReadToEnd();
        if (!process.WaitForExit(12000))
        {
            try { process.Kill(); } catch { }
            throw new TimeoutException("Consulta ao serviço de impressão demorou demais.");
        }
        if (process.ExitCode != 0 || string.IsNullOrWhiteSpace(output))
            throw new IOException("Falha ao consultar o registro de impressão: " + errors);
        using var json = JsonDocument.Parse(output);
        var root = json.RootElement;
        if (!root.GetProperty("enabled").GetBoolean()) return new Result(false, afterId, new());
        long highest = root.GetProperty("max").GetInt64();
        var events = new List<Item>();
        if (root.TryGetProperty("events", out var array) && array.ValueKind == JsonValueKind.Array)
            foreach (var e in array.EnumerateArray())
            {
                long id = e.GetProperty("id").GetInt64();
                string dateText = e.GetProperty("date").GetString() ?? "";
                DateTime.TryParse(dateText, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date);
                int.TryParse(e.GetProperty("pages").ToString(), out int pages);
                events.Add(new Item(id, date == default ? DateTime.Now : date,
                    e.GetProperty("printer").GetString() ?? "", e.GetProperty("document").GetString() ?? "", pages));
            }
        if (root.TryGetProperty("events", out var single) && single.ValueKind == JsonValueKind.Object)
        {
            long id = single.GetProperty("id").GetInt64();
            DateTime.TryParse(single.GetProperty("date").GetString(), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date);
            int.TryParse(single.GetProperty("pages").ToString(), out int pages);
            events.Add(new Item(id, date == default ? DateTime.Now : date,
                single.GetProperty("printer").GetString() ?? "", single.GetProperty("document").GetString() ?? "", pages));
        }
        return new Result(true, highest, events);
    }
}
