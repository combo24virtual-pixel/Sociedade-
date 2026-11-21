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
            lines.AddRange(entries.Where(e => string.Equals(e.Printer.Trim(), TargetPrinter, StringComparison.OrdinalIgnoreCase)
                && e.Date.Year == DateTime.Today.Year && e.Date.Month == DateTime.Today.Month)
                .Select(e => string.Join(";", new[] { Q(e.Date.ToString("dd/MM/yyyy HH:mm:ss")), Q(e.Document), Q(e.Printer), e.Pages.ToString(), Q(e.Copies <= 0 ? "Não confirmado" : e.Copies.ToString()), Q(e.Type), e.Value?.ToString("0.00", CultureInfo.GetCultureInfo("pt-BR")) ?? "" })));
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
$all = @(Get-WinEvent -FilterHashtable @{LogName=$log; Id=307} -MaxEvents 2000 -ErrorAction SilentlyContinue)
$max = 0
if ($all.Count -gt 0) { $max = [long]$all[0].RecordId }
$events = @($all | Where-Object { [long]$_.RecordId -gt $after } | ForEach-Object {
  $v = @($_.Properties | ForEach-Object { [string]$_.Value })
  if ($v.Count -le 4 -or $v[4].Trim() -ine 'EPSON L395 Series') { return }
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
