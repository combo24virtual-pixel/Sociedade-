
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;

namespace MonitorEpsonL395
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MonitorForm());
        }
    }

    public class Registro
    {
        public long Id { get; set; }
        public DateTime Data { get; set; }
        public string Documento { get; set; } = "";
        public string Impressora { get; set; } = "";
        public int Paginas { get; set; }
        public string Papel { get; set; } = "Não identificado";
        public decimal Valor { get; set; }
    }

    public class DadosSalvos
    {
        public long UltimoEvento { get; set; }
        public List<Registro> Registros { get; set; } = new();
    }

    public class EventoImpressao
    {
        public long Id { get; set; }
        public DateTime Data { get; set; }
        public string Documento { get; set; } = "";
        public string Impressora { get; set; } = "";
        public int Paginas { get; set; }
    }

    public class MonitorForm : Form
    {
        private readonly string pasta;
        private readonly string arquivo;
        private readonly DataGridView tabela;
        private readonly Label resumo;
        private readonly Label situacao;
        private readonly NotifyIcon icone;
        private readonly System.Windows.Forms.Timer relogio;

        private DadosSalvos dados = new();
        private bool encerrando;
        private bool consultando;

        public MonitorForm()
        {
            Text = "Monitor Epson L395";
            Width = 1050;
            Height = 650;
            MinimumSize = new Size(750, 450);
            StartPosition = FormStartPosition.CenterScreen;

            pasta = Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "MonitorEpsonL395");

            Directory.CreateDirectory(pasta);
            arquivo = Path.Combine(pasta, "historico.json");

            var painel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 52,
                Padding = new Padding(6),
                AutoScroll = true,
                WrapContents = false
            };

            AdicionarBotao(painel, "Atualizar", () => Atualizar());
            AdicionarBotao(painel, "Exportar CSV mensal", ExportarCSV);
            AdicionarBotao(painel, "Zerar histórico", ZerarHistorico);
            AdicionarBotao(painel, "Verificar registro", VerificarRegistro);
            AdicionarBotao(painel, "Abrir pasta de dados", AbrirPasta);
            AdicionarBotao(painel, "Sair", Sair);

            resumo = new Label
            {
                Dock = DockStyle.Top,
                Height = 42,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(12, 0, 0, 0),
                Font = new Font("Segoe UI", 11, FontStyle.Bold)
            };

            situacao = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                Padding = new Padding(12, 5, 0, 0),
                Text = "Iniciando monitoramento..."
            };

            tabela = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode =
                    DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode =
                    DataGridViewSelectionMode.FullRowSelect
            };

            tabela.Columns.Add("Data", "Data e hora");
            tabela.Columns.Add("Documento", "Documento");
            tabela.Columns.Add("Impressora", "Impressora");
            tabela.Columns.Add("Paginas", "Páginas");
            tabela.Columns.Add("Papel", "Tipo de papel");
            tabela.Columns.Add("Valor", "Valor (R$)");

            Controls.Add(tabela);
            Controls.Add(situacao);
            Controls.Add(resumo);
            Controls.Add(painel);

            icone = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = "Monitor Epson L395",
                Visible = true
            };

            var menu = new ContextMenuStrip();
            menu.Items.Add("Abrir monitor", null,
                (s, e) => MostrarJanela());
            menu.Items.Add("Sair", null,
                (s, e) => Sair());

            icone.ContextMenuStrip = menu;
            icone.DoubleClick += (s, e) => MostrarJanela();

            relogio = new System.Windows.Forms.Timer
            {
                Interval = 5000
            };

            relogio.Tick += (s, e) => Atualizar();

            FormClosing += AoFechar;
            FormClosed += (s, e) => icone.Dispose();

            Carregar();

            // Na primeira execução, ignora impressões anteriores.
            if (!File.Exists(arquivo))
            {
                try
                {
                    var eventos = LerEventos();
                    dados.UltimoEvento = eventos.Count > 0
                        ? eventos.Max(x => x.Id)
                        : 0;
                    Salvar();
                }
                catch (Exception ex)
                {
                    situacao.Text =
                        "Não foi possível iniciar: " + ex.Message;
                }
            }

            MostrarDados();
            relogio.Start();
            Atualizar();
        }

        private void AdicionarBotao(
            FlowLayoutPanel painel,
            string texto,
            Action acao)
        {
            var botao = new Button
            {
                Text = texto,
                AutoSize = true,
                Height = 32,
                Margin = new Padding(3)
            };

            botao.Click += (s, e) => acao();
            painel.Controls.Add(botao);
        }

        private void Carregar()
        {
            if (!File.Exists(arquivo))
                return;

            try
            {
                var conteudo = File.ReadAllText(arquivo);
                dados = JsonSerializer.Deserialize<DadosSalvos>(
                    conteudo) ?? new DadosSalvos();

                dados.Registros ??= new List<Registro>();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Não foi possível carregar o histórico:\n" +
                    ex.Message);
            }
        }

        private void Salvar()
        {
            var json = JsonSerializer.Serialize(
                dados,
                new JsonSerializerOptions
                {
                    WriteIndented = true
                });

            File.WriteAllText(
                arquivo,
                json,
                new UTF8Encoding(false));
        }

        private List<EventoImpressao> LerEventos()
        {
            // Evento 307: documento registrado como impresso
            // pelo serviço de impressão do Windows.
            string script = @"
$ErrorActionPreference = 'Stop'
$lista = @(Get-WinEvent -FilterHashtable @{
    LogName='Microsoft-Windows-PrintService/Operational'
    Id=307
} -MaxEvents 500 -ErrorAction SilentlyContinue)

$resultado = @(
    foreach ($ev in $lista) {
        $xml = [xml]$ev.ToXml()
        $p = @($xml.Event.UserData.DocumentPrinted.Param1,
               $xml.Event.UserData.DocumentPrinted.Param2,
               $xml.Event.UserData.DocumentPrinted.Param3,
               $xml.Event.UserData.DocumentPrinted.Param4,
               $xml.Event.UserData.DocumentPrinted.Param5,
               $xml.Event.UserData.DocumentPrinted.Param6,
               $xml.Event.UserData.DocumentPrinted.Param7,
               $xml.Event.UserData.DocumentPrinted.Param8)

        $nome = [string]$p[4]
        $paginas = 0
        [int]::TryParse([string]$p[7],
            [ref]$paginas) | Out-Null

        if ($nome -like '*EPSON L395*') {
            [PSCustomObject]@{
                Id = [long]$ev.RecordId
                Data = $ev.TimeCreated.ToString('o')
                Documento = [string]$p[1]
                Impressora = $nome
                Paginas = $paginas
            }
        }
    }
)
ConvertTo-Json -InputObject $resultado -Compress -Depth 4
";

            string comando = Convert.ToBase64String(
                Encoding.Unicode.GetBytes(script));

            var inicio = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments =
                    "-NoProfile -NonInteractive -EncodedCommand " +
                    comando,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8
            };

            using var processo = Process.Start(inicio)
                ?? throw new Exception(
                    "Não foi possível abrir o PowerShell.");

            string resposta =
                processo.StandardOutput.ReadToEnd();

            string erros =
                processo.StandardError.ReadToEnd();

            processo.WaitForExit();

            if (processo.ExitCode != 0)
                throw new Exception(erros);

            if (string.IsNullOrWhiteSpace(resposta) ||
                resposta.Trim() == "null")
                return new List<EventoImpressao>();

            return JsonSerializer.Deserialize<
                List<EventoImpressao>>(resposta,
                new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                }) ?? new List<EventoImpressao>();
        }

        private void Atualizar()
        {
            if (consultando)
                return;

            consultando = true;

            try
            {
                var eventos = LerEventos()
                    .OrderBy(x => x.Id)
                    .ToList();

                int novos = 0;

                foreach (var ev in eventos)
                {
                    if (ev.Id <= dados.UltimoEvento)
                        continue;

                    dados.UltimoEvento = ev.Id;

                    if (dados.Registros.Any(x => x.Id == ev.Id))
                        continue;

                    dados.Registros.Add(new Registro
                    {
                        Id = ev.Id,
                        Data = ev.Data,
                        Documento = ev.Documento,
                        Impressora = ev.Impressora,
                        Paginas = Math.Max(0, ev.Paginas),
                        Papel = "Não identificado",
                        Valor = 0m
                    });

                    novos++;
                }

                if (novos > 0)
                    Salvar();

                MostrarDados();

                situacao.Text =
                    "Monitoramento ativo. Última consulta: " +
                    DateTime.Now.ToString("HH:mm:ss") +
                    ". Novos registros: " + novos +
                    ". O tipo de papel não é fornecido " +
                    "pelo evento 307.";
            }
            catch (Exception ex)
            {
                situacao.Text =
                    "Erro na consulta: " + ex.Message;
            }
            finally
            {
                consultando = false;
            }
        }

        private void MostrarDados()
        {
            tabela.Rows.Clear();

            foreach (var r in dados.Registros
                .OrderByDescending(x => x.Data))
            {
                tabela.Rows.Add(
                    r.Data.ToString("dd/MM/yyyy HH:mm:ss"),
                    r.Documento,
                    r.Impressora,
                    r.Paginas,
                    r.Papel,
                    r.Valor.ToString("N2"));
            }

            resumo.Text =
                "Trabalhos: " + dados.Registros.Count +
                "    |    Páginas: " +
                dados.Registros.Sum(x => x.Paginas) +
                "    |    Valor classificado: R$ " +
                dados.Registros.Sum(x => x.Valor)
                    .ToString("N2");
        }

        private void VerificarRegistro()
        {
            Atualizar();

            MessageBox.Show(
                "O monitor consulta os eventos 307 " +
                "do Windows a cada 5 segundos.\n\n" +
                "Para testar, imprima uma página " +
                "na Epson L395 e clique em Atualizar.\n\n" +
                "Somente eventos novos são adicionados.\n\n" +
                "A confirmação do Windows não garante " +
                "que o papel saiu fisicamente.",
                "Verificar registro",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void ZerarHistorico()
        {
            if (MessageBox.Show(
                "Deseja apagar o histórico do monitor?",
                "Confirmar",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            try
            {
                // Mantém a posição do último evento para
                // não recuperar impressões antigas.
                dados.Registros.Clear();

                var eventos = LerEventos();
                if (eventos.Count > 0)
                    dados.UltimoEvento = Math.Max(
                        dados.UltimoEvento,
                        eventos.Max(x => x.Id));

                Salvar();
                MostrarDados();

                situacao.Text =
                    "Histórico zerado. Aguardando " +
                    "novas impressões.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ExportarCSV()
        {
            using var dialogo = new SaveFileDialog
            {
                Filter = "Arquivo CSV (*.csv)|*.csv",
                FileName = "impressoes_" +
                    DateTime.Now.ToString("yyyy_MM") + ".csv"
            };

            if (dialogo.ShowDialog() != DialogResult.OK)
                return;

            string Escapar(string valor)
            {
                return "\"" +
                    (valor ?? "").Replace("\"", "\"\"") +
                    "\"";
            }

            var linhas = new List<string>
            {
                "Data;Documento;Impressora;Paginas;Papel;Valor"
            };

            var registros = dados.Registros.Where(x =>
                x.Data.Year == DateTime.Now.Year &&
                x.Data.Month == DateTime.Now.Month);

            foreach (var r in registros)
            {
                linhas.Add(string.Join(";",
                    Escapar(r.Data.ToString(
                        "dd/MM/yyyy HH:mm:ss")),
                    Escapar(r.Documento),
                    Escapar(r.Impressora),
                    r.Paginas.ToString(),
                    Escapar(r.Papel),
                    r.Valor.ToString(
                        "N2", new CultureInfo("pt-BR"))));
            }

            File.WriteAllLines(
                dialogo.FileName,
                linhas,
                new UTF8Encoding(true));

            MessageBox.Show(
                "Relatório mensal exportado com sucesso.");
        }

        private void AbrirPasta()
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = pasta,
                UseShellExecute = true
            });
        }

        private void MostrarJanela()
        {
            Show();
            WindowState = FormWindowState.Normal;
            Activate();
        }

        private void AoFechar(
            object? sender,
            FormClosingEventArgs e)
        {
            if (encerrando)
                return;

            // O X esconde a janela.
            // O monitor continua funcionando.
            e.Cancel = true;
            Hide();

            icone.BalloonTipTitle = "Monitor Epson L395";
            icone.BalloonTipText =
                "O monitor continua ativo em segundo plano.";
            icone.ShowBalloonTip(2000);
        }

        private void Sair()
        {
            encerrando = true;
            relogio.Stop();
            icone.Visible = false;
            Close();
        }
    }
}
