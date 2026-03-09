using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace InventarioMaquina
{
    public class MainForm : Form
    {
        private TextBox textNome = null!;
        private TextBox textMaquina = null!;
        private TextBox textSetor = null!;
        private TextBox textTipoDisco = null!;
        private TextBox textDisco = null!;
        private TextBox textModelo = null!;
        private TextBox textFabricante = null!;
        private TextBox textSerial = null!;
        private TextBox textProcessador = null!;
        private TextBox textRam = null!;
        private TextBox textAv = null!;
        private TextBox textVersaoWin = null!;
        private TextBox textLicencaAuto = null!;
        private TextBox textLicencaManual = null!;
        private TextBox textSis = null!;
        private ComboBox comboConta = null!;
        private TextBox textContaEmail = null!;
        private TextBox textObs = null!;
        private Button btnAdicionarExcel = null!;
        private Button btnLimpar = null!;
        private PictureBox pictureLogo = null!;
        private TableLayoutPanel table = null!;
        private Panel panelPrincipal = null!;

        public MainForm()
        {
            InicializarTela();
            PreencherCarregando();
            _ = CarregarDadosAsync();
        }

        private void InicializarTela()
        {
            Text = "Inventário de Máquina";
            Size = new Size(1180, 860);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            MinimumSize = new Size(1000, 760);
            BackColor = Color.WhiteSmoke;

            Font fonte = new Font("Segoe UI", 10);

            panelPrincipal = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.WhiteSmoke
            };

            Controls.Add(panelPrincipal);

            pictureLogo = new PictureBox
            {
                Size = new Size(450, 150),
                Location = new Point(350, 20),
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.Transparent
            };

            string caminhoLogo = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "logo.png");
            if (File.Exists(caminhoLogo))
            {
                try
                {
                    pictureLogo.Image = Image.FromFile(caminhoLogo);
                }
                catch
                {
                }
            }

            panelPrincipal.Controls.Add(pictureLogo);

            var titulo = new Label
            {
                Text = "Inventário de Máquina",
                Font = new Font("Segoe UI", 18, FontStyle.Bold),
                ForeColor = Color.FromArgb(35, 35, 35),
                AutoSize = true,
                Location = new Point(420, 170)
            };

            panelPrincipal.Controls.Add(titulo);

            var subtitulo = new Label
            {
                Text = "Levantamento técnico de hardware e software",
                Font = new Font("Segoe UI", 10, FontStyle.Regular),
                ForeColor = Color.DimGray,
                AutoSize = true,
                Location = new Point(420, 205)
            };

            panelPrincipal.Controls.Add(subtitulo);

            table = new TableLayoutPanel
            {
                ColumnCount = 2,
                RowCount = 18,
                Location = new Point(20, 250),
                AutoSize = true,
                Padding = new Padding(14),
                BackColor = Color.White,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
            };

            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 240));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            panelPrincipal.Controls.Add(table);

            textNome = CriarTextBox(fonte);
            textMaquina = CriarTextBox(fonte, true);
            textSetor = CriarTextBox(fonte);
            textTipoDisco = CriarTextBox(fonte, true);
            textDisco = CriarTextBox(fonte, true);
            textModelo = CriarTextBox(fonte, true);
            textFabricante = CriarTextBox(fonte, true);
            textSerial = CriarTextBox(fonte, true);
            textProcessador = CriarTextBox(fonte, true);
            textRam = CriarTextBox(fonte, true);
            textAv = CriarTextBox(fonte, true);
            textVersaoWin = CriarTextBox(fonte, true);
            textLicencaAuto = CriarTextBox(fonte, true);
            textLicencaManual = CriarTextBox(fonte);
            textSis = CriarTextBox(fonte);
            textContaEmail = CriarTextBox(fonte);

            comboConta = new ComboBox
            {
                Font = fonte,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 620
            };
            comboConta.Items.AddRange(new string[] { "Google Drive", "OneDrive", "Dropbox", "Outros" });
            comboConta.SelectedIndex = 0;

            textObs = new TextBox
            {
                Font = fonte,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Width = 620,
                Height = 120,
                BorderStyle = BorderStyle.FixedSingle
            };

            int row = 0;
            AdicionarLinha("Nome do Usuário:", textNome, row++);
            AdicionarLinha("Nome da Máquina:", textMaquina, row++);
            AdicionarLinha("Setor:", textSetor, row++);
            AdicionarLinha("Tipo de Disco:", textTipoDisco, row++);
            AdicionarLinha("Total do Disco:", textDisco, row++);
            AdicionarLinha("Modelo da Máquina:", textModelo, row++);
            AdicionarLinha("Fabricante:", textFabricante, row++);
            AdicionarLinha("Serial da Máquina:", textSerial, row++);
            AdicionarLinha("Processador:", textProcessador, row++);
            AdicionarLinha("Memória RAM:", textRam, row++);
            AdicionarLinha("Antivírus:", textAv, row++);
            AdicionarLinha("Versão do Windows:", textVersaoWin, row++);
            AdicionarLinha("Licença:", textLicencaAuto, row++);
            AdicionarLinha("Licença Manual (se tiver):", textLicencaManual, row++);
            AdicionarLinha("Sistemas (se usar algum):", textSis, row++);
            AdicionarLinha("Tipo da Conta:", comboConta, row++);
            AdicionarLinha("E-mail da Conta:", textContaEmail, row++);
            AdicionarLinha("Observações:", textObs, row++);

            btnLimpar = new Button
            {
                Text = "Limpar",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Width = 120,
                Height = 42,
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Location = new Point(720, table.Bottom + 20)
            };
            btnLimpar.FlatAppearance.BorderSize = 0;
            btnLimpar.Click += (s, e) => LimparCampos();
            panelPrincipal.Controls.Add(btnLimpar);

            btnAdicionarExcel = new Button
            {
                Text = "Adicionar ao Excel",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Width = 190,
                Height = 42,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Location = new Point(850, table.Bottom + 20)
            };

            btnAdicionarExcel.FlatAppearance.BorderSize = 0;
            btnAdicionarExcel.Click += BtnAdicionarAoExcel_Click;
            panelPrincipal.Controls.Add(btnAdicionarExcel);
            
            var rodape = new Label
            {
                Text = "Criado por: Thomas Castro",
                Font = new Font("Segoe UI", 8, FontStyle.Italic),
                ForeColor = Color.Gray,
                AutoSize = true
            };

            panelPrincipal.Controls.Add(rodape);

            panelPrincipal.Resize += (s, e) =>
            {
                rodape.Location = new Point(
                    panelPrincipal.ClientSize.Width - rodape.Width - 20,
                    panelPrincipal.ClientSize.Height - 30
                );
            };


            panelPrincipal.Resize += (s, e) =>
            {
                table.Width = Math.Max(960, panelPrincipal.ClientSize.Width - 60);

                btnAdicionarExcel.Location = new Point(
                    Math.Max(20, panelPrincipal.ClientSize.Width - btnAdicionarExcel.Width - 40),
                    table.Bottom + 20
                );

                btnLimpar.Location = new Point(
                    btnAdicionarExcel.Left - btnLimpar.Width - 12,
                    table.Bottom + 20
                );
            };
        }

        private TextBox CriarTextBox(Font fonte, bool somenteLeitura = false)
        {
            return new TextBox
            {
                Font = fonte,
                Width = 620,
                ReadOnly = somenteLeitura,
                BackColor = somenteLeitura ? Color.FromArgb(245, 245, 245) : Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private void AdicionarLinha(string labelText, Control control, int row)
        {
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var lbl = new Label
            {
                Text = labelText,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(6, 12, 6, 12),
                Font = new Font("Segoe UI", 10, FontStyle.Regular),
                ForeColor = Color.FromArgb(50, 50, 50)
            };

            control.Margin = new Padding(6, 8, 6, 8);
            control.Anchor = AnchorStyles.Left | AnchorStyles.Right;

            table.Controls.Add(lbl, 0, row);
            table.Controls.Add(control, 1, row);
        }

        private void PreencherCarregando()
        {
            textMaquina.Text = "Carregando...";
            textTipoDisco.Text = "Carregando...";
            textDisco.Text = "Carregando...";
            textModelo.Text = "Carregando...";
            textFabricante.Text = "Carregando...";
            textSerial.Text = "Carregando...";
            textProcessador.Text = "Carregando...";
            textRam.Text = "Carregando...";
            textAv.Text = "Carregando...";
            textVersaoWin.Text = "Carregando...";
            textLicencaAuto.Text = "Carregando...";
        }

        private async Task CarregarDadosAsync()
        {
            string maquina = Environment.MachineName;

            var taskTipoDisco = Task.Run(() => GetTipoDiscoSistema());
            var taskDisco = Task.Run(() => GetDiscoC());
            var taskModelo = Task.Run(() => GetModelo());
            var taskFabricante = Task.Run(() => GetFabricante());
            var taskSerial = Task.Run(() => GetSerialMaquina());
            var taskProcessador = Task.Run(() => GetProcessador());
            var taskRam = Task.Run(() => GetRAM());
            var taskAv = Task.Run(() => GetAntivirus());
            var taskVersao = Task.Run(() => GetVersaoWindows());
            var taskLicenca = Task.Run(() => GetLicencaWindows());

            textMaquina.Text = ValorOuPadrao(maquina);
            textTipoDisco.Text = ValorOuPadrao(await taskTipoDisco);
            textDisco.Text = ValorOuPadrao(await taskDisco);
            textModelo.Text = ValorOuPadrao(await taskModelo);
            textFabricante.Text = ValorOuPadrao(await taskFabricante);
            textSerial.Text = ValorOuPadrao(await taskSerial);
            textProcessador.Text = ValorOuPadrao(await taskProcessador);
            textRam.Text = ValorOuPadrao(await taskRam);
            textAv.Text = ValorOuPadrao(await taskAv);
            textVersaoWin.Text = ValorOuPadrao(await taskVersao);
            textLicencaAuto.Text = ValorOuPadrao(await taskLicenca);
        }

        private string ValorOuPadrao(string valor)
        {
            return string.IsNullOrWhiteSpace(valor) ? "Não detectado" : valor.Trim();
        }

        private string GetModelo()
        {
            return ExecutarPowerShell("(Get-CimInstance Win32_ComputerSystem).Model");
        }

        private string GetFabricante()
        {
            return ExecutarPowerShell("(Get-CimInstance Win32_ComputerSystem).Manufacturer");
        }

        private string GetSerialMaquina()
        {
            string result = ExecutarPowerShell("(Get-CimInstance Win32_BIOS).SerialNumber");
            return string.IsNullOrWhiteSpace(result) ? "Não detectado" : result;
        }

        private string GetVersaoWindows()
        {
            return ExecutarPowerShell("(Get-CimInstance Win32_OperatingSystem).Caption");
        }

        private string GetDiscoC()
        {
            string result = ExecutarPowerShell(
                "$d = Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DeviceID -eq 'C:' }; " +
                "if($d){ [math]::Round($d.Size / 1GB, 0).ToString() + ' GB' } else { 'Não detectado' }"
            );

            return string.IsNullOrWhiteSpace(result) ? "Não detectado" : result;
        }

        private string GetRAM()
        {
            string result = ExecutarPowerShell(
                "$c = Get-CimInstance Win32_ComputerSystem; " +
                "if($c){ [math]::Round($c.TotalPhysicalMemory / 1GB, 2).ToString() + ' GB' }"
            );

            return string.IsNullOrWhiteSpace(result) ? "Não detectado" : result;
        }

        private string GetProcessador()
        {
            string result = ExecutarPowerShell(
                "(Get-CimInstance Win32_Processor | Select-Object -First 1).Name"
            );

            return string.IsNullOrWhiteSpace(result) ? "Não detectado" : result;
        }

        private string GetTipoDiscoSistema()
        {
            string result = ExecutarPowerShell(
                "$p = Get-Partition -DriveLetter C -ErrorAction SilentlyContinue; " +
                "if($p){ " +
                "   $d = Get-Disk -Number $p.DiskNumber -ErrorAction SilentlyContinue; " +
                "   if($d){ " +
                "       if($d.MediaType -eq 'SSD'){ 'SSD' } " +
                "       elseif($d.MediaType -eq 'HDD'){ 'HD' } " +
                "       elseif($d.FriendlyName -match 'SSD|NVMe'){ 'SSD' } " +
                "       elseif($d.FriendlyName -match 'HDD|SATA'){ 'HD' } " +
                "       else { 'Não detectado' } " +
                "   } else { 'Não detectado' } " +
                "} else { 'Não detectado' }"
            );

            return string.IsNullOrWhiteSpace(result) ? "Não detectado" : result;
        }

        private string GetAntivirus()
        {
            string result = ExecutarPowerShell(
                "try { " +
                "$a = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct -ErrorAction Stop; " +
                "if($a){ ($a.displayName -join '; ') } else { 'Não detectado' } " +
                "} catch { 'Não detectado' }"
            );

            return string.IsNullOrWhiteSpace(result) ? "Não detectado" : result;
        }

        private string GetLicencaWindows()
        {
            string result = ExecutarPowerShell(
                "try { " +
                "$os = Get-CimInstance SoftwareLicensingProduct | " +
                "Where-Object { $_.ApplicationID -eq '55c92734-d682-4d71-983e-d6ec3f16059f' -and $_.PartialProductKey } | " +
                "Select-Object -First 1; " +
                "if($os){ " +
                "switch ($os.LicenseStatus) { " +
                "0 { 'Unlicensed' } " +
                "1 { 'Licensed' } " +
                "2 { 'Initial grace period' } " +
                "3 { 'Additional grace period' } " +
                "4 { 'Non-genuine grace period' } " +
                "5 { 'Notification' } " +
                "6 { 'Extended grace period' } " +
                "default { 'Desconhecido' } " +
                "} " +
                "} else { 'Não detectado' } " +
                "} catch { 'Não detectado' }"
            );

            return string.IsNullOrWhiteSpace(result) ? "Não detectado" : result;
        }

        private string ExecutarPowerShell(string comando)
        {
            try
            {
                string comandoFinal =
                    "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; " +
                    "$OutputEncoding = [System.Text.Encoding]::UTF8; " +
                    comando;

                using Process process = new Process();
                process.StartInfo.FileName = "powershell.exe";
                process.StartInfo.Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{comandoFinal}\"";
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.StandardOutputEncoding = Encoding.UTF8;
                process.StartInfo.StandardErrorEncoding = Encoding.UTF8;

                process.Start();

                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();

                process.WaitForExit();

                if (!string.IsNullOrWhiteSpace(output))
                    return output.Trim();

                if (!string.IsNullOrWhiteSpace(error))
                    return "Não detectado";

                return "Não detectado";
            }
            catch
            {
                return "Não detectado";
            }
        }

        private void BtnAdicionarAoExcel_Click(object? sender, EventArgs e)
        {
            try
            {
                string pasta = AppDomain.CurrentDomain.BaseDirectory;
                string arquivo = Path.Combine(pasta, "InventarioMaquinas.xlsx");

                bool arquivoExiste = File.Exists(arquivo);

                using XLWorkbook workbook = arquivoExiste ? new XLWorkbook(arquivo) : new XLWorkbook();
                var ws = arquivoExiste
                    ? workbook.Worksheet("Inventário")
                    : workbook.Worksheets.Add("Inventário");

                if (!arquivoExiste)
                {
                    CriarCabecalhoInventario(ws);
                }

                int proximaLinha = GetProximaLinha(ws);

                ws.Cell(proximaLinha, 1).Value = textNome.Text;
                ws.Cell(proximaLinha, 2).Value = textMaquina.Text;
                ws.Cell(proximaLinha, 3).Value = textSetor.Text;
                ws.Cell(proximaLinha, 4).Value = textTipoDisco.Text;
                ws.Cell(proximaLinha, 5).Value = textDisco.Text;
                ws.Cell(proximaLinha, 6).Value = textModelo.Text;
                ws.Cell(proximaLinha, 7).Value = textFabricante.Text;
                ws.Cell(proximaLinha, 8).Value = textSerial.Text;
                ws.Cell(proximaLinha, 9).Value = textProcessador.Text;
                ws.Cell(proximaLinha, 10).Value = textRam.Text;
                ws.Cell(proximaLinha, 11).Value = textAv.Text;
                ws.Cell(proximaLinha, 12).Value = textVersaoWin.Text;
                ws.Cell(proximaLinha, 13).Value = textLicencaAuto.Text;
                ws.Cell(proximaLinha, 14).Value = textLicencaManual.Text;
                ws.Cell(proximaLinha, 15).Value = textSis.Text;
                ws.Cell(proximaLinha, 16).Value = comboConta.SelectedItem?.ToString() ?? "";
                ws.Cell(proximaLinha, 17).Value = textContaEmail.Text;
                ws.Cell(proximaLinha, 18).Value = textObs.Text;

                AplicarBordasLinha(ws, proximaLinha, 18);

                ws.Columns().AdjustToContents();

                ws.Column(1).Width = 20;   // Nome
                ws.Column(2).Width = 20;   // Hostname
                ws.Column(3).Width = 22;   // Setor
                ws.Column(4).Width = 16;   // Tipo de disco
                ws.Column(5).Width = 16;   // Total de disco
                ws.Column(6).Width = 22;   // Modelo
                ws.Column(7).Width = 18;   // Fabricante
                ws.Column(8).Width = 18;   // Serial
                ws.Column(9).Width = 34;   // Processador
                ws.Column(10).Width = 14;  // RAM
                ws.Column(11).Width = 30;  // Antivírus
                ws.Column(12).Width = 30;  // Windows
                ws.Column(13).Width = 14;  // Licença
                ws.Column(14).Width = 22;  // Chave
                ws.Column(15).Width = 24;  // Sistemas
                ws.Column(16).Width = 18;  // Tipo da conta
                ws.Column(17).Width = 28;  // E-mail
                ws.Column(18).Width = 30;  // Observações

                workbook.SaveAs(arquivo);

                MessageBox.Show(
                    $"Máquina adicionada ao inventário com sucesso!\n\nArquivo: {arquivo}",
                    "Sucesso",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                LimparCampos();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Erro ao adicionar ao Excel: " + ex.Message,
                    "Erro",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private void CriarCabecalhoInventario(IXLWorksheet ws)
        {
            ws.Cell("A1").Value = "INVENTÁRIO DE EQUIPAMENTOS TÉCNICOS";
            ws.Range("A1:R1").Merge();
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 16;
            ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EAF7");

            string caminhoLogo = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "logo.png");
            if (File.Exists(caminhoLogo))
            {
                try
                {
                    ws.AddPicture(caminhoLogo).MoveTo(ws.Cell("S1")).Scale(0.30);
                }
                catch
                {
                }
            }

            ws.Cell(2, 1).Value = "NOME";
            ws.Cell(2, 2).Value = "HOSTNAME";
            ws.Cell(2, 3).Value = "SETOR";
            ws.Cell(2, 4).Value = "TIPO DE DISCO";
            ws.Cell(2, 5).Value = "TOTAL DE DISCO";
            ws.Cell(2, 6).Value = "MODELO";
            ws.Cell(2, 7).Value = "FABRICANTE";
            ws.Cell(2, 8).Value = "SERIAL";
            ws.Cell(2, 9).Value = "PROCESSADOR";
            ws.Cell(2, 10).Value = "MEMÓRIA RAM";
            ws.Cell(2, 11).Value = "ANTIVÍRUS";
            ws.Cell(2, 12).Value = "VERSÃO WINDOWS";
            ws.Cell(2, 13).Value = "LICENÇA";
            ws.Cell(2, 14).Value = "CHAVE DA LICENÇA";
            ws.Cell(2, 15).Value = "SISTEMAS";
            ws.Cell(2, 16).Value = "TIPO DA CONTA";
            ws.Cell(2, 17).Value = "E-MAIL DA CONTA";
            ws.Cell(2, 18).Value = "OBSERVAÇÕES";

            var cabecalho = ws.Range("A2:R2");
            cabecalho.Style.Font.Bold = true;
            cabecalho.Style.Font.FontColor = XLColor.White;
            cabecalho.Style.Fill.BackgroundColor = XLColor.Black;
            cabecalho.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cabecalho.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cabecalho.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cabecalho.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.SheetView.FreezeRows(2);
            ws.Range("A2:R2").SetAutoFilter();
        }

        private int GetProximaLinha(IXLWorksheet ws)
        {
            int ultimaLinhaUsada = ws.LastRowUsed()?.RowNumber() ?? 2;
            return ultimaLinhaUsada < 3 ? 3 : ultimaLinhaUsada + 1;
        }

        private void AplicarBordasLinha(IXLWorksheet ws, int linha, int ultimaColuna)
        {
            var range = ws.Range(linha, 1, linha, ultimaColuna);
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private void LimparCampos()
        {
            textNome.Clear();
            textSetor.Clear();
            textLicencaManual.Clear();
            textSis.Clear();
            textContaEmail.Clear();
            textObs.Clear();
            comboConta.SelectedIndex = 0;
        }
    }
}