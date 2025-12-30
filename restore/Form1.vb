Imports System.IO
Imports System.IO.Compression
Imports System.Diagnostics
Imports System.Text
Imports System.Drawing
Imports System.Security.Cryptography.X509Certificates
Imports System.Runtime.InteropServices

Public Class Form1

    Private cancelar As Boolean = False
    Private pastaTemp As String = ""
    Private driversInstalados As String = ""
    Private RestoreMode As Boolean = False
    Private SilentMode As Boolean = False
    Private atualizandoSelecao As Boolean = False


    ' =========================
    ' FORM LOAD
    ' =========================
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.AllowDrop = True

        With ListView1
            .View = View.Details
            .CheckBoxes = True
            .FullRowSelect = True
            .GridLines = True
            .Columns.Clear()
            .Columns.Add("Dispositivo", 300)
            .Columns.Add("INF", 200)
            .Columns.Add("Status", 220)
        End With

        BTNrestaurarDrivers.Enabled = False
        btncancelar.Visible = False
        ProgressBar1.Value = 0
        LBLstatus.Text = "Aguardando backup..."

        ' DEBUG VISUAL
        txtdebug.Clear()
        txtdebug.Visible = False
        txtdebug.BackColor = Color.Black
        txtdebug.ForeColor = Color.Lime
        txtdebug.Font = New Font("Consolas", 9.0F)

        Log("Aplicação iniciada")

        ' Enumerar drivers instalados
        Dim enumResult = ExecutarPnputil("/enum-drivers")
        driversInstalados = enumResult.Output
        Log("Drivers instalados enumerados")

        ' Processar argumentos de linha de comando
        Dim args = Environment.GetCommandLineArgs()
        For i As Integer = 1 To args.Length - 1
            Select Case args(i).ToLower()
                Case "/restore"
                    RestoreMode = True
                Case "/silent"
                    SilentMode = True
                Case Else
                    If File.Exists(args(i)) Then
                        Log("Backup recebido por argumento: " & args(i))
                        CarregarBackup(args(i))
                    End If
            End Select
        Next

        ' Se veio /restore e backup válido, iniciar restauração automaticamente
        If RestoreMode AndAlso ListView1.Items.Count > 0 Then
            BTNrestaurarDrivers_Click(Nothing, Nothing)
        End If

    End Sub

    ' =========================
    ' DEBUG
    ' =========================
    Private Sub CHKdebug_CheckedChanged(sender As Object, e As EventArgs) Handles chkdebug.CheckedChanged
        txtdebug.Visible = chkdebug.Checked
        Me.Height = If(chkdebug.Checked, 564, 489)
    End Sub

    Private Sub Log(msg As String)
        If Not chkdebug.Checked Then Exit Sub
        txtdebug.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}")
    End Sub

    ' =========================
    ' BOTÃO ABRIR BACKUP
    ' =========================
    Private Sub BTNabrirbackup_Click(sender As Object, e As EventArgs) Handles BTNabrirbackup.Click
        Using ofd As New OpenFileDialog()
            ofd.Filter = "Backup de Drivers (*.drvbackup)|*.drvbackup"
            ofd.Title = "Abrir backup de drivers"

            If ofd.ShowDialog() = DialogResult.OK Then
                Log("Backup selecionado manualmente: " & ofd.FileName)
                CarregarBackup(ofd.FileName)
            End If
        End Using
    End Sub

    ' =========================
    ' DRAG & DROP
    ' =========================
    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        Dim files = CType(e.Data.GetData(DataFormats.FileDrop), String())
        If files.Length > 0 AndAlso File.Exists(files(0)) Then
            Log("Drag & Drop: " & files(0))
            CarregarBackup(files(0))
        End If
    End Sub

    ' =========================
    ' CARREGAR BACKUP
    ' =========================
    Private Sub CarregarBackup(arquivo As String)
        Try
            If Not File.Exists(arquivo) Then Throw New FileNotFoundException("Arquivo de backup não encontrado.")

            ListView1.Items.Clear()
            ProgressBar1.Value = 0
            BTNrestaurarDrivers.Enabled = False

            pastaTemp = Path.Combine(Path.GetTempPath(), "DriverRestore_" & Guid.NewGuid().ToString())
            Directory.CreateDirectory(pastaTemp)
            Log("Extraindo backup para: " & pastaTemp)

            ZipFile.ExtractToDirectory(arquivo, pastaTemp)

            Dim jsonPath = Path.Combine(pastaTemp, "drivers.json")
            If Not File.Exists(jsonPath) Then Throw New FileNotFoundException("drivers.json não encontrado no backup.")

            Dim json = File.ReadAllText(jsonPath, Encoding.UTF8)
            Log("drivers.json carregado")

            For Each bloco In json.Split("{"c)
                If Not bloco.Contains("""inf""") Then Continue For
                Dim device = Extrair(bloco, "device")
                Dim inf = Extrair(bloco, "inf")
                Dim platform = Extrair(bloco, "platform")

                Log($"Driver encontrado: {device} ({inf}) Plataforma: {platform}")

                Dim item As New ListViewItem(device)
                item.SubItems.Add(inf)
                item.SubItems.Add("")
                item.Tag = platform ' armazenar plataforma ou outras info para tooltips

                ListView1.Items.Add(item)
                AnalisarDriver(item)
                item.Checked = True
            Next

            BTNrestaurarDrivers.Enabled = ListView1.Items.Count > 0
            LBLstatus.Text = $"Drivers carregados: {ListView1.Items.Count}"

        Catch ex As Exception
            Log("ERRO CarregarBackup:")
            Log(ex.ToString())
            If Not SilentMode Then
                MessageBox.Show(ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Try
    End Sub

    ' =========================
    ' ANALISAR DRIVER
    ' =========================
    Private Sub AnalisarDriver(item As ListViewItem)
        Dim inf = item.SubItems(1).Text.ToLower()
        Log("Analisando driver: " & inf)

        ' Verificar plataforma 32/64 bits
        Dim driverPlatform = If(item.Tag?.ToString(), "")
        Dim sistema = If(Environment.Is64BitOperatingSystem, "x64", "x86")
        If Not String.IsNullOrEmpty(driverPlatform) AndAlso Not driverPlatform.ToLower().Contains(sistema.ToLower()) Then
            item.SubItems(2).Text = "Plataforma incompatível"
            item.BackColor = Color.LightPink
            Log("ERRO: driver não compatível com a plataforma")
            Exit Sub
        End If

        If driversInstalados.ToLower().Contains(inf) Then
            item.SubItems(2).Text = "Já instalado"
            item.BackColor = Color.LightBlue
            item.Checked = False
            Log("→ Já instalado")
            Exit Sub
        End If

        Dim pastaDriver = Path.Combine(pastaTemp, inf)
        If Not Directory.Exists(pastaDriver) Then
            item.SubItems(2).Text = "Pasta ausente"
            item.BackColor = Color.LightCoral
            Log("ERRO: pasta do driver não encontrada")
            Exit Sub
        End If

        Dim cats = Directory.GetFiles(pastaDriver, "*.cat", SearchOption.AllDirectories)
        If cats.Length = 0 Then
            item.SubItems(2).Text = "Sem assinatura"
            item.BackColor = Color.Orange
            Log("AVISO: sem arquivo .cat")
            Exit Sub
        End If

        If Not CertificadoValido(cats(0)) Then
            item.SubItems(2).Text = "Assinatura inválida"
            item.BackColor = Color.OrangeRed
            Exit Sub
        End If

        item.SubItems(2).Text = "Pronto para restaurar"
        item.BackColor = Color.White
        Log("OK para restauração")
    End Sub

    ' =========================
    ' RESTAURAR / DRY RUN
    ' =========================
    Private Sub BTNrestaurarDrivers_Click(sender As Object, e As EventArgs) Handles BTNrestaurarDrivers.Click
        If ListView1.CheckedItems.Count = 0 Then
            If Not SilentMode Then
                MessageBox.Show("Selecione ao menos um driver.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            Exit Sub
        End If

        cancelar = False
        btncancelar.Visible = Not SilentMode
        BTNrestaurarDrivers.Enabled = False

        ProgressBar1.Maximum = ListView1.CheckedItems.Count
        ProgressBar1.Value = 0

        For Each item As ListViewItem In ListView1.CheckedItems
            If cancelar Then Exit For

            Dim inf = item.SubItems(1).Text
            Dim pastaDriver = Path.Combine(pastaTemp, inf)

            item.SubItems(2).Text = "Processando..."
            item.BackColor = Color.Khaki
            Application.DoEvents()

            Log("Executando pnputil para: " & inf)

            If CHKdryrun.Checked Then
                item.SubItems(2).Text = "Dry Run — OK"
                item.BackColor = Color.LightGreen
                Log("Dry run executado")
            Else
                Dim r = ExecutarPnputil($"/add-driver ""{pastaDriver}\*.inf"" /install")

                Log($"pnputil ExitCode: {r.ExitCode}")
                If Not String.IsNullOrWhiteSpace(r.Output) Then Log("STDOUT:" & Environment.NewLine & r.Output)
                If Not String.IsNullOrWhiteSpace(r.StdErr) Then Log("STDERR:" & Environment.NewLine & r.StdErr)

                If r.ExitCode = 0 Then
                    item.SubItems(2).Text = "Instalado com sucesso"
                    item.BackColor = Color.LightGreen
                Else
                    item.SubItems(2).Text = $"Erro (ExitCode {r.ExitCode})"
                    item.BackColor = Color.LightCoral
                End If
            End If

            ProgressBar1.Value += 1
        Next

        btncancelar.Visible = False
        BTNrestaurarDrivers.Enabled = True
        LBLstatus.Text = "Processo finalizado"

        GerarRelatorio()
    End Sub

    ' =========================
    ' CANCELAR
    ' =========================
    Private Sub BTNcancelar_Click(sender As Object, e As EventArgs) Handles btncancelar.Click
        cancelar = True
        Log("Cancelamento solicitado")
        LBLstatus.Text = "Cancelando..."
    End Sub

    ' =========================
    ' UTILITÁRIOS
    ' =========================
    Private Function ExecutarPnputil(args As String) As (ExitCode As Integer, Output As String, StdErr As String)
        Dim p As New Process
        p.StartInfo.FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "pnputil.exe")
        p.StartInfo.Arguments = args
        p.StartInfo.UseShellExecute = False
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.RedirectStandardError = True
        p.StartInfo.CreateNoWindow = True

        p.Start()
        Dim stdout = p.StandardOutput.ReadToEnd()
        Dim stderr = p.StandardError.ReadToEnd()
        p.WaitForExit()
        Return (p.ExitCode, stdout, stderr)
    End Function

    Private Function CertificadoValido(catPath As String) As Boolean
        Try
            Log("Validando certificado: " & Path.GetFileName(catPath))
            Dim cert As New X509Certificate2(catPath)
            Log("  Subject: " & cert.Subject)
            Log("  Issuer : " & cert.Issuer)
            Log("  Válido de: " & cert.NotBefore)
            Log("  Válido até: " & cert.NotAfter)

            If DateTime.Now < cert.NotBefore OrElse DateTime.Now > cert.NotAfter Then
                Log("  ❌ Certificado fora do período de validade")
                Return False
            End If

            Dim chain As New X509Chain()
            chain.ChainPolicy.RevocationMode = X509RevocationMode.Online
            chain.ChainPolicy.RevocationFlag = X509RevocationFlag.EntireChain

            If Not chain.Build(cert) Then
                Log("  ❌ Falha na cadeia de certificação:")
                For Each s In chain.ChainStatus
                    Log("    - " & s.StatusInformation.Trim())
                Next
                Return False
            End If

            Log("  ✔ Certificado válido")
            Return True
        Catch ex As Exception
            Log("  ❌ Erro ao validar certificado: " & ex.Message)
            Return False
        End Try
    End Function


    Private Function Extrair(bloco As String, chave As String) As String
        Dim chaveBusca = $"""{chave}"""
        Dim pos = bloco.IndexOf(chaveBusca)
        If pos = -1 Then Return ""
        pos = bloco.IndexOf(":", pos)
        If pos = -1 Then Return ""
        Dim inicio = bloco.IndexOf("""", pos + 1)
        If inicio = -1 Then Return ""
        Dim fim = bloco.IndexOf("""", inicio + 1)
        If fim = -1 Then Return ""
        Return bloco.Substring(inicio + 1, fim - inicio - 1)
    End Function

    ' =========================
    ' GERAR RELATÓRIO HTML5 COM BOOTSTRAP E BADGES
    ' =========================
    Private Sub GerarRelatorio()
        Try
            Dim sb As New StringBuilder()
            sb.AppendLine("<!DOCTYPE html>")
            sb.AppendLine("<html lang='pt-br'>")
            sb.AppendLine("<head>")
            sb.AppendLine("<meta charset='UTF-8'>")
            sb.AppendLine("<meta name='viewport' content='width=device-width, initial-scale=1'>")
            sb.AppendLine("<link href='https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css' rel='stylesheet'>")
            sb.AppendLine("<title>Driver Restore Report</title>")
            sb.AppendLine("</head>")
            sb.AppendLine("<body class='p-4'>")
            sb.AppendLine("<div class='container'>")
            sb.AppendLine("<h2 class='mb-4'>Relatório de Restauração de Drivers</h2>")
            sb.AppendLine("<table class='table table-bordered table-striped'>")
            sb.AppendLine("<thead class='table-dark'><tr><th>Dispositivo</th><th>INF</th><th>Status</th></tr></thead>")
            sb.AppendLine("<tbody>")

            For Each item As ListViewItem In ListView1.Items
                Dim statusText = item.SubItems(2).Text
                Dim detalhes = item.Tag?.ToString() ' Tooltip / info extra
                Dim statusBadge As String = $"<span class='badge bg-warning text-dark'>{statusText}</span>"

                If statusText.ToLower().Contains("sucesso") OrElse statusText.ToLower().Contains("ok") Then
                    statusBadge = $"<span class='badge bg-success'>{statusText}</span>"
                ElseIf statusText.ToLower().Contains("erro") Then
                    statusBadge = $"<span class='badge bg-danger'>{statusText}</span>"
                End If

                sb.AppendLine($"<tr><td title='{detalhes}'>{item.Text}</td><td>{item.SubItems(1).Text}</td><td>{statusBadge}</td></tr>")
            Next

            sb.AppendLine("</tbody></table></div></body></html>")

            Dim reportPath As String = Path.Combine(Path.GetTempPath(), "DriverRestore_Report.html")
            File.WriteAllText(reportPath, sb.ToString(), Encoding.UTF8)
            Log("Relatório gerado em: " & reportPath)

            If Not SilentMode Then
                Dim result As DialogResult = MessageBox.Show("Deseja abrir o relatório no navegador padrão?", "Abrir relatório", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.Yes Then
                    Process.Start(New ProcessStartInfo With {.FileName = reportPath, .UseShellExecute = True})
                End If
            End If

        Catch ex As Exception
            Log("Erro ao gerar relatório: " & ex.Message)
            If Not SilentMode Then
                MessageBox.Show("Erro ao gerar relatório: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Try
    End Sub

    ' =========================
    ' LIMPAR PASTA TEMP AO FECHAR
    ' =========================
    Private Sub LimparPastaTemp()
        If String.IsNullOrEmpty(pastaTemp) OrElse Not Directory.Exists(pastaTemp) Then Return

        Try
            For i = 1 To 5
                Try
                    Directory.Delete(pastaTemp, True)
                    Log("Pasta temporária limpa: " & pastaTemp)
                    Exit For
                Catch ex As IOException
                    Log("Tentativa " & i & " de apagar a pasta falhou, retrying...")
                    Threading.Thread.Sleep(200)
                End Try
            Next
        Catch ex As Exception
            Log("Erro ao limpar pasta temporária: " & ex.Message)
        End Try
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        LimparPastaTemp()
    End Sub

    Private Sub CHKSelecionarTodos_CheckedChanged(sender As Object, e As EventArgs) Handles CHKSelecionarTodos.CheckedChanged
        If atualizandoSelecao Then Exit Sub

        atualizandoSelecao = True

        For Each item As ListViewItem In ListView1.Items
            ' Só mexe nos que podem ser restaurados
            If item.SubItems(2).Text = "Pronto para restaurar" _
               OrElse item.SubItems(2).Text.StartsWith("Dry Run") Then
                item.Checked = CHKSelecionarTodos.Checked
            End If
        Next

        atualizandoSelecao = False
    End Sub

    Private Sub ListView1_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles ListView1.ItemChecked
        If atualizandoSelecao Then Exit Sub

        atualizandoSelecao = True

        ' Conta apenas os itens válidos
        Dim totalValidos = ListView1.Items.
            Cast(Of ListViewItem)().
            Count(Function(i) i.SubItems(2).Text = "Pronto para restaurar" _
                               OrElse i.SubItems(2).Text.StartsWith("Dry Run"))

        Dim marcadosValidos = ListView1.CheckedItems.
            Cast(Of ListViewItem)().
            Count(Function(i) i.SubItems(2).Text = "Pronto para restaurar" _
                               OrElse i.SubItems(2).Text.StartsWith("Dry Run"))

        CHKSelecionarTodos.Checked = (totalValidos > 0 AndAlso totalValidos = marcadosValidos)

        atualizandoSelecao = False
    End Sub

End Class
