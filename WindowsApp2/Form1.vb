Imports System.IO
Imports System.IO.Compression
Imports System.Management
Imports System.Diagnostics
Imports System.Text
Imports System.Drawing

Public Class Form1

    Private ignorarEventos As Boolean = False
    Private cancelarBackup As Boolean = False
    Private processoAtual As Process = Nothing
    Private arquivoDestinoBackup As String = Nothing
    Private pastaTempBackup As String = Nothing

    ' ==============================
    ' FORM LOAD
    ' ==============================
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With ListView1
            .View = View.Details
            .CheckBoxes = True
            .FullRowSelect = True
            .GridLines = True
            .Columns.Clear()
            .Columns.Add("Dispositivo", 320)
            .Columns.Add("INF", 180)
            .Columns.Add("Fornecedor", 200)
        End With

        ProgressBar1.Minimum = 0
        ProgressBar1.Value = 0

        btncancelar.Visible = False
        LBLstatus.Text = "Listando drivers..."

        ListarDrivers()

    End Sub

    ' ==============================
    ' LISTAR DRIVERS REAIS
    ' ==============================
    Private Sub ListarDrivers()

        ignorarEventos = True
        CHKSelecionarTodos.Checked = False
        ignorarEventos = False

        ListView1.Items.Clear()

        Dim searcher As New ManagementObjectSearcher(
            "SELECT * FROM Win32_PnPSignedDriver WHERE DeviceClass IS NOT NULL"
        )

        For Each drv As ManagementObject In searcher.Get()

            Dim inf As String = TryCast(drv("InfName"), String)
            Dim deviceName As String = TryCast(drv("DeviceName"), String)
            Dim provider As String = TryCast(drv("DriverProviderName"), String)
            Dim deviceId As String = TryCast(drv("DeviceID"), String)

            If String.IsNullOrEmpty(inf) OrElse String.IsNullOrEmpty(deviceName) Then Continue For

            If deviceId.StartsWith("ROOT\") _
                OrElse deviceId.StartsWith("SWD\") _
                OrElse deviceId.StartsWith("HTREE\") Then Continue For

            If provider = "Microsoft" Then Continue For

            Dim item As New ListViewItem(deviceName)
            item.SubItems.Add(inf)
            item.SubItems.Add(provider)

            ListView1.Items.Add(item)
        Next

        LBLstatus.Text = $"Drivers encontrados: {ListView1.Items.Count}"

    End Sub

    ' ==============================
    ' BACKUP
    ' ==============================
    Private Sub BTNbackup_Click(sender As Object, e As EventArgs) Handles BTNbackup.Click

        If ListView1.CheckedItems.Count = 0 Then
            MessageBox.Show("Selecione ao menos um driver.",
                            "Aviso",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim sfd As New SaveFileDialog With {
            .Title = "Salvar backup de drivers",
            .Filter = "Backup de Drivers (*.drvbackup)|*.drvbackup"
        }

        If sfd.ShowDialog() <> DialogResult.OK Then Exit Sub

        cancelarBackup = False
        arquivoDestinoBackup = sfd.FileName

        Dim tempDir As String = Path.Combine(
            Path.GetTempPath(),
            "DriverBackup_" & Guid.NewGuid().ToString()
        )

        pastaTempBackup = tempDir
        Directory.CreateDirectory(tempDir)

        BTNbackup.Visible = False
        btncancelar.Visible = True

        Dim pnputilPath As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.System),
            "pnputil.exe"
        )

        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = ListView1.CheckedItems.Count
        ProgressBar1.Value = 0

        Dim json As New StringBuilder
        json.AppendLine("[")

        For Each item As ListViewItem In ListView1.CheckedItems

            If cancelarBackup Then Exit For

            Dim inf As String = item.SubItems(1).Text
            item.BackColor = Color.Khaki

            LBLstatus.Text =
                $"Copiando {inf} ({ProgressBar1.Value + 1}/{ProgressBar1.Maximum})"

            Application.DoEvents()

            Try
                Dim dest As String = Path.Combine(tempDir, inf)
                Directory.CreateDirectory(dest)

                processoAtual = New Process
                With processoAtual.StartInfo
                    .FileName = pnputilPath
                    .Arguments = $"/export-driver {inf} ""{dest}"""
                    .UseShellExecute = False
                    .CreateNoWindow = True
                End With

                processoAtual.Start()

                Do While Not processoAtual.HasExited
                    Application.DoEvents()
                    If cancelarBackup Then
                        Try
                            processoAtual.Kill()
                        Catch
                        End Try
                        Exit Do
                    End If
                    Threading.Thread.Sleep(50)
                Loop

                If cancelarBackup Then Exit For

                If processoAtual.ExitCode = 0 Then
                    item.BackColor = Color.LightGreen
                    json.AppendLine("  {")
                    json.AppendLine($"    ""device"": ""{item.Text.Replace("""", "'")}"",")
                    json.AppendLine($"    ""inf"": ""{inf}""")
                    json.AppendLine("  },")
                Else
                    item.BackColor = Color.LightCoral
                End If

            Catch
                item.BackColor = Color.LightCoral
            End Try

            ProgressBar1.Value += 1
        Next

        btncancelar.Visible = False
        BTNbackup.Visible = True

        If cancelarBackup Then

            Try
                If File.Exists(arquivoDestinoBackup) Then File.Delete(arquivoDestinoBackup)
            Catch
            End Try

            Try
                If Directory.Exists(pastaTempBackup) Then Directory.Delete(pastaTempBackup, True)
            Catch
            End Try

            ProgressBar1.Value = 0
            LBLstatus.Text = "Backup cancelado."

            MessageBox.Show("Backup cancelado pelo usuário.",
                            "Cancelado",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Exit Sub
        End If

        If json.ToString().EndsWith("," & Environment.NewLine) Then
            json.Length -= ("," & Environment.NewLine).Length
            json.AppendLine()
        End If

        json.AppendLine("]")

        File.WriteAllText(
            Path.Combine(tempDir, "drivers.json"),
            json.ToString(),
            Encoding.UTF8
        )

        ZipFile.CreateFromDirectory(tempDir, arquivoDestinoBackup)
        Directory.Delete(tempDir, True)

        ProgressBar1.Value = ProgressBar1.Maximum
        LBLstatus.Text = "Backup concluído (100%)"

        MessageBox.Show("Backup de drivers concluído com sucesso!",
                        "Concluído",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

    End Sub

    ' ==============================
    ' CANCELAR
    ' ==============================
    Private Sub BTNcancelar_Click(sender As Object, e As EventArgs) Handles btncancelar.Click

        Dim r = MessageBox.Show(
            "Deseja cancelar o backup em andamento?",
            "Cancelar backup",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )

        If r <> DialogResult.Yes Then Exit Sub

        cancelarBackup = True
        LBLstatus.Text = "Cancelando backup..."

    End Sub

    ' ==============================
    ' CHECKBOX SELECIONAR TODOS
    ' ==============================
    Private Sub CHKSelecionarTodos_CheckedChanged(
        sender As Object,
        e As EventArgs
    ) Handles CHKSelecionarTodos.CheckedChanged

        If ignorarEventos Then Exit Sub

        ignorarEventos = True
        ListView1.BeginUpdate()

        For Each item As ListViewItem In ListView1.Items
            item.Checked = CHKSelecionarTodos.Checked
        Next

        ListView1.EndUpdate()
        ignorarEventos = False

    End Sub

    Private Sub ListView1_ItemChecked(
        sender As Object,
        e As ItemCheckedEventArgs
    ) Handles ListView1.ItemChecked

        If ignorarEventos Then Exit Sub

        ignorarEventos = True

        If Not e.Item.Checked AndAlso CHKSelecionarTodos.Checked Then
            CHKSelecionarTodos.Checked = False
        End If

        If e.Item.Checked Then
            Dim todosMarcados As Boolean = True
            For Each item As ListViewItem In ListView1.Items
                If Not item.Checked Then
                    todosMarcados = False
                    Exit For
                End If
            Next
            If todosMarcados Then CHKSelecionarTodos.Checked = True
        End If

        ignorarEventos = False

    End Sub

End Class
