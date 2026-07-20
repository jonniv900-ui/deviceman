Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms

''' <summary>
''' Gerenciador de tarefas: lista de processos com CPU%, memória e I/O de disco calculados
''' por PID (sem ambiguidade de instância), atualização assíncrona sem travar a UI,
''' ordenação por coluna, filtro por nome, menu de contexto e barra de status.
''' Controles montados em código - não depende de um tskmgr.Designer.vb.
''' </summary>
Public Class tskmgr
    Inherits Form

    Private ReadOnly service As New ProcessSnapshotService()
    Private cts As CancellationTokenSource
    Private refreshIntervalMs As Integer = 1000
    Private lastSnapshot As New List(Of ProcessRow)
    Private sortColumn As Integer = -1
    Private sortAscending As Boolean = True

    ' --- Controles ---
    Private txtFilter As TextBox
    Private cmbInterval As ComboBox
    Private btnRefreshNow As Button
    Private btnKill As Button
    Private btnKillTree As Button
    Private btnNewProcess As Button
    Private lvProcesses As ListView
    Private statusStrip As StatusStrip
    Private lblStatusCount As ToolStripStatusLabel
    Private lblStatusCpu As ToolStripStatusLabel
    Private lblStatusMem As ToolStripStatusLabel

    Private cmProcess As ContextMenuStrip
    Private miKill As ToolStripMenuItem
    Private miKillTree As ToolStripMenuItem
    Private miOpenLocation As ToolStripMenuItem
    Private miProperties As ToolStripMenuItem
    Private miPriority As ToolStripMenuItem
    Private miCopyPid As ToolStripMenuItem
    Private miCopyName As ToolStripMenuItem

    Private processIcons As ImageList
    Private ReadOnly iconKeysAdded As New HashSet(Of String)

    Public Sub New()
        InitializeComponent()
    End Sub

    ' ============================================================
    ' MONTAGEM DA TELA
    ' ============================================================
    Private Sub InitializeComponent()
        Me.Text = "Gerenciador de Tarefas"
        Me.Width = 900
        Me.Height = 600
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimumSize = New Size(700, 400)

        BuildTopPanel()
        BuildStatusStrip()
        BuildListView()
        BuildContextMenu()

        AddHandler Me.Load, AddressOf tskmgr_Load
        AddHandler Me.FormClosing, AddressOf tskmgr_FormClosing
    End Sub

    Private Sub BuildTopPanel()
        Dim panel As New Panel With {.Dock = DockStyle.Top, .Height = 76}
        Me.Controls.Add(panel)

        ' Linha 1: filtro e intervalo de atualização
        Dim lblFiltro As New Label With {.Text = "Filtrar:", .Left = 8, .Top = 14, .Width = 45}
        txtFilter = New TextBox With {.Left = 55, .Top = 10, .Width = 200}
        AddHandler txtFilter.TextChanged, Sub() RefreshListView()

        Dim lblIntervalo As New Label With {.Text = "Atualizar a cada:", .Left = 270, .Top = 14, .Width = 100}
        cmbInterval = New ComboBox With {.Left = 375, .Top = 10, .Width = 120, .DropDownStyle = ComboBoxStyle.DropDownList}
        cmbInterval.Items.AddRange({"500 ms", "1 segundo", "2 segundos", "5 segundos"})
        cmbInterval.SelectedIndex = 1
        AddHandler cmbInterval.SelectedIndexChanged, AddressOf CmbInterval_SelectedIndexChanged

        ' Linha 2: ações
        btnNewProcess = New Button With {.Text = "Novo Processo...", .Left = 8, .Top = 42, .Width = 130}
        AddHandler btnNewProcess.Click, AddressOf BtnNewProcess_Click

        btnRefreshNow = New Button With {.Text = "Atualizar Agora", .Left = 145, .Top = 42, .Width = 120}
        AddHandler btnRefreshNow.Click, AddressOf BtnRefreshNow_Click

        btnKill = New Button With {.Text = "Finalizar Processo", .Left = 275, .Top = 42, .Width = 130}
        AddHandler btnKill.Click, AddressOf BtnKill_Click

        btnKillTree = New Button With {.Text = "Finalizar Árvore", .Left = 415, .Top = 42, .Width = 120}
        AddHandler btnKillTree.Click, AddressOf BtnKillTree_Click

        panel.Controls.AddRange({lblFiltro, txtFilter, lblIntervalo, cmbInterval,
                                  btnNewProcess, btnRefreshNow, btnKill, btnKillTree})
    End Sub

    Private Sub BuildListView()
        processIcons = New ImageList With {.ImageSize = New Size(16, 16), .ColorDepth = ColorDepth.Depth32Bit}
        Try
            processIcons.Images.Add("generic", SystemIcons.Application)
        Catch
        End Try
        iconKeysAdded.Add("generic")

        lvProcesses = New ListView With {
            .Dock = DockStyle.Fill,
            .View = View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .MultiSelect = False,
            .HideSelection = False,
            .SmallImageList = processIcons
        }
        lvProcesses.Columns.Add("Nome", 180)
        lvProcesses.Columns.Add("PID", 70, HorizontalAlignment.Right)
        lvProcesses.Columns.Add("CPU (%)", 70, HorizontalAlignment.Right)
        lvProcesses.Columns.Add("Memória (MB)", 100, HorizontalAlignment.Right)
        lvProcesses.Columns.Add("Disco (KB/s)", 100, HorizontalAlignment.Right)
        lvProcesses.Columns.Add("Prioridade", 120)
        lvProcesses.Columns.Add("Caminho", 260)

        AddHandler lvProcesses.ColumnClick, AddressOf LvProcesses_ColumnClick
        AddHandler lvProcesses.DoubleClick, Sub() ShowProperties()

        Me.Controls.Add(lvProcesses)
    End Sub

    Private Sub BuildContextMenu()
        cmProcess = New ContextMenuStrip()

        miKill = New ToolStripMenuItem("Finalizar Processo")
        AddHandler miKill.Click, Sub() KillSelected(False)

        miKillTree = New ToolStripMenuItem("Finalizar Árvore de Processos")
        AddHandler miKillTree.Click, Sub() KillSelected(True)

        miOpenLocation = New ToolStripMenuItem("Abrir Local do Arquivo")
        AddHandler miOpenLocation.Click, AddressOf MiOpenLocation_Click

        miProperties = New ToolStripMenuItem("Propriedades")
        AddHandler miProperties.Click, Sub() ShowProperties()

        miPriority = New ToolStripMenuItem("Definir Prioridade")
        For Each pc In {ProcessPriorityClass.Idle, ProcessPriorityClass.BelowNormal, ProcessPriorityClass.Normal,
                        ProcessPriorityClass.AboveNormal, ProcessPriorityClass.High, ProcessPriorityClass.RealTime}
            Dim mi As New ToolStripMenuItem(PriorityLabel(pc)) With {.Tag = pc}
            AddHandler mi.Click, AddressOf MiPriority_Click
            miPriority.DropDownItems.Add(mi)
        Next

        miCopyPid = New ToolStripMenuItem("Copiar PID")
        AddHandler miCopyPid.Click, AddressOf MiCopyPid_Click

        miCopyName = New ToolStripMenuItem("Copiar Nome")
        AddHandler miCopyName.Click, AddressOf MiCopyName_Click

        cmProcess.Items.AddRange({miKill, miKillTree, New ToolStripSeparator(),
                                   miOpenLocation, miProperties, miPriority, New ToolStripSeparator(),
                                   miCopyPid, miCopyName})

        AddHandler cmProcess.Opening, AddressOf CmProcess_Opening
        lvProcesses.ContextMenuStrip = cmProcess
    End Sub

    Private Sub BuildStatusStrip()
        statusStrip = New StatusStrip()
        lblStatusCount = New ToolStripStatusLabel With {.Text = "Processos: -", .Spring = False, .BorderSides = ToolStripStatusLabelBorderSides.Right}
        lblStatusCpu = New ToolStripStatusLabel With {.Text = "CPU: -", .BorderSides = ToolStripStatusLabelBorderSides.Right}
        lblStatusMem = New ToolStripStatusLabel With {.Text = "Memória: -", .Spring = True, .TextAlign = ContentAlignment.MiddleLeft}
        statusStrip.Items.AddRange({lblStatusCount, lblStatusCpu, lblStatusMem})
        Me.Controls.Add(statusStrip)
    End Sub

    ' ============================================================
    ' CICLO DE VIDA / LOOP DE ATUALIZAÇÃO
    ' ============================================================
    Private Sub tskmgr_Load(sender As Object, e As EventArgs)
        cts = New CancellationTokenSource()
        StartRefreshLoop()
    End Sub

    Private Sub tskmgr_FormClosing(sender As Object, e As FormClosingEventArgs)
        cts?.Cancel()
    End Sub

    ''' <summary>Loop assíncrono: a coleta roda em background (Task.Run); a atualização da UI
    ''' acontece automaticamente na thread da UI ao continuar depois do Await, sem Invoke manual.</summary>
    Private Async Sub StartRefreshLoop()
        Try
            While Not cts.Token.IsCancellationRequested
                Dim rows = Await Task.Run(Function() service.GetSnapshot())
                If cts.Token.IsCancellationRequested Then Exit While

                Dim newIcons = Await Task.Run(Function() ExtractIconsForPaths(rows))
                ApplyNewIcons(newIcons)

                lastSnapshot = rows
                RefreshListView()
                UpdateStatusBar(rows)

                Await Task.Delay(refreshIntervalMs, cts.Token)
            End While
        Catch ex As OperationCanceledException
            ' Encerramento normal (form fechado).
        Catch ex As Exception
            ' Loop não deve derrubar o app; próxima abertura tenta de novo.
        End Try
    End Sub

    Private Sub CmbInterval_SelectedIndexChanged(sender As Object, e As EventArgs)
        Select Case cmbInterval.SelectedIndex
            Case 0 : refreshIntervalMs = 500
            Case 1 : refreshIntervalMs = 1000
            Case 2 : refreshIntervalMs = 2000
            Case 3 : refreshIntervalMs = 5000
        End Select
    End Sub

    Private Async Sub BtnRefreshNow_Click(sender As Object, e As EventArgs)
        Dim rows = Await Task.Run(Function() service.GetSnapshot())
        Dim newIcons = Await Task.Run(Function() ExtractIconsForPaths(rows))
        ApplyNewIcons(newIcons)
        lastSnapshot = rows
        RefreshListView()
        UpdateStatusBar(rows)
    End Sub

    ''' <summary>Roda em background: extrai o ícone só de caminhos ainda não vistos (cache por caminho).</summary>
    Private Function ExtractIconsForPaths(rows As List(Of ProcessRow)) As Dictionary(Of String, Icon)
        Dim newIcons As New Dictionary(Of String, Icon)
        Dim distinctPaths = rows.Select(Function(r) r.FilePath).Where(Function(p) Not String.IsNullOrEmpty(p)).Distinct()
        For Each path In distinctPaths
            If Not iconKeysAdded.Contains(path) Then
                Try
                    Dim ico = Icon.ExtractAssociatedIcon(path)
                    If ico IsNot Nothing Then newIcons(path) = ico
                Catch
                    ' Sem acesso ao arquivo ou sem ícone associado - fica com o genérico.
                End Try
            End If
        Next
        Return newIcons
    End Function

    ''' <summary>Roda na UI thread: adiciona os ícones novos ao ImageList e libera o handle original.</summary>
    Private Sub ApplyNewIcons(newIcons As Dictionary(Of String, Icon))
        For Each kvp In newIcons
            Try
                processIcons.Images.Add(kvp.Key, kvp.Value)
                iconKeysAdded.Add(kvp.Key)
            Catch
            Finally
                kvp.Value.Dispose() ' o ImageList já copiou a imagem internamente
            End Try
        Next
    End Sub

    ' ============================================================
    ' ATUALIZAÇÃO DA LISTVIEW (filtro + diff, sem recriar tudo)
    ' ============================================================
    Private Sub RefreshListView()
        Dim filterText = txtFilter.Text.Trim().ToLowerInvariant()
        Dim filteredRows = If(String.IsNullOrEmpty(filterText),
                               lastSnapshot,
                               lastSnapshot.Where(Function(r) r.Name.ToLowerInvariant().Contains(filterText)).ToList())

        Dim existingByPid = lvProcesses.Items.Cast(Of ListViewItem)().
                             Where(Function(it) it.Tag IsNot Nothing).
                             ToDictionary(Function(it) CInt(it.Tag))

        Dim selectedPid As Integer? = Nothing
        If lvProcesses.SelectedItems.Count > 0 Then selectedPid = CInt(lvProcesses.SelectedItems(0).Tag)

        lvProcesses.BeginUpdate()
        Try
            Dim seenPids As New HashSet(Of Integer)
            For Each row In filteredRows
                seenPids.Add(row.PID)
                If existingByPid.ContainsKey(row.PID) Then
                    UpdateItemText(existingByPid(row.PID), row)
                Else
                    Dim item As New ListViewItem(row.Name)
                    item.Tag = row.PID
                    For i = 1 To 6
                        item.SubItems.Add("")
                    Next
                    UpdateItemText(item, row)
                    lvProcesses.Items.Add(item)
                End If
            Next

            For Each kvp In existingByPid
                If Not seenPids.Contains(kvp.Key) Then
                    lvProcesses.Items.Remove(kvp.Value)
                End If
            Next
        Finally
            lvProcesses.EndUpdate()
        End Try

        If sortColumn >= 0 Then lvProcesses.Sort()

        If selectedPid.HasValue Then
            For Each it As ListViewItem In lvProcesses.Items
                If CInt(it.Tag) = selectedPid.Value Then
                    it.Selected = True
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub UpdateItemText(item As ListViewItem, row As ProcessRow)
        item.Text = row.Name
        item.Tag = row.PID
        item.ImageKey = If(Not String.IsNullOrEmpty(row.FilePath) AndAlso iconKeysAdded.Contains(row.FilePath), row.FilePath, "generic")
        item.SubItems(1).Text = row.PID.ToString()
        item.SubItems(2).Text = row.CpuPercent.ToString("0.0", CultureInfo.InvariantCulture)
        item.SubItems(3).Text = row.MemoryMB.ToString("0", CultureInfo.InvariantCulture)
        item.SubItems(4).Text = row.DiskKBs.ToString("0", CultureInfo.InvariantCulture)
        item.SubItems(5).Text = row.Priority
        item.SubItems(6).Text = row.FilePath
    End Sub

    Private Sub UpdateStatusBar(rows As List(Of ProcessRow))
        lblStatusCount.Text = $"Processos: {rows.Count}"

        Dim totalCpu = Math.Min(100.0, rows.Sum(Function(r) r.CpuPercent))
        lblStatusCpu.Text = $"CPU: {totalCpu:0.0}%"

        Dim mem = service.GetSystemMemoryInfo()
        lblStatusMem.Text = $"Memória: {mem.UsedGB:0.0} GB / {mem.TotalGB:0.0} GB ({mem.Percent}%)"
    End Sub

    ' ============================================================
    ' ORDENAÇÃO POR COLUNA
    ' ============================================================
    Private Sub LvProcesses_ColumnClick(sender As Object, e As ColumnClickEventArgs)
        If e.Column = sortColumn Then
            sortAscending = Not sortAscending
        Else
            sortColumn = e.Column
            sortAscending = True
        End If
        lvProcesses.ListViewItemSorter = New ProcessListViewComparer(sortColumn, sortAscending)
        lvProcesses.Sort()
    End Sub

    Private Class ProcessListViewComparer
        Implements IComparer

        Private ReadOnly col As Integer
        Private ReadOnly asc As Boolean

        Public Sub New(col As Integer, asc As Boolean)
            Me.col = col
            Me.asc = asc
        End Sub

        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim itemX = CType(x, ListViewItem)
            Dim itemY = CType(y, ListViewItem)
            Dim textX = itemX.SubItems(col).Text
            Dim textY = itemY.SubItems(col).Text

            Dim numX As Double, numY As Double
            Dim result As Integer
            If Double.TryParse(textX, NumberStyles.Any, CultureInfo.InvariantCulture, numX) AndAlso
               Double.TryParse(textY, NumberStyles.Any, CultureInfo.InvariantCulture, numY) Then
                result = numX.CompareTo(numY)
            Else
                result = String.Compare(textX, textY, StringComparison.OrdinalIgnoreCase)
            End If

            Return If(asc, result, -result)
        End Function
    End Class

    ' ============================================================
    ' AÇÕES: FINALIZAR / NOVO PROCESSO
    ' ============================================================
    Private Function GetSelectedPid() As Integer?
        If lvProcesses.SelectedItems.Count = 0 Then Return Nothing
        Return CInt(lvProcesses.SelectedItems(0).Tag)
    End Function

    Private Function GetSelectedRow() As ProcessRow?
        Dim pid = GetSelectedPid()
        If Not pid.HasValue Then Return Nothing
        For Each r In lastSnapshot
            If r.PID = pid.Value Then Return r
        Next
        Return Nothing
    End Function

    Private Sub BtnKill_Click(sender As Object, e As EventArgs)
        KillSelected(False)
    End Sub

    Private Sub BtnKillTree_Click(sender As Object, e As EventArgs)
        KillSelected(True)
    End Sub

    Private Sub KillSelected(asTree As Boolean)
        Dim pid = GetSelectedPid()
        If Not pid.HasValue Then
            MessageBox.Show("Selecione um processo primeiro.", "Finalizar Processo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim row = GetSelectedRow()
        Dim nomeExibicao = If(row.HasValue, row.Value.Name, pid.Value.ToString())
        Dim pergunta = If(asTree,
            $"Finalizar '{nomeExibicao}' (PID {pid.Value}) e todos os processos filhos dele?",
            $"Finalizar o processo '{nomeExibicao}' (PID {pid.Value})?")

        If MessageBox.Show(pergunta, "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
            Return
        End If

        Try
            If asTree Then
                Dim psi As New ProcessStartInfo("taskkill.exe", $"/PID {pid.Value} /T /F") With {
                    .UseShellExecute = False, .CreateNoWindow = True,
                    .RedirectStandardOutput = True, .RedirectStandardError = True
                }
                Using proc = Process.Start(psi)
                    proc.WaitForExit(3000)
                End Using
            Else
                Process.GetProcessById(pid.Value).Kill()
            End If
        Catch ex As Exception
            MessageBox.Show("Não foi possível finalizar o processo: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnNewProcess_Click(sender As Object, e As EventArgs)
        Using dlg As New OpenFileDialog With {.Title = "Escolha um programa para executar", .Filter = "Executáveis|*.exe"}
            If dlg.ShowDialog() = DialogResult.OK Then
                Try
                    Process.Start(dlg.FileName)
                Catch ex As Exception
                    MessageBox.Show("Não foi possível iniciar o processo: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    ' ============================================================
    ' MENU DE CONTEXTO
    ' ============================================================
    Private Sub CmProcess_Opening(sender As Object, e As CancelEventArgs)
        Dim row = GetSelectedRow()
        Dim hasSelection = row.HasValue
        Dim hasPath = hasSelection AndAlso Not String.IsNullOrEmpty(row.Value.FilePath)

        miKill.Enabled = hasSelection
        miKillTree.Enabled = hasSelection
        miPriority.Enabled = hasSelection
        miCopyPid.Enabled = hasSelection
        miCopyName.Enabled = hasSelection
        miOpenLocation.Enabled = hasPath
        miProperties.Enabled = hasPath
    End Sub

    Private Sub MiOpenLocation_Click(sender As Object, e As EventArgs)
        Dim row = GetSelectedRow()
        If Not row.HasValue OrElse String.IsNullOrEmpty(row.Value.FilePath) Then
            MessageBox.Show("Caminho do executável não disponível para este processo.", "Abrir Local", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Try
            Process.Start("explorer.exe", $"/select,""{row.Value.FilePath}""")
        Catch ex As Exception
            MessageBox.Show("Não foi possível abrir o local do arquivo: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ShowProperties()
        Dim row = GetSelectedRow()
        If Not row.HasValue OrElse String.IsNullOrEmpty(row.Value.FilePath) Then Return
        Try
            Dim psi As New ProcessStartInfo(row.Value.FilePath) With {.UseShellExecute = True, .Verb = "properties"}
            Process.Start(psi)
        Catch ex As Exception
            MessageBox.Show("Não foi possível abrir as propriedades: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MiPriority_Click(sender As Object, e As EventArgs)
        Dim pid = GetSelectedPid()
        If Not pid.HasValue Then Return
        Dim novaPrioridade = CType(CType(sender, ToolStripMenuItem).Tag, ProcessPriorityClass)
        Try
            Dim p = Process.GetProcessById(pid.Value)
            p.PriorityClass = novaPrioridade
        Catch ex As Exception
            MessageBox.Show("Não foi possível alterar a prioridade (processo protegido ou sem permissão): " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MiCopyPid_Click(sender As Object, e As EventArgs)
        Dim pid = GetSelectedPid()
        If pid.HasValue Then Clipboard.SetText(pid.Value.ToString())
    End Sub

    Private Sub MiCopyName_Click(sender As Object, e As EventArgs)
        Dim row = GetSelectedRow()
        If row.HasValue Then Clipboard.SetText(row.Value.Name)
    End Sub

    Private Function PriorityLabel(pc As ProcessPriorityClass) As String
        Select Case pc
            Case ProcessPriorityClass.Idle : Return "Baixa"
            Case ProcessPriorityClass.BelowNormal : Return "Abaixo do Normal"
            Case ProcessPriorityClass.Normal : Return "Normal"
            Case ProcessPriorityClass.AboveNormal : Return "Acima do Normal"
            Case ProcessPriorityClass.High : Return "Alta"
            Case ProcessPriorityClass.RealTime : Return "Tempo Real"
            Case Else : Return "N/D"
        End Select
    End Function

End Class