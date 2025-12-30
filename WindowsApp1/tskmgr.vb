Imports System.Diagnostics
Imports System.Threading
Imports System.Windows.Forms

Public Class tskmgr
    ' ===== PerformanceCounter por PID =====
    Private cpuCounters As New Dictionary(Of Integer, PerformanceCounter)
    Private t As Thread

    Private Sub tskmgr_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Configura ListView
        ListView1.View = View.Details
        ListView1.Columns.Clear()
        ListView1.Columns.Add("Processo", 200)
        ListView1.Columns.Add("CPU (%)", 70)
        ListView1.Columns.Add("Memória (MB)", 90)
        ListView1.Columns.Add("Disco (MB/s)", 90)

        ' Inicia thread de atualização
        t = New Thread(AddressOf UpdateProcesses)
        t.IsBackground = True
        t.Start()
    End Sub

    Private Sub UpdateProcesses()
        While True
            Try
                Dim processes = Process.GetProcesses()
                ' Cria dicionário de itens existentes pelo PID
                Dim existingItems = ListView1.Items.Cast(Of ListViewItem)().
                                    Where(Function(it) it.Tag IsNot Nothing).
                                    ToDictionary(Function(it) CInt(it.Tag), Function(it) it)

                For Each p As Process In processes
                    Try
                        Dim name = p.ProcessName
                        Dim pid = p.Id
                        Dim memMB As Single = p.WorkingSet64 / 1024 / 1024

                        ' CPU
                        Dim cpuPerc As Single = 0
                        If Not cpuCounters.ContainsKey(pid) Then
                            Dim pc As New PerformanceCounter("Process", "% Processor Time", name, True)
                            pc.NextValue()
                            cpuCounters(pid) = pc
                            cpuPerc = 0
                        Else
                            cpuPerc = cpuCounters(pid).NextValue() / Environment.ProcessorCount
                        End If

                        ' Disco (simplificado como 0 por hora, pode implementar disco por PID)
                        Dim diskPerc As Single = 0

                        ' Atualiza UI de forma segura
                        If ListView1.InvokeRequired Then
                            ListView1.Invoke(Sub()
                                                 If existingItems.ContainsKey(pid) Then
                                                     Dim item As ListViewItem = existingItems(pid)
                                                     item.SubItems(1).Text = $"{cpuPerc:0}"
                                                     item.SubItems(2).Text = $"{memMB:0}"
                                                     item.SubItems(3).Text = $"{diskPerc:0}"
                                                     existingItems.Remove(pid)
                                                 Else
                                                     Dim lv As New ListViewItem(name)
                                                     lv.Tag = pid
                                                     lv.SubItems.Add($"{cpuPerc:0}")
                                                     lv.SubItems.Add($"{memMB:0}")
                                                     lv.SubItems.Add($"{diskPerc:0}")
                                                     ListView1.Items.Add(lv)
                                                 End If
                                             End Sub)
                        Else
                            If existingItems.ContainsKey(pid) Then
                                Dim item As ListViewItem = existingItems(pid)
                                item.SubItems(1).Text = $"{cpuPerc:0}"
                                item.SubItems(2).Text = $"{memMB:0}"
                                item.SubItems(3).Text = $"{diskPerc:0}"
                                existingItems.Remove(pid)
                            Else
                                Dim lv As New ListViewItem(name)
                                lv.Tag = pid
                                lv.SubItems.Add($"{cpuPerc:0}")
                                lv.SubItems.Add($"{memMB:0}")
                                lv.SubItems.Add($"{diskPerc:0}")
                                ListView1.Items.Add(lv)
                            End If
                        End If

                    Catch
                        ' Ignora processos inacessíveis
                    End Try
                Next

                ' Remove processos que encerraram
                If ListView1.InvokeRequired Then
                    ListView1.Invoke(Sub()
                                         For Each key In existingItems.Keys
                                             ListView1.Items.Remove(existingItems(key))
                                         Next
                                     End Sub)
                Else
                    For Each key In existingItems.Keys
                        ListView1.Items.Remove(existingItems(key))
                    Next
                End If

            Catch
                ' ignora erros
            End Try

            Thread.Sleep(1000)
        End While
    End Sub

    ' ===== Botão Encerrar =====
    Private Sub button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ListView1.SelectedItems.Count > 0 Then
            Dim pid As Integer = CInt(ListView1.SelectedItems(0).Tag)
            Try
                Dim p As Process = Process.GetProcessById(pid)
                p.Kill()
            Catch ex As Exception
                MessageBox.Show("Não foi possível encerrar o processo: " & ex.Message)
            End Try
        End If
    End Sub

    ' ===== Botão Novo Processo =====
    Private Sub button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim dlg As New OpenFileDialog()
        dlg.Title = "Escolha um programa para executar"
        dlg.Filter = "Executáveis|*.exe"
        If dlg.ShowDialog() = DialogResult.OK Then
            Try
                Process.Start(dlg.FileName)
            Catch ex As Exception
                MessageBox.Show("Não foi possível iniciar o processo: " & ex.Message)
            End Try
        End If
    End Sub
End Class
