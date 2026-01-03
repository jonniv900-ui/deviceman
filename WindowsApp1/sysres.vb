Imports System.Diagnostics
Imports System.Management
Imports System.IO
Imports OpenHardwareMonitor.Hardware
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices
Imports System.Net.NetworkInformation
Imports System.Linq

Public Class SysRes
    ' --- APIs Windows para Movimentação e Resolução ---
    <DllImport("user32.dll")> Private Shared Function ReleaseCapture() As Boolean : End Function
    <DllImport("user32.dll")> Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As Integer, lParam As Integer) As Integer : End Function
    <DllImport("user32.dll")> Private Shared Function EnumDisplaySettings(deviceName As String, modeNum As Integer, ByRef devMode As DEVMODE) As Boolean : End Function

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
    Public Structure DEVMODE
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public dmDeviceName As String
        Public dmSpecVersion, dmDriverVersion, dmSize, dmDriverExtra, dmFields As Integer
        Public dmPositionX, dmPositionY, dmDisplayOrientation, dmDisplayFixedOutput, dmColor, dmDuplex, dmYResolution, dmTTOption, dmCollate As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public dmFormName As String
        Public dmLogPixels, dmBitsPerPel, dmPelsWidth, dmPelsHeight, dmDisplayFlags, dmDisplayFrequency As Integer
        Public dmICMMethod, dmICMIntent, dmMediaType, dmDitherType, dmReserved1, dmReserved2, dmPanningWidth, dmPanningHeight As Integer
    End Structure

    ' ===== Variáveis de Dados =====
    Private cpuCounters() As PerformanceCounter
    Private cpuUsages() As Single
    Private cpuHistory() As List(Of Single)
    Private coreColors() As Brush
    Private ramTotal As Long
    Private ramCounter, diskActivityCounter, netDownCounter, netUpCounter As PerformanceCounter

    Private cpuFullName As String = "CPU"
    Private gpuFullName As String = "GPU"
    Private screenRes As String = ""
    Private screenHz As String = ""
    Private diskModels As New Dictionary(Of String, String)
    Private drives() As DriveInfo
    Private computer As Computer
    Private cpuTemp, gpuTemp, gpuUsage, diskActivity, downloadSpeed, uploadSpeed As Single

    ' ===== UI =====
    Private WithEvents t As Timer
    Private WithEvents btnCpuToggle As Button
    Private WithEvents chkTopMost, chkCompact As CheckBox
    Private showThreads As Boolean = True
    Private isCompact As Boolean = False
    Private normalSize As Size = New Size(1100, 850)
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTCAPTION As Integer = 2

    Private Sub SysRes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.DoubleBuffered = True
        Me.Size = normalSize
        Me.BackColor = Color.FromArgb(240, 242, 245)
        Me.Text = "System Resources Monitor"

        ' Inicialização dos Controles
        chkTopMost = New CheckBox() With {.Text = "Topo", .Location = New Point(15, 12), .AutoSize = True, .Cursor = Cursors.Hand}
        chkCompact = New CheckBox() With {.Text = "Overlay", .Location = New Point(80, 12), .AutoSize = True, .Cursor = Cursors.Hand}
        btnCpuToggle = New Button() With {.Text = "Threads: ON", .Location = New Point(160, 10), .Size = New Size(100, 24), .BackColor = Color.White, .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}

        Me.Controls.AddRange({chkTopMost, chkCompact, btnCpuToggle})

        InitHardware()
        SetupCPUCounters()
        SetupNetwork()

        ramCounter = New PerformanceCounter("Memory", "Available Bytes")
        ramTotal = GetTotalPhysicalMemory()
        diskActivityCounter = New PerformanceCounter("PhysicalDisk", "% Disk Time", "_Total")

        ' Inicializa Sensores (Requer Admin)
        computer = New Computer With {.CPUEnabled = True, .GPUEnabled = True}
        Try : computer.Open() : Catch : End Try

        t = New Timer() With {.Interval = 1000, .Enabled = True}
    End Sub

    Private Sub InitHardware()
        Try
            Dim dm As New DEVMODE() : dm.dmSize = CShort(Marshal.SizeOf(dm))
            If EnumDisplaySettings(Nothing, -1, dm) Then
                screenRes = $"{dm.dmPelsWidth}x{dm.dmPelsHeight}"
                screenHz = $"{dm.dmDisplayFrequency}Hz"
            End If

            Using s As New ManagementObjectSearcher("SELECT Name FROM Win32_Processor")
                For Each o In s.Get() : cpuFullName = o("Name").ToString().Trim() : Next
            End Using
            Using s As New ManagementObjectSearcher("SELECT Name FROM Win32_VideoController")
                For Each o In s.Get() : gpuFullName = o("Name").ToString().Trim() : Next
            End Using

            diskModels.Clear()
            Using s As New ManagementObjectSearcher("SELECT DeviceID, Model FROM Win32_DiskDrive")
                For Each drive In s.Get()
                    Dim dID = drive("DeviceID").ToString()
                    Using s2 As New ManagementObjectSearcher($"ASSOCIATORS OF {{Win32_DiskDrive.DeviceID='{dID}'}} WHERE AssocClass = Win32_DiskDriveToDiskPartition")
                        For Each part In s2.Get()
                            Using s3 As New ManagementObjectSearcher($"ASSOCIATORS OF {{Win32_DiskPartition.DeviceID='{part("DeviceID")}'}} WHERE AssocClass = Win32_LogicalDiskToPartition")
                                For Each log In s3.Get() : diskModels(log("Name").ToString()) = drive("Model").ToString() : Next
                            End Using
                        Next
                    End Using
                Next
            End Using
            drives = DriveInfo.GetDrives().Where(Function(d) d.IsReady).ToArray()
        Catch : End Try
    End Sub

    Private Sub SetupCPUCounters()
        Dim count = If(showThreads, Environment.ProcessorCount, Math.Max(1, Environment.ProcessorCount \ 2))
        ReDim cpuCounters(count - 1), cpuUsages(count - 1), cpuHistory(count - 1), coreColors(count - 1)
        Dim rnd As New Random()
        For i = 0 To count - 1
            cpuCounters(i) = New PerformanceCounter("Processor", "% Processor Time", i.ToString())
            cpuHistory(i) = New List(Of Single)()
            coreColors(i) = New SolidBrush(Color.FromArgb(60, 120, rnd.Next(160, 255)))
        Next
    End Sub

    Private Sub SetupNetwork()
        Try
            Dim cat As New PerformanceCounterCategory("Network Interface")
            Dim instances = cat.GetInstanceNames()
            ' Busca interface ativa ignorando filtros virtuais comuns
            Dim inst = instances.FirstOrDefault(Function(n) (n.ToLower().Contains("ethernet") OrElse n.ToLower().Contains("wi-fi") OrElse n.ToLower().Contains("realtek") OrElse n.ToLower().Contains("intel")) AndAlso Not n.ToLower().Contains("pseudo") AndAlso Not n.ToLower().Contains("virtual"))

            If inst IsNot Nothing Then
                netDownCounter = New PerformanceCounter("Network Interface", "Bytes Received/sec", inst)
                netUpCounter = New PerformanceCounter("Network Interface", "Bytes Sent/sec", inst)
            End If
        Catch : End Try
    End Sub

    Private Sub t_Tick(sender As Object, e As EventArgs) Handles t.Tick
        Try
            ' Coleta CPU
            For i = 0 To cpuCounters.Length - 1
                cpuUsages(i) = cpuCounters(i).NextValue()
                cpuHistory(i).Add(cpuUsages(i))
                If cpuHistory(i).Count > 30 Then cpuHistory(i).RemoveAt(0)
            Next

            UpdateSensors()

            ' Coleta Rede (Conversão para Mbps)
            If netDownCounter IsNot Nothing Then
                downloadSpeed = (netDownCounter.NextValue() * 8) / 1000000
                uploadSpeed = (netUpCounter.NextValue() * 8) / 1000000
            End If

            diskActivity = diskActivityCounter.NextValue()
        Catch : End Try
        Me.Invalidate()
    End Sub

    Private Sub UpdateSensors()
        If computer Is Nothing Then Exit Sub
        For Each hw In computer.Hardware
            hw.Update()
            For Each s In hw.Sensors
                If s.Value.HasValue Then
                    If hw.HardwareType = HardwareType.CPU And s.SensorType = SensorType.Temperature Then cpuTemp = s.Value.Value
                    If (hw.HardwareType = HardwareType.GpuNvidia Or hw.HardwareType = HardwareType.GpuAti) Then
                        If s.SensorType = SensorType.Temperature Then gpuTemp = s.Value.Value
                        If s.SensorType = SensorType.Load Then gpuUsage = s.Value.Value
                    End If
                End If
            Next
        Next
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g = e.Graphics : g.SmoothingMode = SmoothingMode.AntiAlias
        If isCompact Then DrawOverlay(g) Else DrawFull(g)
    End Sub

    Private Sub DrawFull(g As Graphics)
        Dim colW = (Me.Width \ 2) - 40

        ' --- LADO ESQUERDO: CPU E GPU ---
        DrawCard(g, 20, 50, colW, 400, "PROCESSADOR")
        g.DrawString($"{cpuFullName} | {screenRes} @ {screenHz}", New Font("Segoe UI", 8, FontStyle.Bold), Brushes.Black, 35, 75)

        Dim cpuTBrush = If(cpuTemp > 80, Brushes.Red, Brushes.Orange)
        DrawBar(g, 35, 95, colW - 40, 15, cpuTemp, $"CPU Temp: {cpuTemp:0}°C", cpuTBrush)

        Dim curX = 35, curY = 145
        For i = 0 To cpuUsages.Length - 1
            If curX + 110 > colW Then : curX = 35 : curY += 80 : End If
            DrawCore(g, curX, curY, cpuUsages(i), cpuHistory(i), coreColors(i), i)
            curX += 115
        Next

        DrawCard(g, 20, 470, colW, 200, "PLACA DE VÍDEO")
        g.DrawString(gpuFullName, New Font("Segoe UI", 9, FontStyle.Bold), Brushes.Black, 35, 495)
        DrawBar(g, 35, 520, colW - 40, 20, gpuUsage, $"GPU Carga: {gpuUsage:0}%", Brushes.Crimson)

        Dim gpuTBrush = If(gpuTemp > 80, Brushes.Red, Brushes.DarkRed)
        DrawBar(g, 35, 580, colW - 40, 20, gpuTemp, $"GPU Temp: {gpuTemp:0}°C", gpuTBrush)

        ' --- LADO DIREITO: RAM, REDE E DISCOS ---
        Dim rx = colW + 40, py = 50
        DrawCard(g, rx, py, colW, 620, "SISTEMA E ARMAZENAMENTO")

        Dim pRam = (1 - ramCounter.NextValue() / ramTotal) * 100
        DrawBar(g, rx + 15, py + 40, colW - 30, 20, pRam, $"Memória RAM: {pRam:0}%", Brushes.RoyalBlue)

        ' Rede com Download (Barra) e Upload (Texto)
        Dim netTxt = $"Download: ↓ {downloadSpeed:F1} Mbps | Upload: ↑ {uploadSpeed:F1} Mbps"
        DrawBar(g, rx + 15, py + 100, colW - 30, 20, Math.Min(downloadSpeed, 100), netTxt, Brushes.Teal)

        py += 180
        DrawBar(g, rx + 15, py, colW - 30, 20, diskActivity, $"Atividade de Disco (IO): {diskActivity:0}%", Brushes.Purple)

        py += 70
        For Each d In drives
            Dim model = If(diskModels.ContainsKey(d.Name.Replace("\", "")), diskModels(d.Name.Replace("\", "")), "Disco")
            Dim p = CSng(((d.TotalSize - d.AvailableFreeSpace) / d.TotalSize) * 100)
            DrawBar(g, rx + 15, py, colW - 30, 15, p, $"{d.Name} [{model}] {p:0}%", Brushes.Gray)
            py += 55
        Next
    End Sub

    Private Sub DrawOverlay(g As Graphics)
        g.Clear(Color.FromArgb(35, 35, 35))
        Dim w = Me.Width - 20
        DrawBar(g, 10, 40, w, 12, cpuUsages.Average(), $"CPU: {cpuTemp:0}°C | IO: {diskActivity:0}%", Brushes.DodgerBlue)
        DrawBar(g, 10, 85, w, 12, gpuUsage, $"GPU: {gpuTemp:0}°C | Load: {gpuUsage:0}%", Brushes.Crimson)
        DrawBar(g, 10, 130, w, 12, Math.Min(downloadSpeed, 100), $"↓ {downloadSpeed:F1} | ↑ {uploadSpeed:F1} Mbps", Brushes.Teal)
    End Sub

    ' --- Auxiliares de Desenho ---
    Private Sub DrawCard(g As Graphics, x As Integer, y As Integer, w As Integer, h As Integer, title As String)
        GraphicsExtensions.FillRoundedRectangle(g, Brushes.White, x, y, w, h, 10)
        g.DrawString(title, New Font("Segoe UI", 7, FontStyle.Bold), Brushes.DarkGray, x + 10, y + 8)
    End Sub

    Private Sub DrawBar(g As Graphics, x As Integer, y As Integer, w As Integer, h As Integer, val As Single, txt As String, b As Brush)
        g.FillRectangle(New SolidBrush(Color.FromArgb(40, 150, 150, 150)), x, y, w, h)
        g.FillRectangle(b, x, y, CInt(Math.Min(val, 100) / 100 * w), h)
        g.DrawString(txt, New Font("Segoe UI", 8, FontStyle.Bold), If(isCompact, Brushes.White, Brushes.Black), x, y + h + 2)
    End Sub

    Private Sub DrawCore(g As Graphics, x As Integer, y As Integer, val As Single, hist As List(Of Single), b As Brush, i As Integer)
        g.FillRectangle(b, x, y + (40 - CInt(val / 100 * 40)), 20, CInt(val / 100 * 40))
        g.DrawRectangle(Pens.LightGray, x, y, 20, 40)
        If hist.Count > 1 Then
            Dim pts = hist.Select(Function(v, idx) New PointF(x + 25 + (idx * 2), y + 40 - (v / 100 * 40))).ToArray()
            g.DrawLines(New Pen(b, 1), pts)
        End If
        g.DrawString($"{i}: {val:0}%", New Font("Segoe UI", 7, FontStyle.Bold), Brushes.Gray, x, y + 42)
    End Sub

    ' --- Eventos de UI ---
    Private Sub chkCompact_CheckedChanged(sender As Object, e As EventArgs) Handles chkCompact.CheckedChanged
        isCompact = chkCompact.Checked
        If isCompact Then
            Me.FormBorderStyle = FormBorderStyle.None : Me.Size = New Size(240, 180) : Me.Opacity = 0.85
            btnCpuToggle.Visible = False
        Else
            Me.FormBorderStyle = FormBorderStyle.Sizable : Me.Size = normalSize : Me.Opacity = 1.0
            btnCpuToggle.Visible = True
        End If
    End Sub

    Private Sub chkTopMost_CheckedChanged(sender As Object, e As EventArgs) Handles chkTopMost.CheckedChanged
        Me.TopMost = chkTopMost.Checked : End Sub

    Private Sub btnCpuToggle_Click(sender As Object, e As EventArgs) Handles btnCpuToggle.Click
        showThreads = Not showThreads
        btnCpuToggle.Text = If(showThreads, "Threads: ON", "Cores: ON")
        SetupCPUCounters()
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        If isCompact Then
            ' Revela checkboxes no overlay ao passar o mouse no topo
            chkTopMost.Visible = (e.Y < 40) : chkCompact.Visible = (e.Y < 40)
        End If
    End Sub

    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            ReleaseCapture() : SendMessage(Me.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        End If
    End Sub

    Private Function GetTotalPhysicalMemory() As Long
        Using s As New ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")
            For Each o In s.Get() : Return CLng(o("TotalPhysicalMemory")) : Next
        End Using
        Return 0
    End Function
End Class

' --- Módulo de Extensão (Obrigatório fora da classe) ---
Public Module GraphicsExtensions
    <System.Runtime.CompilerServices.Extension()>
    Public Sub FillRoundedRectangle(g As Graphics, brush As Brush, x As Integer, y As Integer, width As Integer, height As Integer, radius As Integer)
        Using path As New GraphicsPath()
            Dim d = radius * 2
            If d > width Then d = width
            If d > height Then d = height
            path.AddArc(x, y, d, d, 180, 90)
            path.AddArc(x + width - d, y, d, d, 270, 90)
            path.AddArc(x + width - d, y + height - d, d, d, 0, 90)
            path.AddArc(x, y + height - d, d, d, 90, 90)
            path.CloseAllFigures()
            g.FillPath(brush, path)
        End Using
    End Sub
End Module