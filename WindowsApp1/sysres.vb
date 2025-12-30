Imports System.Diagnostics
Imports System.Management
Imports System.IO
Imports OpenHardwareMonitor.Hardware

Public Class SysRes
    ' ===== PerformanceCounters =====
    Private cpuCounters() As PerformanceCounter
    Private cpuUsages() As Single
    Private ramTotal As Long
    Private ramCounter As PerformanceCounter
    Private diskActivityCounter As PerformanceCounter
    Private diskActivity As Single = 0
    Private drives() As DriveInfo
    Private gpuUsage As Single = 0

    ' ===== Temperaturas =====
    Private cpuTemp As Single = 0
    Private gpuTemp As Single = 0

    ' ===== Timer =====
    Private t As Timer

    ' ===== OpenHardwareMonitor =====
    Private computer As Computer

    Private Sub SysRes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' ==== CPU ====
        Dim coreCount As Integer = Environment.ProcessorCount
        ReDim cpuCounters(coreCount - 1)
        ReDim cpuUsages(coreCount - 1)
        For i = 0 To coreCount - 1
            cpuCounters(i) = New PerformanceCounter("Processor", "% Processor Time", i.ToString())
            cpuCounters(i).NextValue()
        Next

        ' ==== Memória ====
        ramCounter = New PerformanceCounter("Memory", "Available Bytes")
        ramTotal = CLng(New ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem").[Get]().Cast(Of ManagementObject)().FirstOrDefault()("TotalPhysicalMemory"))

        ' ==== Disco total e atividade ====
        diskActivityCounter = New PerformanceCounter("LogicalDisk", "% Disk Time", "_Total")
        diskActivityCounter.NextValue() ' Primeiro valor inicial
        drives = DriveInfo.GetDrives().Where(Function(d) d.IsReady).ToArray()

        ' ==== OpenHardwareMonitor ====
        computer = New Computer With {.CPUEnabled = True, .GPUEnabled = True}
        computer.Open()

        ' ==== Timer ====
        t = New Timer()
        t.Interval = 1000
        AddHandler t.Tick, AddressOf TimerTick
        t.Start()

        Me.DoubleBuffered = True
    End Sub

    Private Sub TimerTick(sender As Object, e As EventArgs)
        ' ==== CPU uso ====
        For i = 0 To cpuCounters.Length - 1
            cpuUsages(i) = cpuCounters(i).NextValue()
        Next

        ' ==== Temperaturas via OpenHardwareMonitor ====
        cpuTemp = GetAverageCPUTemp()
        gpuTemp = GetGPUTemp()
        gpuUsage = GetGPUUsage()

        ' ==== Disco atividade ====
        diskActivity = diskActivityCounter.NextValue()

        Me.Invalidate()
    End Sub

    Private Sub SysRes_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        Dim g As Graphics = e.Graphics
        g.Clear(Me.BackColor)

        Dim padding As Integer = 20
        Dim x As Integer = padding
        Dim y As Integer = padding
        Dim barWidth As Integer = 40
        Dim barHeight As Integer = 150

        ' ==== CPU vertical ====
        For i = 0 To cpuUsages.Length - 1
            Dim height As Integer = CInt(cpuUsages(i) / 100 * barHeight)
            g.FillRectangle(Brushes.Green, x, y + (barHeight - height), barWidth, height)
            g.DrawRectangle(Pens.Black, x, y, barWidth, barHeight)
            g.DrawString($"{cpuUsages(i):0}%", Me.Font, Brushes.Black, x, y + barHeight + 2)
            g.DrawString($"CPU {i}", Me.Font, Brushes.Black, x, y + barHeight + 20)
            x += barWidth + padding
        Next

        y += barHeight + 60
        x = padding

        ' ==== Memória horizontal ====
        Dim ramUsed As Single = ramTotal - ramCounter.NextValue()
        DrawHorizontalBar(g, x, y, 400, 25, ramUsed / ramTotal * 100, $"Memória ({ramUsed / 1024 / 1024:0}MB/{ramTotal / 1024 / 1024:0}MB)")
        y += 60

        ' ==== Temperatura CPU ====
        DrawHorizontalBar(g, x, y, 400, 25, cpuTemp, $"Temp CPU: {cpuTemp:0}°C", True)
        y += 60

        ' ==== Temperatura GPU ====
        DrawHorizontalBar(g, x, y, 400, 25, gpuTemp, $"Temp GPU: {gpuTemp:0}°C", True)
        y += 60

        ' ==== Disco total horizontal (atividade) ====
        DrawHorizontalBar(g, x, y, 400, 25, diskActivity, "Atividade Disco (%)")
        y += 60

        ' ==== Cada unidade de disco (uso de espaço) ====
        For Each d As DriveInfo In drives
            Dim percentUsed As Single = 100 * (d.TotalSize - d.AvailableFreeSpace) / d.TotalSize
            DrawHorizontalBar(g, x + 20, y, 250, 20, percentUsed, $"Drive {d.Name}")
            y += 40
        Next

        ' ==== GPU Usage ====
        DrawHorizontalBar(g, x, y, 400, 25, gpuUsage, $"GPU Usage: {gpuUsage:0}%")

        ' ==== Ajusta tamanho do form automaticamente ====
        Me.Height = y + 100
        Me.Width = 500
    End Sub

    Private Sub DrawHorizontalBar(g As Graphics, x As Integer, y As Integer, width As Integer, height As Integer, value As Single, label As String, Optional isTemp As Boolean = False)
        Dim percent As Single = Math.Min(value, 100)
        Dim fillWidth As Integer = CInt(width * percent / 100)
        Dim brush As Brush

        If isTemp Then
            If percent < 60 Then
                brush = Brushes.Green
            ElseIf percent < 80 Then
                brush = Brushes.Yellow
            Else
                brush = Brushes.Red
            End If
        Else
            If percent < 70 Then
                brush = Brushes.Blue
            ElseIf percent < 90 Then
                brush = Brushes.Yellow
            Else
                brush = Brushes.Red
            End If
        End If

        g.FillRectangle(brush, x, y, fillWidth, height)
        g.DrawRectangle(Pens.Black, x, y, width, height)
        g.DrawString(label, Me.Font, Brushes.Black, x, y + height + 2)
    End Sub

    ' ==== Funções OpenHardwareMonitor ====
    Private Function GetAverageCPUTemp() As Single
        Dim total As Single = 0
        Dim count As Integer = 0

        For Each hardware In computer.Hardware
            If hardware.HardwareType = HardwareType.CPU Then
                hardware.Update()
                For Each sensor In hardware.Sensors
                    If sensor.SensorType = SensorType.Temperature AndAlso sensor.Value.HasValue Then
                        total += sensor.Value.Value
                        count += 1
                    End If
                Next
            End If
        Next

        Return If(count > 0, total / count, 0)
    End Function

    Private Function GetGPUTemp() As Single
        For Each hardware In computer.Hardware
            If hardware.HardwareType = HardwareType.GpuNvidia Or hardware.HardwareType = HardwareType.GpuAti Then
                hardware.Update()
                For Each sensor In hardware.Sensors
                    If sensor.SensorType = SensorType.Temperature AndAlso sensor.Value.HasValue Then
                        Return sensor.Value.Value
                    End If
                Next
            End If
        Next
        Return 0
    End Function

    Private Function GetGPUUsage() As Single
        For Each hardware In computer.Hardware
            If hardware.HardwareType = HardwareType.GpuNvidia Or hardware.HardwareType = HardwareType.GpuAti Then
                hardware.Update()
                For Each sensor In hardware.Sensors
                    If sensor.SensorType = SensorType.Load AndAlso sensor.Value.HasValue Then
                        Return sensor.Value.Value
                    End If
                Next
            End If
        Next
        Return 0
    End Function
End Class
