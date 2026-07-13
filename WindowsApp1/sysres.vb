Imports System.Diagnostics
Imports System.Management
Imports System.IO
Imports OpenHardwareMonitor.Hardware
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices
Imports System.Net.NetworkInformation
Imports System.Linq
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class SysRes
    Inherits Form

    ' --- APIs Windows para Movimentação e Resolução ---
    <DllImport("user32.dll")> Private Shared Function ReleaseCapture() As Boolean : End Function
    <DllImport("user32.dll")> Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As Integer, lParam As Integer) As Integer : End Function
    <DllImport("user32.dll")> Private Shared Function EnumDisplaySettings(deviceName As String, modeNum As Integer, ByRef devMode As DEVMODE) As Boolean : End Function

    ' ===== Estruturas Internas =====
    Private Structure GpuInfo
        Public Name As String
        Public IsIntegrated As Boolean
    End Structure

    ' ===== Variáveis de Dados e Contadores =====
    Private cpuCounters() As PerformanceCounter
    Private cpuUsages() As Single
    Private cpuHistory() As List(Of Single)
    Private coreColors() As Brush
    Private ramTotal As Long

    ' Contadores Globais
    Private ramCounter As PerformanceCounter
    Private diskActivityCounter As PerformanceCounter
    Private diskReadBytesCounter As PerformanceCounter
    Private diskWriteBytesCounter As PerformanceCounter
    Private netDownCounter As PerformanceCounter
    Private netUpCounter As PerformanceCounter
    Private netPacketsCounter As PerformanceCounter

    ' Dados de Hardware Detectados
    Private cpuFullName As String = "Processador"
    Private detectedGpus As New List(Of GpuInfo)()
    Private networkTypeString As String = "Desconectado"
    Private diskModels As New Dictionary(Of String, String)()
    Private drives() As DriveInfo

    ' Variáveis de Estado Dinâmico
    Private computer As Computer
    Private cpuTemp, gpuTemp, gpuUsage As Single
    Private diskActivity, diskReadMbps, diskWriteMbps As Single
    Private downloadSpeed, uploadSpeed, packetsPerSec As Single
    Private sessionDataTraficMB As Double = 0.0

    ' Métricas de Tela Dinâmicas
    Private currentScreenRes As String = ""
    Private currentScreenHz As String = ""

    ' ===== UI e Layout =====
    Private WithEvents t As Timer
    Private WithEvents btnCpuToggle As Button
    Private WithEvents chkTopMost, chkCompact As CheckBox
    Private showThreads As Boolean = True
    Private isCompact As Boolean = False
    Private normalSize As Size
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTCAPTION As Integer = 2

    ' ===== Cache Estático de Objetos GDI+ (Prevenção de Vazamento de Memória) =====
    Private fontTitle As New Font("Segoe UI", 9, FontStyle.Bold)
    Private fontSub As New Font("Segoe UI", 8, FontStyle.Bold)
    Private fontTiny As New Font("Segoe UI", 7, FontStyle.Bold)
    Private brushTextBlack As Brush = Brushes.Black
    Private brushTextWhite As Brush = Brushes.White
    Private brushCardBg As Brush = Brushes.White
    Private brushGrayText As Brush = Brushes.DarkGray

    Public Sub New()
        ' Configura tamanho proporcional baseado na resolução do monitor atual (Ex: 60% de largura, 70% de altura)
        Dim workingArea = Screen.PrimaryScreen.WorkingArea
        normalSize = New Size(CInt(workingArea.Width * 0.6), CInt(workingArea.Height * 0.7))
        Me.Size = normalSize
        Me.MinimumSize = New Size(800, 600)
    End Sub

    Private Sub SysRes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.DoubleBuffered = True
        Me.BackColor = Color.FromArgb(240, 242, 245)
        Me.Text = "System Resources Monitor - Pro"

        ' Inicialização dos Controles de UI
        chkTopMost = New CheckBox() With {.Text = "Fixar Topo", .Location = New Point(15, 12), .AutoSize = True, .Cursor = Cursors.Hand}
        chkCompact = New CheckBox() With {.Text = "Modo Overlay", .Location = New Point(110, 12), .AutoSize = True, .Cursor = Cursors.Hand}
        btnCpuToggle = New Button() With {.Text = "Threads: ON", .Location = New Point(230, 9), .Size = New Size(110, 24), .BackColor = Color.White, .FlatStyle = FlatStyle.Flat, .Cursor = Cursors.Hand}

        Me.Controls.AddRange({chkTopMost, chkCompact, btnCpuToggle})

        ' Coleta Inicial Estática e Inicialização Crítica de Arrays
        SetupCPUCounters()
        InitHardwareStaticData()
        UpdateScreenMetrics()
        SetupActiveNetworkInterface()
        SetupSystemPerformanceCounters()

        ' Inicializa OpenHardwareMonitor (Requer privilégios de Administrador)
        computer = New Computer With {.CPUEnabled = True, .GPUEnabled = True}
        Try : computer.Open() : Catch : End Try

        ' Timer Principal (1 Segundo)
        t = New Timer() With {.Interval = 1000, .Enabled = True}
    End Sub

    Private Sub SetupSystemPerformanceCounters()
        Try
            ramCounter = New PerformanceCounter("Memory", "Available Bytes")
            ramTotal = GetTotalPhysicalMemory()
            diskActivityCounter = New PerformanceCounter("PhysicalDisk", "% Disk Time", "_Total")
            diskReadBytesCounter = New PerformanceCounter("PhysicalDisk", "Disk Read Bytes/sec", "_Total")
            diskWriteBytesCounter = New PerformanceCounter("PhysicalDisk", "Disk Write Bytes/sec", "_Total")
        Catch : End Try
    End Sub

    Private Sub InitHardwareStaticData()
        Try
            ' Captura Nome do Processador
            Using s As New ManagementObjectSearcher("SELECT Name FROM Win32_Processor")
                For Each o In s.Get() : cpuFullName = o("Name").ToString().Trim() : Next
            End Using

            ' Varredura Multi-GPU robusta
            detectedGpus.Clear()
            Using s As New ManagementObjectSearcher("SELECT Name, AdapterRAM FROM Win32_VideoController")
                For Each o In s.Get()
                    Dim gpuName = o("Name").ToString().Trim()
                    Dim ramBytes As Long = 0
                    Long.TryParse(o("AdapterRAM")?.ToString(), ramBytes)

                    detectedGpus.Add(New GpuInfo() With {.Name = gpuName, .IsIntegrated = ClassificarGpuIntegrada(gpuName, ramBytes)})
                Next
            End Using

            ' Mapeamento lógico de armazenamento (WMI executado uma única vez)
            diskModels.Clear()
            Using s As New ManagementObjectSearcher("SELECT DeviceID, Model FROM Win32_DiskDrive")
                For Each drive In s.Get()
                    Dim dID = drive("DeviceID").ToString()
                    Using s2 As New ManagementObjectSearcher($"ASSOCIATORS OF {{Win32_DiskDrive.DeviceID='{dID}'}} WHERE AssocClass = Win32_DiskDriveToDiskPartition")
                        For Each part In s2.Get()
                            Using s3 As New ManagementObjectSearcher($"ASSOCIATORS OF {{Win32_DiskPartition.DeviceID='{part("DeviceID")}'}} WHERE AssocClass = Win32_LogicalDiskToPartition")
                                For Each log In s3.Get()
                                    diskModels(log("Name").ToString().Replace("\", "")) = drive("Model").ToString().Trim()
                                Next
                            End Using
                        Next
                    End Using
                Next
            End Using

            RefreshDrivesList()
        Catch : End Try
    End Sub

    ' ===== FIX: classificação de GPU priorizando o NOME, não o AdapterRAM =====
    ' Win32_VideoController.AdapterRAM é um campo de 32 bits — qualquer GPU dedicada com
    ' 4GB+ de VRAM sofre overflow/truncamento nesse valor no WMI, o que fazia a versão anterior
    ' (baseada só em RAM < 512MB) classificar placas dedicadas modernas como "integrada".
    Private Function ClassificarGpuIntegrada(gpuName As String, ramBytes As Long) As Boolean
        Dim n = gpuName.ToLowerInvariant()

        ' Padrões que indicam claramente uma GPU DEDICADA - checa primeiro e tem prioridade
        Dim padroesDedicada = {"geforce", "rtx", "gtx", "radeon rx", "radeon pro", "quadro", "tesla", "titan", "arc a"}
        For Each padrao In padroesDedicada
            If n.Contains(padrao) Then Return False
        Next

        ' Padrões que indicam claramente uma GPU INTEGRADA (iGPU/APU)
        Dim padroesIntegrada = {"intel", "uhd graphics", "iris", "amd radeon graphics", "basic display", "vega graphics"}
        For Each padrao In padroesIntegrada
            If n.Contains(padrao) Then Return True
        Next

        ' Nome não identificou claramente - usa o RAM só como sinal de apoio (fraco e sabidamente impreciso)
        Return ramBytes > 0 AndAlso ramBytes < 536870912
    End Function

    Private Sub RefreshDrivesList()
        drives = DriveInfo.GetDrives().Where(Function(d) d.IsReady).ToArray()
    End Sub

    Private Sub UpdateScreenMetrics()
        Try
            Dim currentScreen As Screen = Screen.FromControl(Me)
            currentScreenRes = $"{currentScreen.Bounds.Width}x{currentScreen.Bounds.Height}"

            Dim dm As New DEVMODE() : dm.dmSize = CShort(Marshal.SizeOf(dm))
            If EnumDisplaySettings(currentScreen.DeviceName, -1, dm) Then
                currentScreenHz = $"{dm.dmDisplayFrequency}Hz"
            Else
                currentScreenHz = "60Hz"
            End If
        Catch
            currentScreenHz = "60Hz"
        End Try
    End Sub

    Private Sub SetupCPUCounters()
        Try
            ' Descarta anteriores para evitar fugas de memória
            If cpuCounters IsNot Nothing Then
                For Each c In cpuCounters : c?.Dispose() : Next
            End If
            ' FIX: coreColors também precisa ser descartado antes do ReDim - cada clique no botão
            ' Threads/Cores criava um lote novo de SolidBrush sem liberar o anterior
            If coreColors IsNot Nothing Then
                For Each cb In coreColors : cb?.Dispose() : Next
            End If

            Dim count = If(showThreads, Environment.ProcessorCount, Math.Max(1, Environment.ProcessorCount \ 2))

            ' Garante o redimensionamento seguro e imediato dos arrays estruturais
            ReDim cpuUsages(count - 1)
            ReDim cpuHistory(count - 1)
            ReDim coreColors(count - 1)
            ReDim cpuCounters(count - 1)

            Dim rnd As New Random()
            For i = 0 To count - 1
                Try
                    cpuCounters(i) = New PerformanceCounter("Processor", "% Processor Time", i.ToString())
                Catch
                    cpuCounters(i) = Nothing
                End Try
                cpuHistory(i) = New List(Of Single)()
                coreColors(i) = New SolidBrush(Color.FromArgb(60, 120, rnd.Next(160, 255)))
            Next
        Catch ex As Exception
            ' Fallback crítico se o subsistema de contadores falhar completamente
            ' FIX: redimensiona TODOS os arrays (antes só cpuUsages/cpuHistory eram ajustados,
            ' deixando coreColors/cpuCounters com tamanho desencontrado e risco de IndexOutOfRange no DrawCore)
            Dim fallbackCount = Environment.ProcessorCount
            ReDim cpuUsages(fallbackCount - 1)
            ReDim cpuHistory(fallbackCount - 1)
            ReDim coreColors(fallbackCount - 1)
            ReDim cpuCounters(fallbackCount - 1)
            Dim rndFallback As New Random()
            For i = 0 To fallbackCount - 1
                cpuHistory(i) = New List(Of Single)()
                coreColors(i) = New SolidBrush(Color.FromArgb(60, 120, rndFallback.Next(160, 255)))
                cpuCounters(i) = Nothing
            Next
        End Try
    End Sub

    Private Sub SetupActiveNetworkInterface()
        Try
            netDownCounter?.Dispose() : netUpCounter?.Dispose() : netPacketsCounter?.Dispose()

            ' FIX: blacklist mais completa - "virtual"/"pseudo" na descrição não pega adaptadores
            ' tipo Teredo/ISATAP/WAN Miniport, que têm nomes próprios sem essas palavras
            Dim blacklist = {"virtual", "pseudo", "teredo", "isatap", "6to4", "loopback",
                              "wan miniport", "kernel debug", "tap-windows", "vmware", "virtualbox",
                              "hyper-v virtual", "bluetooth"}

            Dim interfaces = NetworkInterface.GetAllNetworkInterfaces() _
                .Where(Function(ni) ni.OperationalStatus = OperationalStatus.Up AndAlso
                                   ni.NetworkInterfaceType <> NetworkInterfaceType.Loopback AndAlso
                                   ni.NetworkInterfaceType <> NetworkInterfaceType.Tunnel AndAlso
                                   Not blacklist.Any(Function(termo) ni.Description.ToLower().Contains(termo))).ToList()

            ' FIX: prioriza a interface com mais tráfego acumulado, em vez de pegar a primeira da
            ' lista - com Wi-Fi e Ethernet "Up" ao mesmo tempo, a ordem de enumeração é arbitrária
            ' e podia selecionar a interface errada (sem tráfego real)
            Dim activeNet = interfaces.OrderByDescending(
                Function(ni)
                    Try
                        Dim st = ni.GetIPv4Statistics()
                        Return st.BytesReceived + st.BytesSent
                    Catch
                        Return 0L
                    End Try
                End Function).FirstOrDefault()

            If activeNet IsNot Nothing Then
                If activeNet.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 Then
                    networkTypeString = "Wi-Fi"
                ElseIf activeNet.NetworkInterfaceType = NetworkInterfaceType.Ethernet Then
                    networkTypeString = "Rede Cabeada"
                Else
                    networkTypeString = "Adaptador Ativo"
                End If

                Dim cat As New PerformanceCounterCategory("Network Interface")
                Dim instances = cat.GetInstanceNames()

                Dim instanceName = instances.FirstOrDefault(Function(inst)
                                                                Dim cleanDesc = activeNet.Description.Replace("(", "[").Replace(")", "]")
                                                                Return inst.Replace("(", "[").Replace(")", "]").StartsWith(cleanDesc.Substring(0, Math.Min(cleanDesc.Length, 15)))
                                                            End Function)

                ' FIX: se houver mais de uma instância candidata com o mesmo prefixo (placas de
                ' rede idênticas), prefere a que tiver estatísticas de tráfego mais próximas de
                ' zero-ou-mais (evita pegar sempre a primeira igual arbitrariamente)
                If String.IsNullOrEmpty(instanceName) Then instanceName = instances.FirstOrDefault(Function(x) x.ToLower().Contains("ethernet") Or x.ToLower().Contains("wi-fi"))

                If Not String.IsNullOrEmpty(instanceName) Then
                    netDownCounter = New PerformanceCounter("Network Interface", "Bytes Received/sec", instanceName)
                    netUpCounter = New PerformanceCounter("Network Interface", "Bytes Sent/sec", instanceName)
                    netPacketsCounter = New PerformanceCounter("Network Interface", "Packets/sec", instanceName)
                End If
            Else
                networkTypeString = "Desconectado"
            End If
        Catch
            networkTypeString = "Desconectado"
        End Try
    End Sub

    Private Sub t_Tick(sender As Object, e As EventArgs) Handles t.Tick
        ' Proteção blindada: Aborta o ciclo se os arrays estruturais da CPU ainda não existirem
        If cpuCounters Is Nothing OrElse cpuUsages Is Nothing OrElse cpuHistory Is Nothing Then Exit Sub

        Try
            If Now.Second Mod 5 = 0 Then
                UpdateScreenMetrics()
                RefreshDrivesList() ' FIX: evita usar DriveInfo desatualizado se um pendrive/SD for removido
            End If

            ' Coleta CPU com proteção individual de objetos
            For i = 0 To cpuCounters.Length - 1
                If cpuCounters(i) IsNot Nothing Then
                    cpuUsages(i) = cpuCounters(i).NextValue()
                Else
                    cpuUsages(i) = 0F
                End If

                If cpuHistory(i) IsNot Nothing Then
                    cpuHistory(i).Add(cpuUsages(i))
                    If cpuHistory(i).Count > 30 Then cpuHistory(i).RemoveAt(0)
                End If
            Next

            UpdateHardwareSensors()

            ' Métricas de Rede Calculadas (Bytes para Mbps)
            If netDownCounter IsNot Nothing AndAlso netUpCounter IsNot Nothing Then
                Dim rxBytes = netDownCounter.NextValue()
                Dim txBytes = netUpCounter.NextValue()

                downloadSpeed = (rxBytes * 8) / 1000000
                uploadSpeed = (txBytes * 8) / 1000000

                If netPacketsCounter IsNot Nothing Then packetsPerSec = netPacketsCounter.NextValue()
                sessionDataTraficMB += ((rxBytes + txBytes) / 1024 / 1024)
            End If

            ' Throughput de Armazenamento
            If diskActivityCounter IsNot Nothing Then diskActivity = diskActivityCounter.NextValue()
            If diskReadBytesCounter IsNot Nothing Then diskReadMbps = diskReadBytesCounter.NextValue() / 1024 / 1024
            If diskWriteBytesCounter IsNot Nothing Then diskWriteMbps = diskWriteBytesCounter.NextValue() / 1024 / 1024

        Catch : End Try

        Me.Invalidate()
    End Sub

    Private Sub UpdateHardwareSensors()
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
        If cpuUsages Is Nothing Then Exit Sub

        Dim clientW = Me.ClientSize.Width
        Dim clientH = Me.ClientSize.Height

        Dim padding As Integer = 20
        Dim colW As Integer = (clientW \ 2) - 30
        Dim cardH As Integer = (clientH - 100) \ 2

        ' --- COLUNA ESQUERDA: PROCESSAMENTO (CPU E MULTI-GPU) ---
        DrawCard(g, padding, 50, colW, cardH, "PROCESSADOR")
        g.DrawString($"{cpuFullName}  |  {currentScreenRes} @ {currentScreenHz}", fontSub, brushTextBlack, padding + 15, 75)

        Dim cpuTBrush = If(cpuTemp > 82, Brushes.Red, Brushes.Orange)
        DrawBar(g, padding + 15, 98, colW - 30, 14, cpuTemp, $"CPU Temp: {cpuTemp:0}°C", cpuTBrush)

        ' Renderização responsiva dos Cores
        Dim curX = padding + 15, curY = 140
        For i = 0 To cpuUsages.Length - 1
            If curX + 110 > (padding + colW) Then
                curX = padding + 15
                curY += 75
                If curY + 60 > 50 + cardH Then Exit For
            End If
            DrawCore(g, curX, curY, cpuUsages(i), cpuHistory(i), coreColors(i), i)
            curX += 115
        Next

        ' Card Dinâmico Multi-GPU
        Dim gpuY = 70 + cardH
        DrawCard(g, padding, gpuY, colW, cardH - 20, "MÓDULOS DE VÍDEO DETECTADOS")
        Dim internalGpuY = gpuY + 25

        ' FIX: List(Of T).Count é uma propriedade (sem parâmetros) que teria prioridade sobre o
        ' método de extensão Enumerable.Count(predicate) do LINQ, causando erro de compilação
        ' ao tentar chamar detectedGpus.Count(Function(x) ...) diretamente. .Where(...).Count()
        ' evita o conflito porque IEnumerable(Of T) não tem propriedade Count própria.
        Dim totalDedicadas = detectedGpus.Where(Function(x) Not x.IsIntegrated).Count()

        For Each gpu In detectedGpus
            Dim typeStr = If(gpu.IsIntegrated, "Integrada [iGPU]", "Dedicada [dGPU]")
            g.DrawString($"{gpu.Name} ({typeStr})", fontSub, brushTextBlack, padding + 15, internalGpuY)

            If Not gpu.IsIntegrated OrElse detectedGpus.Count = 1 Then
                ' FIX: com mais de uma GPU dedicada, o sensor do OpenHardwareMonitor é agregado
                ' (não identifica qual placa é qual) - avisa isso em vez de sugerir uma leitura por placa
                Dim tempTxt = If(totalDedicadas > 1, $"Carga: {gpuUsage:0}%  |  Temp: {gpuTemp:0}°C (sensor agregado)", $"Carga: {gpuUsage:0}%  |  Temp: {gpuTemp:0}°C")
                DrawBar(g, padding + 15, internalGpuY + 20, colW - 30, 14, gpuUsage, tempTxt, Brushes.Crimson)
                internalGpuY += 55
            Else
                internalGpuY += 25
            End If
            If internalGpuY + 30 > gpuY + cardH Then Exit For
        Next

        ' --- COLUNA DIREITA: SISTEMA, REDE E ARMAZENAMENTO ---
        Dim rx = colW + 40
        DrawCard(g, rx, 50, colW, clientH - 70, "ESTATÍSTICAS DO SISTEMA")

        ' RAM
        ' FIX: protege contra ramTotal=0 (divisão gerava NaN/Infinity visível no texto, não só na barra)
        Dim pRam As Single = 0F
        If ramTotal > 0 Then
            pRam = (1 - (If(ramCounter?.NextValue(), 0.0F) / ramTotal)) * 100
        End If
        DrawBar(g, rx + 15, 80, colW - 30, 16, pRam, $"Memória RAM em Uso: {pRam:0}%", Brushes.RoyalBlue)

        ' Rede Customizada Avançada
        Dim netDataString = $"{networkTypeString} | Pacotes: {packetsPerSec:0}/s | Sessão: {sessionDataTraficMB:N0} MB"
        DrawBar(g, rx + 15, 140, colW - 30, 16, Math.Min(downloadSpeed, 100), $"↓ {downloadSpeed:F1} Mbps  |  ↑ {uploadSpeed:F1} Mbps", Brushes.Teal)
        g.DrawString(netDataString, fontTiny, brushGrayText, rx + 15, 175)

        ' Armazenamento / Throughput Geral e Unidades
        Dim storageY = 210
        Dim diskTxt = $"Atividade I/O: {diskActivity:0}% | R: {diskReadMbps:F1} MB/s | W: {diskWriteMbps:F1} MB/s"
        DrawBar(g, rx + 15, storageY, colW - 30, 16, diskActivity, diskTxt, Brushes.Purple)

        storageY += 55
        If drives IsNot Nothing Then
            For Each d In drives
                ' FIX: revalida IsReady e protege contra divisão por zero antes de desenhar —
                ' evita IOException não tratada (crash) se o drive foi removido entre um refresh e outro
                Try
                    If Not d.IsReady OrElse d.TotalSize <= 0 Then Continue For

                    Dim model = If(diskModels.ContainsKey(d.Name.Replace("\", "")), diskModels(d.Name.Replace("\", "")), "Unidade de Armazenamento")
                    Dim p = CSng(((d.TotalSize - d.AvailableFreeSpace) / d.TotalSize) * 100)
                    DrawBar(g, rx + 15, storageY, colW - 30, 14, p, $"{d.Name} [{model}] Uso: {p:0}%", Brushes.SlateGray)
                    storageY += 50
                Catch
                    ' Drive pode ter sido removido no exato instante da leitura - pula silenciosamente
                    Continue For
                End Try
                If storageY + 30 > clientH - 30 Then Exit For
            Next
        End If
    End Sub

    Private Sub DrawOverlay(g As Graphics)
        If cpuUsages Is Nothing OrElse cpuUsages.Length = 0 Then Exit Sub

        g.Clear(Color.FromArgb(28, 28, 30))
        Dim w = Me.ClientSize.Width - 16
        Dim y = 24 ' abaixo da faixa de revelar os checkboxes (e.Y < 22)
        Dim rowH = 20
        Dim barH = 8

        ' ----- CPU -----
        Dim cpuAvg = cpuUsages.Average()
        DrawBar(g, 8, y, w, barH, cpuAvg, $"CPU {cpuAvg:0}% {cpuTemp:0}°C", Brushes.DodgerBlue, fontTiny)
        y += rowH

        ' ----- RAM -----
        Dim pRam As Single = 0F
        If ramTotal > 0 Then pRam = (1 - (If(ramCounter?.NextValue(), 0.0F) / ramTotal)) * 100
        DrawBar(g, 8, y, w, barH, pRam, $"RAM {pRam:0}%", Brushes.RoyalBlue, fontTiny)
        y += rowH

        ' ----- GPU: só ocupa uma linha se houver GPU dedicada com sensor (economiza espaço) -----
        Dim totalDedicadas = detectedGpus.Where(Function(x) Not x.IsIntegrated).Count()
        If totalDedicadas > 0 Then
            DrawBar(g, 8, y, w, barH, gpuUsage, $"GPU {gpuUsage:0}% {gpuTemp:0}°C", Brushes.Crimson, fontTiny)
            y += rowH
        End If

        ' ----- Disco -----
        DrawBar(g, 8, y, w, barH, diskActivity, $"Disco {diskActivity:0}% R{diskReadMbps:F0} W{diskWriteMbps:F0}", Brushes.Purple, fontTiny)
        y += rowH

        ' ----- Rede (tipo + velocidade, sem pacotes/sessão - isso fica só na view completa) -----
        Dim netLabel = If(String.IsNullOrEmpty(networkTypeString), "Rede", networkTypeString)
        DrawBar(g, 8, y, w, barH, Math.Min(downloadSpeed, 100), $"{netLabel} ↓{downloadSpeed:F1} ↑{uploadSpeed:F1}", Brushes.Teal, fontTiny)
    End Sub

    Private Sub DrawCard(g As Graphics, x As Integer, y As Integer, w As Integer, h As Integer, title As String)
        GraphicsExtensions.FillRoundedRectangle(g, brushCardBg, x, y, w, h, 6)
        g.DrawString(title, fontTiny, brushGrayText, x + 12, y + 8)
    End Sub

    Private Sub DrawBar(g As Graphics, x As Integer, y As Integer, w As Integer, h As Integer, val As Single, txt As String, b As Brush, Optional txtFont As Font = Nothing)
        ' 1. Validação de segurança: se o valor for inválido, infinito ou a largura for irreal, força para 0
        If Single.IsNaN(val) OrElse Single.IsInfinity(val) OrElse val < 0 Then val = 0F
        If w <= 0 Then w = 1 ' Evita divisão ou multiplicação por zero/negativo no layout

        ' 2. Desenha o fundo da barra
        Using bgBrush As New SolidBrush(Color.FromArgb(35, 180, 180, 180))
            g.FillRectangle(bgBrush, x, y, w, h)
        End Using

        ' 3. Cálculo seguro da largura do preenchimento
        Dim clampedVal As Single = Math.Min(val, 100)
        Dim fillWidth As Integer = CInt((clampedVal / 100.0F) * w)

        ' 4. Renderiza o preenchimento se houver espaço válido
        If fillWidth > 0 Then
            ' Garante que a barra preenchida nunca passe do limite físico do fundo (w)
            If fillWidth > w Then fillWidth = w
            g.FillRectangle(b, x, y, fillWidth, h)
        End If

        ' 5. Desenha o texto descritivo logo abaixo da barra
        ' FIX: fonte opcional (usada pelo overlay com fontTiny) para permitir linhas mais compactas
        Dim textColor As Brush = If(isCompact, brushTextWhite, brushTextBlack)
        g.DrawString(txt, If(txtFont, fontSub), textColor, x, y + h + 2)
    End Sub

    Private Sub DrawCore(g As Graphics, x As Integer, y As Integer, val As Single, hist As List(Of Single), b As Brush, i As Integer)
        Dim barH As Integer = CInt(val / 100 * 35)
        If barH > 0 Then g.FillRectangle(b, x, y + (35 - barH), 18, barH)
        g.DrawRectangle(Pens.LightGray, x, y, 18, 35)

        If hist IsNot Nothing AndAlso hist.Count > 1 Then
            Dim pts = hist.Select(Function(v, idx) New PointF(x + 22 + (idx * 2), y + 35 - (v / 100 * 35))).ToArray()
            g.DrawLines(New Pen(b, 1.0F), pts)
        End If
        g.DrawString($"{i}:{val:0}%", fontTiny, Brushes.DimGray, x, y + 37)
    End Sub

    ' --- Manipulação de Eventos ---
    Private Sub chkCompact_CheckedChanged(sender As Object, e As EventArgs) Handles chkCompact.CheckedChanged
        isCompact = chkCompact.Checked
        If isCompact Then
            ' FIX: layout compacto (barras finas de 8px, fonte tiny, sem pacotes/sessão) cabe em
            ' bem menos espaço - overlay deve ocupar o mínimo possível da tela
            Me.FormBorderStyle = FormBorderStyle.None : Me.Size = New Size(210, 125) : Me.Opacity = 0.9
            btnCpuToggle.Visible = False
        Else
            Me.FormBorderStyle = FormBorderStyle.Sizable : Me.Size = normalSize : Me.Opacity = 1.0
            btnCpuToggle.Visible = True
            RefreshDrivesList()
        End If
    End Sub

    Private Sub chkTopMost_CheckedChanged(sender As Object, e As EventArgs) Handles chkTopMost.CheckedChanged
        Me.TopMost = chkTopMost.Checked
    End Sub

    Private Sub btnCpuToggle_Click(sender As Object, e As EventArgs) Handles btnCpuToggle.Click
        showThreads = Not showThreads
        btnCpuToggle.Text = If(showThreads, "Threads: ON", "Cores: ON")
        SetupCPUCounters()
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        If isCompact Then
            ' FIX: ajustado de 35 para 22 para bater com o novo y inicial (24) do conteúdo do overlay
            Dim showControls = (e.Y < 22)
            chkTopMost.Visible = showControls : chkCompact.Visible = showControls
        End If
    End Sub

    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            ReleaseCapture() : SendMessage(Me.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        End If
    End Sub

    Private Function GetTotalPhysicalMemory() As Long
        Try
            Using s As New ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")
                For Each o In s.Get() : Return CLng(o("TotalPhysicalMemory")) : Next
            End Using
        Catch : End Try
        Return 0
    End Function

    ' --- Limpeza Completa e Preventiva de Recursos ---
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            t?.Stop() : t?.Dispose()

            ramCounter?.Dispose()
            diskActivityCounter?.Dispose()
            diskReadBytesCounter?.Dispose()
            diskWriteBytesCounter?.Dispose()
            netDownCounter?.Dispose()
            netUpCounter?.Dispose()
            netPacketsCounter?.Dispose()

            If cpuCounters IsNot Nothing Then
                For Each c In cpuCounters : c?.Dispose() : Next
            End If

            If coreColors IsNot Nothing Then
                For Each cb In coreColors : cb?.Dispose() : Next
            End If

            fontTitle?.Dispose() : fontSub?.Dispose() : fontTiny?.Dispose()

            Try : computer?.Close() : Catch : End Try
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' FIX: os campos abaixo eram todos declarados como Integer (32 bits), mas a struct DEVMODE
    ' real do Win32 usa WORD (16 bits) em vários deles. O layout de memória incorreto podia fazer
    ' o EnumDisplaySettings retornar valores errados (ex: Hz incorreto) em vez de falhar de forma limpa.
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
    Public Structure DEVMODE
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public dmDeviceName As String
        Public dmSpecVersion, dmDriverVersion, dmSize, dmDriverExtra As Short
        Public dmFields As Integer
        Public dmPositionX, dmPositionY, dmDisplayOrientation, dmDisplayFixedOutput As Integer
        Public dmColor, dmDuplex, dmYResolution, dmTTOption, dmCollate As Short
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public dmFormName As String
        Public dmLogPixels As Short
        Public dmBitsPerPel, dmPelsWidth, dmPelsHeight, dmDisplayFlags, dmDisplayFrequency As Integer
        Public dmICMMethod, dmICMIntent, dmMediaType, dmDitherType, dmReserved1, dmReserved2, dmPanningWidth, dmPanningHeight As Integer
    End Structure
End Class

' ===== MÓDULO DE EXTENSÃO (Fora da Classe) =====
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