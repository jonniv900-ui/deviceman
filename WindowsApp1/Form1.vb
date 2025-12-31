Imports System.Globalization
Imports System.IO
Imports System.Management
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Drawing.Icon
Imports System.Diagnostics
Imports System.Linq
Imports System.Drawing




Public Class Form1

    Private Enum SHSTOCKICONID
        SIID_DRIVEFIXED = &H8
        SIID_DRIVEREMOVABLE = &H9
        SIID_DRIVENET = &HA
        SIID_DRIVERAM = &HB
        SIID_COMPUTER = &H10
        SIID_PROCESSOR = &H16
        SIID_MEMORY = &H17
        SIID_USB = &H28
        SIID_WARNING = &H4F
        SIID_INFO = &H4E
    End Enum

    Private Const SHGSI_ICON As UInteger = &H100
    Private Const SHGSI_SMALLICON As UInteger = &H1
    Private Const SHGSI_LARGEICON As UInteger = &H0
    Private statusTimer As Timer

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure SHSTOCKICONINFO
        Public cbSize As UInteger
        Public hIcon As IntPtr
        Public iSysImageIndex As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public szPath As String
    End Structure
    '<DllImport("shell32.dll")>
    <DllImport("user32.dll")>
    Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
    End Function
    Private Shared Function SHGetStockIconInfo(
    ByVal siid As SHSTOCKICONID,
    ByVal uFlags As UInteger,
    ByRef psii As SHSTOCKICONINFO
) As Integer
    End Function
    ' Declare objetos para monitoramento
    Private cpuCounter As PerformanceCounter
    Private ramCounter As PerformanceCounter

    Private Sub InitStatusBar()
        ' CPU %
        cpuCounter = New PerformanceCounter("Processor", "% Processor Time", "_Total")
        cpuCounter.NextValue() ' Inicializa

        ' Memória disponível
        ramCounter = New PerformanceCounter("Memory", "Available Bytes")

        ' Atualiza StatusBar imediatamente
        UpdateStatusBar()

        ' Timer para atualizar a cada 1s
        statusTimer = New Timer()
        statusTimer.Interval = 1000
        AddHandler statusTimer.Tick, AddressOf UpdateStatusBar
        statusTimer.Start()
    End Sub

    Private Sub UpdateStatusBar()
        ' CPU %
        Dim cpuUso = cpuCounter.NextValue()

        ' RAM total
        Dim cs As ManagementObject = New ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem").Get().Cast(Of ManagementObject)().FirstOrDefault()
        Dim ramTotal As Long = If(cs IsNot Nothing, CLng(cs("TotalPhysicalMemory")), 0)

        ' RAM disponível
        Dim ramDisp As Long = CLng(ramCounter.NextValue())
        Dim ramUsada As Long = ramTotal - ramDisp
        Dim ramPercent As Double = If(ramTotal > 0, ramUsada / ramTotal * 100, 0)

        ' Processador e RAM total (primeiro CPU)
        Dim cpuNome As String = ""
        For Each cpu As ManagementObject In New ManagementObjectSearcher("SELECT Name FROM Win32_Processor").Get()
            cpuNome = WmiStr(cpu, "Name")
            Exit For
        Next

        ' Atualiza StatusBar
        statusbar1.Text = $"CPU: {cpuNome} | Uso CPU: {cpuUso:0.0}% | RAM: {FormatSize2(ramTotal)} | Usada: {FormatSize2(ramUsada)} ({ramPercent:0.0}%)"
    End Sub

    ' Função auxiliar de formatação de bytes
    Private Function FormatSize2(ByVal bytes As Long) As String
        If bytes >= 1024 ^ 3 Then
            Return $"{bytes / 1024 ^ 3:0.00} GB"
        ElseIf bytes >= 1024 ^ 2 Then
            Return $"{bytes / 1024 ^ 2:0.00} MB"
        ElseIf bytes >= 1024 Then
            Return $"{bytes / 1024:0.00} KB"
        Else
            Return $"{bytes} B"
        End If
    End Function

    Private Function GetStockIcon(id As SHSTOCKICONID, small As Boolean) As Icon

        Dim info As New SHSTOCKICONINFO()
        info.cbSize = CUInt(Marshal.SizeOf(info))

        Dim flags As UInteger = SHGSI_ICON
        flags = flags Or If(small, SHGSI_SMALLICON, SHGSI_LARGEICON)

        SHGetStockIconInfo(id, flags, info)

        If info.hIcon = IntPtr.Zero Then Return Nothing
        Dim ico = Icon.FromHandle(info.hIcon)
        Dim clone = CType(ico.Clone(), Icon)
        DestroyIcon(info.hIcon)
        Return clone

    End Function




    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LVdetalhes.View = View.Details
        LVdetalhes.Columns.Add("Propriedade", 220)
        LVdetalhes.Columns.Add("Valor", 450)
        BuildTree()
        TVdispositivos.SelectedNode = TVdispositivos.Nodes(0)
        InitStatusBar()


    End Sub
    Private Sub BuildTree()

        TVdispositivos.Nodes.Clear()

        ' ===== ROOTS =====
        Dim rootResumo = TVdispositivos.Nodes.Add("Resumo do Sistema")
        rootResumo.Tag = "RESUMO"
        rootResumo.ImageKey = "info"
        rootResumo.SelectedImageKey = "info"

        Dim rootDisp = TVdispositivos.Nodes.Add("Dispositivos")
        rootDisp.Tag = "DISPOSITIVOS"
        rootDisp.ImageKey = "folder"
        rootDisp.SelectedImageKey = "folder"

        Dim rootSistema = TVdispositivos.Nodes.Add("Sistema")
        rootSistema.Tag = "SISTEMA"
        rootSistema.ImageKey = "folder"        ' Ícone do software/sistema
        rootSistema.SelectedImageKey = "folder"

        ' Sistema Operacional
        rootSistema.Nodes.Add("Sistema Operacional").Tag = "OS"

        ' Programas Instalados
        rootSistema.Nodes.Add("Programas Instalados").Tag = "PROGRAMAS"

        ' Drivers
        rootSistema.Nodes.Add("Drivers Instalados").Tag = "DRIVERS"

        ' Atualizações
        rootSistema.Nodes.Add("Atualizações").Tag = "UPDATES"



        ' ===== PROCESSADOR =====
        Dim cpuNode = rootDisp.Nodes.Add("Processador")
        cpuNode.Tag = "CPU"
        cpuNode.ImageKey = "cpu"
        cpuNode.SelectedImageKey = "cpu"

        ' ===== MEMÓRIA RAM =====
        Dim ramNode = rootDisp.Nodes.Add("Memória RAM")
        ramNode.Tag = "RAM"
        ramNode.ImageKey = "ram"
        ramNode.SelectedImageKey = "ram"

        Dim ramResumo = ramNode.Nodes.Add("Resumo")
        ramResumo.Tag = "RAM_RESUMO"
        ramResumo.ImageKey = "report"
        ramResumo.SelectedImageKey = "report"

        For Each m As ManagementObject In New ManagementObjectSearcher(
        "SELECT BankLabel FROM Win32_PhysicalMemory").Get()

            Dim banco = WmiStr(m, "BankLabel", "Slot desconhecido")
            Dim slotNode = ramNode.Nodes.Add(banco)
            slotNode.Tag = "RAM_SLOT"
            slotNode.ImageKey = "chip"
            slotNode.SelectedImageKey = "chip"
        Next

        ' ===== PLACA-MÃE =====
        Dim mbNode = rootDisp.Nodes.Add("Placa-mãe")
        mbNode.Tag = "MB"
        mbNode.ImageKey = "mb"
        mbNode.SelectedImageKey = "mb"
        ' ===== GPU / VÍDEO =====
        Dim gpuNode = rootDisp.Nodes.Add("Placa de Vídeo")
        gpuNode.Tag = "GPU"
        gpuNode.ImageKey = "gpu"
        gpuNode.SelectedImageKey = "gpu"

        For Each gpu As ManagementObject In New ManagementObjectSearcher(
    "SELECT Name FROM Win32_VideoController").Get()

            Dim nome = WmiStr(gpu, "Name", "GPU desconhecida")
            Dim n = gpuNode.Nodes.Add(nome)
            n.Tag = "GPU_ITEM"
            n.ImageKey = "gpu"
            n.SelectedImageKey = "gpu"
        Next
        ' ===== REDE =====
        Dim redeNode = rootDisp.Nodes.Add("Rede")
        redeNode.Tag = "REDE"
        redeNode.ImageKey = "network"          ' Chave do ícone no ImageList
        redeNode.SelectedImageKey = "network"

        ' ===== Sub-nodes para cada adaptador =====
        For Each nic As ManagementObject In New ManagementObjectSearcher(
    "SELECT * FROM Win32_NetworkAdapter").Get()

            Dim nome = WmiStr(nic, "Name")
            Dim status As String

            If WmiBool(nic, "NetEnabled") Then
                status = "Ativo"
            ElseIf String.IsNullOrEmpty(WmiStr(nic, "PNPDeviceID")) Then
                status = "Sem driver"
            Else
                status = "Desativado"
            End If

            Dim node = redeNode.Nodes.Add($"{nome} ({status})")
            node.Tag = WmiStr(nic, "DeviceID")
            node.ImageKey = "nic"              ' Ícone específico para NIC
            node.SelectedImageKey = "nic"
        Next




        ' ===== DISPOSITIVOS PnP =====
        Dim pnpRoot = rootDisp.Nodes.Add("Dispositivos Plug and Play")
        pnpRoot.Tag = "PNP"
        pnpRoot.ImageKey = "pnp"
        pnpRoot.SelectedImageKey = "pnp"

        Dim barramentos As New Dictionary(Of String, TreeNode)

        For Each dev As ManagementObject In New ManagementObjectSearcher(
    "SELECT * FROM Win32_PnPEntity WHERE Present = TRUE AND PNPDeviceID IS NOT NULL").Get()

            Dim pnpId As String = WmiStr(dev, "PNPDeviceID")
            If Not IsRealPnPDevice(pnpId) Then Continue For

            Dim nome As String = WmiStr(dev, "Name", "Dispositivo desconhecido")
            If String.IsNullOrEmpty(nome) Then Continue For

            Dim barramento As String = GetBusFromPNP(pnpId)

            ' cria nó do barramento se não existir
            If Not barramentos.ContainsKey(barramento) Then
                Dim busNode = pnpRoot.Nodes.Add(barramento)
                busNode.Tag = "PNP_BUS"

                Dim iconKey = GetBusIconKey(barramento)
                busNode.ImageKey = iconKey
                busNode.SelectedImageKey = iconKey

                barramentos(barramento) = busNode
            End If

            ' adiciona o dispositivo
            Dim estado As String = GetDeviceState(dev)

            Dim devNode = barramentos(barramento).Nodes.Add(nome)
            devNode.Tag = pnpId

            ' ícone base do dispositivo (ex: barramento ou genérico)
            Dim baseIcon As String = GetBusIconKey(barramento)
            devNode.ImageKey = baseIcon
            devNode.SelectedImageKey = baseIcon

            ' aplica ícone de estado SOMENTE se não for OK
            Dim stateIcon As String = GetStateIconKey(estado)
            If Not String.IsNullOrEmpty(stateIcon) Then
                devNode.ImageKey = stateIcon
                devNode.SelectedImageKey = stateIcon
            End If


        Next


        rootResumo.Expand()
        rootDisp.Expand()


    End Sub


    Private Function WmiStr(
    obj As ManagementBaseObject,
    prop As String,
    Optional def As String = ""
) As String

        Try
            If obj Is Nothing Then Return def
            If Not obj.Properties.Cast(Of PropertyData)().
               Any(Function(p) p.Name.Equals(prop, StringComparison.OrdinalIgnoreCase)) Then
                Return def
            End If

            Dim v = obj(prop)
            If v Is Nothing Then Return def

            Return v.ToString()

        Catch ex As ManagementException
            Return def
        Catch ex As Exception
            Return def
        End Try

    End Function



    Private Sub tvDispositivos_AfterSelect(
        sender As Object,
        e As TreeViewEventArgs) Handles TVdispositivos.AfterSelect

        LVdetalhes.Items.Clear()
        Dim tag = If(e.Node.Tag, "").ToString()

        Select Case tag

            Case "RESUMO"
                ShowResumo()

            Case "CPU"
                ShowCPU()

            Case "MB"
                ShowMotherboard()

            Case "RAM"
                ShowRAMResumo()

            Case "RAM_RESUMO"
                ShowRAMResumo()

            Case "RAM_SLOT"
                ShowRAMSlot(e.Node.Text)

            Case "GPU"
                ShowGPUResumo()

            Case "GPU_ITEM"
                ShowGPU(e.Node.Text)

            Case "REDE"
                ' Node raiz
                ShowNetworkResumo()

            Case "SISTEMA"
                        ' Node raiz: talvez resumo geral de software
                        ShowSistemaResumo()

            Case "OS"
                ShowOSDetails()

                    Case "PROGRAMAS"
                        ShowInstalledPrograms()

                    Case "DRIVERS"
                        ShowInstalledDrivers()

                    Case "UPDATES"
                        ShowSystemUpdates()


                    Case Else
                        ' PnP → Tag contém DeviceID
                        If CStr(e.Node.Parent?.Tag) <> "PNP" AndAlso
                       CStr(e.Node.Parent?.Parent?.Tag) = "PNP" Then
                            ShowPnPDevice(e.Node.Tag.ToString())
                            ShowPnPDriver(e.Node.Tag.ToString())
                        End If

                        ' Rede → clique em adaptador específico
                        If CStr(e.Node.Parent?.Tag) = "REDE" Then
                            ShowNetworkAdapter(e.Node.Tag.ToString())
                        End If
                End Select
    End Sub
    Private Sub ShowSistemaResumo()
        LVdetalhes.Items.Clear()

        ' ===== SISTEMA OPERACIONAL =====
        For Each os As ManagementObject In New ManagementObjectSearcher(
        "SELECT Caption, Version, BuildNumber FROM Win32_OperatingSystem").Get()

            AddDetail("Sistema Operacional", $"{WmiStr(os, "Caption")} {WmiStr(os, "Version")} (Build {WmiStr(os, "BuildNumber")})")
            Exit For
        Next

        ' ===== PROGRAMAS INSTALADOS =====
        Dim totalProgramas As Integer = 0
        Dim uninstallKeys = {
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    }

        For Each keyPath In uninstallKeys
            Using regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(keyPath)
                If regKey IsNot Nothing Then
                    totalProgramas += regKey.GetSubKeyNames().Length
                End If
            End Using
        Next

        AddDetail("Programas Instalados", $"{totalProgramas} programas encontrados")

        ' ===== DRIVERS =====
        Dim totalDrivers As Integer = 0
        For Each drv As ManagementObject In New ManagementObjectSearcher(
        "SELECT DeviceName FROM Win32_PnPSignedDriver").Get()
            totalDrivers += 1
        Next

        AddDetail("Drivers Instalados", $"{totalDrivers} drivers encontrados")

        ' ===== ATUALIZAÇÕES =====
        Dim totalUpdates As Integer = 0
        For Each update As ManagementObject In New ManagementObjectSearcher(
        "SELECT HotFixID FROM Win32_QuickFixEngineering").Get()
            totalUpdates += 1
        Next

        AddDetail("Atualizações Aplicadas", $"{totalUpdates} atualizações instaladas")
    End Sub

    Private Sub ShowOSDetails()
        For Each os As ManagementObject In New ManagementObjectSearcher(
        "SELECT Caption, Version, BuildNumber, InstallDate FROM Win32_OperatingSystem").Get()

            Dim dataInstalacao = WmiStr(os, "InstallDate")
            Dim dataFmt = If(dataInstalacao.Length >= 8,
                         $"{dataInstalacao.Substring(0, 4)}-{dataInstalacao.Substring(4, 2)}-{dataInstalacao.Substring(6, 2)}",
                         "Desconhecida")

            AddDetail("Sistema Operacional", $"{WmiStr(os, "Caption")} {WmiStr(os, "Version")} (Build {WmiStr(os, "BuildNumber")})")
            AddDetail("Caminho do Sistema", Environment.GetFolderPath(Environment.SpecialFolder.Windows))
            AddDetail("Data de Instalação", dataFmt)
            Exit For
        Next
    End Sub
    Private Sub ShowInstalledPrograms()
        ' Consulta registry (64 e 32-bit)
        Dim uninstallKeys = {
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    }

        For Each keyPath In uninstallKeys
            Using regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(keyPath)
                If regKey IsNot Nothing Then
                    For Each subKeyName In regKey.GetSubKeyNames()
                        Using subKey = regKey.OpenSubKey(subKeyName)
                            Dim nome = subKey.GetValue("DisplayName")
                            Dim versao = subKey.GetValue("DisplayVersion")
                            If nome IsNot Nothing Then
                                AddDetail("Programa", $"{nome} {If(versao, "")}")
                            End If
                        End Using
                    Next
                End If
            End Using
        Next
    End Sub
    Private Sub ShowInstalledDrivers()
        For Each drv As ManagementObject In New ManagementObjectSearcher(
        "SELECT DeviceName, DriverVersion, Manufacturer, DriverDate FROM Win32_PnPSignedDriver").Get()

            AddDetail("Driver",
                  $"{WmiStr(drv, "DeviceName")} — {WmiStr(drv, "Manufacturer")} v{WmiStr(drv, "DriverVersion")} ({WmiStr(drv, "DriverDate")})")
        Next
    End Sub
    Private Sub ShowSystemUpdates()
        For Each update As ManagementObject In New ManagementObjectSearcher(
        "SELECT HotFixID, InstalledOn FROM Win32_QuickFixEngineering").Get()

            AddDetail("Atualização",
                  $"{WmiStr(update, "HotFixID")} — {WmiStr(update, "InstalledOn")}")
        Next
    End Sub

    Private Sub ShowNetworkResumo()
        LVdetalhes.Items.Clear()

        For Each nic As ManagementObject In New ManagementObjectSearcher(
        "SELECT * FROM Win32_NetworkAdapter").Get()

            Dim nome = WmiStr(nic, "Name")
            Dim status As String

            If WmiBool(nic, "NetEnabled") Then
                status = "Ativo"
            ElseIf String.IsNullOrEmpty(WmiStr(nic, "PNPDeviceID")) Then
                status = "Sem driver"
            Else
                status = "Desativado"
            End If

            Dim tipo = WmiStr(nic, "AdapterType")
            AddDetail($"{nome}", $"{tipo} — {status}")
        Next
    End Sub

    Private Sub ShowNetworkAdapter(deviceID As String)
        LVdetalhes.Items.Clear()

        Dim query = $"SELECT * FROM Win32_NetworkAdapter WHERE DeviceID='{deviceID}'"
        For Each nic As ManagementObject In New ManagementObjectSearcher(query).Get()
            Dim nome = WmiStr(nic, "Name")
            Dim status As String = If(WmiBool(nic, "NetEnabled"), "Ativo", "Desativado")
            Dim tipo = WmiStr(nic, "AdapterType")
            Dim mac = WmiStr(nic, "MACAddress")
            Dim speed = WmiLng(nic, "Speed")

            ' IPs → consulta separada via Win32_NetworkAdapterConfiguration
            Dim ipList As New List(Of String)
            For Each cfg As ManagementObject In New ManagementObjectSearcher(
            $"SELECT * FROM Win32_NetworkAdapterConfiguration WHERE Index={WmiInt(nic, "Index")} AND IPEnabled=True").Get()
                Dim ips = TryCast(cfg("IPAddress"), String())
                If ips IsNot Nothing Then ipList.AddRange(ips)
            Next

            AddDetail("Nome", nome)
            AddDetail("Status", status)
            AddDetail("Tipo", tipo)
            AddDetail("MAC", mac)
            AddDetail("Velocidade", If(speed > 0, $"{speed / 1_000_000} Mbps", "Desconhecida"))
            AddDetail("IPs", If(ipList.Count > 0, String.Join(" | ", ipList), "Nenhum"))
        Next
    End Sub

    Private Sub ShowGPUResumo()

        For Each gpu As ManagementObject In New ManagementObjectSearcher(
        "SELECT * FROM Win32_VideoController").Get()

            AddDetail("Nome", WmiStr(gpu, "Name"))
            AddDetail("Fabricante", WmiStr(gpu, "AdapterCompatibility"))
            AddDetail("Memória de Vídeo", FormatBytes(WmiLng(gpu, "AdapterRAM")))
            AddDetail("Driver", WmiStr(gpu, "DriverVersion"))
            AddDetail("Status", WmiStr(gpu, "Status"))
           

        Next

    End Sub
    Private Sub ShowGPU(nomeGpu As String)

        For Each gpu As ManagementObject In New ManagementObjectSearcher(
        "SELECT * FROM Win32_VideoController").Get()

            If WmiStr(gpu, "Name") = nomeGpu Then

                AddDetail("Nome", WmiStr(gpu, "Name"))
                AddDetail("Fabricante", WmiStr(gpu, "AdapterCompatibility"))
                AddDetail("Descrição", WmiStr(gpu, "Description"))
                AddDetail("Memória de Vídeo", FormatBytes(WmiLng(gpu, "AdapterRAM")))
                AddDetail("Resolução Atual", $"{WmiInt(gpu, "CurrentHorizontalResolution")} x {WmiInt(gpu, "CurrentVerticalResolution")}")
                AddDetail("Frequência", WmiStr(gpu, "CurrentRefreshRate") & " Hz")
                AddDetail("Driver", WmiStr(gpu, "DriverVersion"))
                AddDetail("Data do Driver", WmiStr(gpu, "DriverDate"))
                AddDetail("PNP Device ID", WmiStr(gpu, "PNPDeviceID"))


                Exit Sub
            End If
        Next

    End Sub
    Private Function WmiLng(obj As ManagementBaseObject, prop As String) As Long
        If obj(prop) Is Nothing Then Return 0
        Return CLng(obj(prop))
    End Function

    Private Function WmiInt(obj As ManagementBaseObject, prop As String) As Integer
        If obj(prop) Is Nothing Then Return 0
        Return CInt(obj(prop))
    End Function
    Private Function FormatBytes(bytes As Long) As String

        If bytes <= 0 Then Return "Desconhecido"

        Dim gb = bytes / 1024.0 / 1024.0 / 1024.0

        If gb >= 1 Then
            Return gb.ToString("0.##") & " GB"
        End If

        Return (bytes / 1024.0 / 1024.0).ToString("0") & " MB"

    End Function


    Private Sub ShowPnPDriver(deviceId As String)

        For Each d As ManagementObject In New ManagementObjectSearcher(
        "SELECT * FROM Win32_PnPSignedDriver WHERE DeviceID='" &
        deviceId.Replace("\", "\\") & "'").Get()

            AddDetail("Driver", WmiStr(d, "DriverName"))
            AddDetail("Versão", WmiStr(d, "DriverVersion"))
            AddDetail("Fabricante", WmiStr(d, "DriverProviderName"))
            AddDetail("INF", WmiStr(d, "InfName"))
            AddDetail("Data", WmiStr(d, "DriverDate"))
            AddDetail("Arquivo", WmiStr(d, "DriverName"))
            AddDetail("Diretório", WmiStr(d, "DriverPath"))

            Exit For
        Next

    End Sub



    Private Sub ShowResumo()
        LVdetalhes.Items.Clear()

        ' ===== CPU =====
        Dim vt As String = "?"
        Dim slat As String = "?"
        Dim hyperv As String = "?"

        For Each cpu As ManagementObject In New ManagementObjectSearcher(
        "SELECT Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, " &
        "VirtualizationFirmwareEnabled, SecondLevelAddressTranslationExtensions " &
        "FROM Win32_Processor").Get()

            Dim nome = WmiStr(cpu, "Name")
            Dim cores = WmiInt(cpu, "NumberOfCores")
            Dim threads = WmiInt(cpu, "NumberOfLogicalProcessors")
            Dim freq = WmiInt(cpu, "MaxClockSpeed") / 1000.0

            AddDetail("Processador", $"{nome} ({cores}/{threads} @ {freq:0.00} GHz)")

            vt = If(WmiBool(cpu, "VirtualizationFirmwareEnabled"), "✓", "✗")
            slat = If(WmiBool(cpu, "SecondLevelAddressTranslationExtensions"), "✓", "✗")

            Exit For
        Next

        ' ---- Hyper-V ----
        For Each cs As ManagementObject In New ManagementObjectSearcher(
        "SELECT HypervisorPresent FROM Win32_ComputerSystem").Get()

            hyperv = If(WmiBool(cs, "HypervisorPresent"), "✓", "✗")
            Exit For
        Next

        AddDetail("Virtualização", $"VT-x / AMD-V {vt} | SLAT {slat} | Hyper-V {hyperv}")

        ' ===== INSTRUÇÕES DO PROCESSADOR =====
        For Each cpu As ManagementObject In New ManagementObjectSearcher("SELECT * FROM Win32_Processor").Get()
            Dim inst As New List(Of String)
            If WmiBool(cpu, "PAEEnabled") Then inst.Add("PAE")
            If WmiBool(cpu, "SecondLevelAddressTranslationExtensions") Then inst.Add("SLAT")
            If WmiBool(cpu, "VirtualizationFirmwareEnabled") Then inst.Add("VT-x/AMD-V")
            AddDetail("Instruções CPU", If(inst.Count > 0, String.Join(", ", inst), "Não reportadas"))
            Exit For
        Next

        ' ===== MEMÓRIA =====
        Dim totalRam As Long = 0
        Dim freqMax As Integer = 0
        Dim tipoMem As String = ""
        Dim slotsOcupados As Integer = 0
        Dim slotsTotal As Integer = 0

        For Each cs As ManagementObject In New ManagementObjectSearcher(
        "SELECT TotalPhysicalMemory FROM Win32_ComputerSystem").Get()
            totalRam = WmiLng(cs, "TotalPhysicalMemory")
        Next

        For Each m As ManagementObject In New ManagementObjectSearcher(
        "SELECT Speed, MemoryType FROM Win32_PhysicalMemory").Get()
            freqMax = Math.Max(freqMax, WmiInt(m, "Speed"))
            If tipoMem = "" Then tipoMem = MemoryType(WmiInt(m, "MemoryType"))
            slotsOcupados += 1
        Next

        For Each arr As ManagementObject In New ManagementObjectSearcher(
        "SELECT MemoryDevices FROM Win32_PhysicalMemoryArray").Get()
            slotsTotal = WmiInt(arr, "MemoryDevices")
        Next

        AddDetail("Memória", $"{FormatSize(totalRam)} ({freqMax} MHz | {tipoMem} | {slotsOcupados}/{slotsTotal})")

        ' ===== SISTEMA OPERACIONAL =====
        For Each os As ManagementObject In New ManagementObjectSearcher(
        "SELECT Caption, Version, ServicePackMajorVersion, BuildNumber, InstallDate FROM Win32_OperatingSystem").Get()

            Dim spNum = WmiInt(os, "ServicePackMajorVersion")
            Dim spTxt = If(spNum > 0, $"SP {spNum}", "Sem Service Pack")

            ' Data de instalação
            Dim installRaw = WmiStr(os, "InstallDate")
            Dim dtInstall As String = ""
            If installRaw.Length >= 14 Then
                Dim tempDt As DateTime
                If DateTime.TryParseExact(installRaw.Substring(0, 14), "yyyyMMddHHmmss", Nothing, Globalization.DateTimeStyles.None, tempDt) Then
                    dtInstall = tempDt.ToShortDateString()
                End If
            End If

            AddDetail("Sistema Operacional", $"{WmiStr(os, "Caption")} {WmiStr(os, "Version")} (Build {WmiStr(os, "BuildNumber")}) — {spTxt}")
            AddDetail("Caminho do Sistema Operacional", Environment.GetFolderPath(Environment.SpecialFolder.Windows))
            AddDetail("Data de Instalação do Sistema Operacional", dtInstall)

            Exit For
        Next

        ' ===== ARMAZENAMENTO =====
        For Each d As ManagementObject In New ManagementObjectSearcher(
        "SELECT Model, Size, MediaType, InterfaceType FROM Win32_DiskDrive").Get()
            Dim tipo = DiskType(d)
            AddDetail($"Armazenamento {tipo}", $"{WmiStr(d, "Model")} — {FormatSize(WmiLng(d, "Size"))}")
        Next

        ' ===== GPU =====
        Dim gpus As New List(Of String)
        For Each gpu As ManagementObject In New ManagementObjectSearcher(
        "SELECT Name, AdapterRAM FROM Win32_VideoController").Get()
            Dim nome = WmiStr(gpu, "Name")
            Dim vram = FormatBytes(WmiLng(gpu, "AdapterRAM"))
            If nome <> "" Then gpus.Add($"{nome} ({vram})")
        Next
        If gpus.Count > 0 Then AddDetail("Placa(s) de Vídeo", String.Join(" | ", gpus))

        ' ===== PLACA-MÃE =====
        For Each b As ManagementObject In New ManagementObjectSearcher(
        "SELECT Manufacturer, Product, SerialNumber FROM Win32_BaseBoard").Get()
            Dim marca = WmiStr(b, "Manufacturer")
            Dim modelo = WmiStr(b, "Product")
            Dim serie = WmiStr(b, "SerialNumber")
            Dim info = $"{marca} {modelo}".Trim()
            If serie <> "" AndAlso serie.ToLower() <> "default string" Then info &= $" (S/N: {serie})"
            AddDetail("Placa-mãe", info)
            Exit For
        Next

        ' ===== BIOS =====
        For Each bios As ManagementObject In New ManagementObjectSearcher(
        "SELECT Manufacturer, SMBIOSBIOSVersion, ReleaseDate FROM Win32_BIOS").Get()
            Dim fabricante = WmiStr(bios, "Manufacturer")
            Dim versao = WmiStr(bios, "SMBIOSBIOSVersion")
            Dim dataRaw = WmiStr(bios, "ReleaseDate")
            Dim dataFmt As String = ""
            If dataRaw.Length >= 8 Then
                dataFmt = $"{dataRaw.Substring(0, 4)}-{dataRaw.Substring(4, 2)}-{dataRaw.Substring(6, 2)}"
            End If
            Dim info = $"{fabricante} v{versao}"
            If dataFmt <> "" Then info &= $" ({dataFmt})"
            AddDetail("BIOS", info)
            Exit For
        Next

        ' ===== TPM =====
        Try
            Dim scope As New ManagementScope("root\CIMV2\Security\MicrosoftTpm")
            scope.Connect()
            Dim query As New ObjectQuery("SELECT * FROM Win32_Tpm")
            For Each tpm As ManagementObject In New ManagementObjectSearcher(scope, query).Get()
                Dim versao = WmiStr(tpm, "SpecVersion")
                AddDetail("TPM", $"Presente (v{versao})")
                Exit For
            Next
        Catch
            AddDetail("TPM", "Não presente")
        End Try

        ' ===== SECURE BOOT =====
        Try
            Dim scope As New ManagementScope("root\Microsoft\Windows\HardwareManagement")
            scope.Connect()
            Dim query As New ObjectQuery("SELECT * FROM MSFT_SecureBoot")
            For Each sb As ManagementObject In New ManagementObjectSearcher(scope, query).Get()
                Dim ativo As Boolean = False
                If sb.Properties("SecureBootEnabled") IsNot Nothing AndAlso sb("SecureBootEnabled") IsNot Nothing Then
                    ativo = CBool(sb("SecureBootEnabled"))
                End If
                AddDetail("Secure Boot", If(ativo, "Ativo ✓", "Desativado ✗"))
                Exit For
            Next
        Catch
            AddDetail("Secure Boot", "Não suportado / Legacy BIOS")
        End Try

        ' ===== MODO DE BOOT =====
        AddDetail("Modo de Boot", If(IsUEFI(), "UEFI", "Legacy"))

        ' ===== PLACAS DE REDE =====
        For Each nic As ManagementObject In New ManagementObjectSearcher(
        "SELECT * FROM Win32_NetworkAdapter WHERE NetEnabled=True").Get()

            Dim nome = WmiStr(nic, "Name")
            Dim mac = WmiStr(nic, "MACAddress")
            Dim ips As New List(Of String)

            For Each cfg As ManagementObject In New ManagementObjectSearcher(
            $"SELECT * FROM Win32_NetworkAdapterConfiguration WHERE MACAddress='{mac}'").Get()
                Dim ipArr = cfg("IPAddress")
                If ipArr IsNot Nothing Then
                    For Each ip In CType(ipArr, String())
                        ips.Add(ip)
                    Next
                End If
            Next

            AddDetail($"Rede: {nome}", $"MAC: {mac} | IP: {String.Join(", ", ips)}")
        Next

        ' ===== NOME DO COMPUTADOR =====
        AddDetail("Nome do Computador", Environment.MachineName)
    End Sub

    Private Function WmiBool(obj As ManagementBaseObject, prop As String) As Boolean
        If obj Is Nothing Then Return False

        Try
            If obj.Properties(prop) Is Nothing Then Return False
            If obj.Properties(prop).Value Is Nothing Then Return False

            Return Convert.ToBoolean(obj.Properties(prop).Value)
        Catch
            Return False
        End Try
    End Function
    Private Function IsUEFI() As Boolean
        Try
            ' Procura no registro se existe a chave de boot UEFI
            Using key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Control")
                If key IsNot Nothing Then
                    Dim firmwareType = key.GetValue("PEFirmwareType")
                    ' 1 = BIOS, 2 = UEFI
                    Return firmwareType IsNot Nothing AndAlso Convert.ToInt32(firmwareType) = 2
                End If
            End Using
        Catch
        End Try
        Return False
    End Function

    Private Function DiskType(d As ManagementBaseObject) As String

        Dim iface = WmiStr(d, "InterfaceType").ToUpper()
        Dim media = WmiStr(d, "MediaType").ToUpper()

        If iface.Contains("USB") Then Return "USB"
        If media.Contains("SOLID") Or media.Contains("SSD") Then Return "SSD"

        Return "HDD"

    End Function




    Private Sub ShowCPU()

        LVdetalhes.Items.Clear()

        For Each cpu As ManagementObject In New ManagementObjectSearcher(
        "SELECT * FROM Win32_Processor").Get()

            ' ===== IDENTIFICAÇÃO =====
            AddDetail("Nome", WmiStr(cpu, "Name"))
            Dim rawVendor = WmiStr(cpu, "Manufacturer")
            AddDetail("Fabricante", NormalizeCpuVendor(rawVendor))
            AddDetail("ID do Processador", WmiStr(cpu, "ProcessorId"))

            ' ===== TOPOLOGIA =====
            AddDetail("Núcleos Físicos", WmiStr(cpu, "NumberOfCores"))
            AddDetail("Threads", WmiStr(cpu, "NumberOfLogicalProcessors"))

            ' ===== CLOCK =====
            AddDetail("Clock Atual (MHz)", WmiStr(cpu, "CurrentClockSpeed"))
            AddDetail("Clock Máximo (MHz)", WmiStr(cpu, "MaxClockSpeed"))

            ' ===== ARQUITETURA =====
            AddDetail("Arquitetura", If(WmiStr(cpu, "AddressWidth") = "64", "64 bits", "32 bits"))
            AddDetail("Família", WmiStr(cpu, "Family"))
            AddDetail("Modelo", WmiStr(cpu, "Model"))
            AddDetail("Stepping", WmiStr(cpu, "Stepping"))

            ' ===== RECURSOS =====
            AddDetail("Hyper-Threading",
            If(CInt(WmiStr(cpu, "NumberOfLogicalProcessors", "0")) >
               CInt(WmiStr(cpu, "NumberOfCores", "0")),
               "Sim", "Não"))

            AddDetail("DEP / NX", WmiStr(cpu, "ExecuteDisableBitAvailable"))

            ' ===== VIRTUALIZAÇÃO =====
            AddDetail("Virtualização Suportada", WmiStr(cpu, "VMMonitorModeExtensions"))
            AddDetail("Virtualização Ativa no BIOS", WmiStr(cpu, "VirtualizationFirmwareEnabled"))
            AddDetail("SLAT (EPT/RVI)", WmiStr(cpu, "SecondLevelAddressTranslationExtensions"))

            ' ===== CACHE =====
            AddDetail("Cache L2 (KB)", WmiStr(cpu, "L2CacheSize"))
            AddDetail("Cache L3 (KB)", WmiStr(cpu, "L3CacheSize"))

            Exit For ' normalmente só existe um CPU
        Next

    End Sub
    Private Function NormalizeCpuVendor(vendor As String) As String

        If String.IsNullOrWhiteSpace(vendor) Then Return "Desconhecido"

        vendor = vendor.ToLowerInvariant()

        If vendor.Contains("genuineintel") OrElse vendor.Contains("intel") Then
            Return "Intel Corporation"
        End If

        If vendor.Contains("authenticamd") OrElse vendor.Contains("advanced micro devices") OrElse vendor.Contains("amd") Then
            Return "Advanced Micro Devices, Inc."
        End If

        If vendor.Contains("qualcomm") Then
            Return "Qualcomm Technologies, Inc."
        End If

        If vendor.Contains("apple") Then
            Return "Apple Inc."
        End If

        Return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(vendor)
    End Function

    Private Sub ShowRAMSlot(bank As String)

        For Each m In New ManagementObjectSearcher(
        "SELECT * FROM Win32_PhysicalMemory").Get()

            If WmiStr(m, "BankLabel") = bank Then
                AddDetail("Slot", bank)
                AddDetail("Capacidade", FormatSize(CLng(m("Capacity"))))
                AddDetail("Tipo", MemoryType(CInt(m("SMBIOSMemoryType"))))
                AddDetail("Frequência", m("Speed").ToString() & " MHz")
                AddDetail("Fabricante", m("Manufacturer").ToString())
                AddDetail("Serial", m("SerialNumber").ToString())
            End If

        Next

    End Sub

    Private Sub ShowRAMResumo()

        Dim totalSlots As Integer = 0
        Dim slotsUsados As Integer = 0
        Dim totalMem As Long = 0
        Dim freq As Integer = 0

        For Each m In New ManagementObjectSearcher(
        "SELECT * FROM Win32_PhysicalMemory").Get()

            totalSlots += 1
            slotsUsados += 1
            totalMem += CLng(m("Capacity"))
            freq = Math.Max(freq, CInt(m("Speed")))
        Next

        AddDetail("Total de RAM", FormatSize(totalMem))
        AddDetail("Slots Totais", totalSlots.ToString())
        AddDetail("Slots Ocupados", slotsUsados.ToString())
        AddDetail("Slots Livres", (totalSlots - slotsUsados).ToString())
        AddDetail("Maior Frequência", freq & " MHz")

    End Sub


    Private Function MemoryType(t As Integer) As String
        Select Case t
            Case 20 : Return "DDR"
            Case 21 : Return "DDR2"
            Case 24 : Return "DDR3"
            Case 26 : Return "DDR4"
            Case 34 : Return "DDR5"
            Case Else : Return "Desconhecido"
        End Select
    End Function
    Private Sub ShowPnPDevice(deviceId As String)

        For Each d In New ManagementObjectSearcher(
        "SELECT * FROM Win32_PnPEntity").Get()

            If d("DeviceID").ToString() = deviceId Then

                For Each prop In d.Properties
                    If prop.Value IsNot Nothing Then
                        AddDetail(prop.Name, prop.Value.ToString())
                    End If
                Next

            End If
        Next

    End Sub
    Private Sub ShowDisk(deviceId As String)

        For Each d In New ManagementObjectSearcher(
        "SELECT * FROM Win32_DiskDrive").Get()

            If d("DeviceID").ToString() = deviceId Then
                Dim alerta = d("Status").ToString() <> "OK"

                AddDetail("Modelo", d("Model").ToString(), alerta)
                AddDetail("Interface", d("InterfaceType").ToString())
                AddDetail("Capacidade", FormatSize(CLng(d("Size"))))
                AddDetail("SMART", d("Status").ToString(), alerta)
            End If

        Next

    End Sub
    Private Sub AddDetail(nome As String, valor As String, Optional alerta As Boolean = False)
        Dim item = LVdetalhes.Items.Add(nome)
        item.SubItems.Add(valor)
        If alerta Then item.BackColor = Color.LightSalmon
    End Sub

    Private Function FormatSize(bytes As Long) As String
        Return (bytes / 1024 / 1024 / 1024).ToString("0.00") & " GB"
    End Function



    Private Sub ShowMotherboard()

        LVdetalhes.Items.Clear()

        ' ===== PLACA-MÃE =====
        For Each b As ManagementObject In New ManagementObjectSearcher(
        "SELECT * FROM Win32_BaseBoard").Get()

            AddDetail("Fabricante", WmiStr(b, "Manufacturer"))
            AddDetail("Modelo", WmiStr(b, "Product"))
            AddDetail("Número de Série", WmiStr(b, "SerialNumber"))
            AddDetail("Asset Tag", WmiStr(b, "Tag"))
        Next

        ' ===== BIOS =====
        For Each bios As ManagementObject In New ManagementObjectSearcher(
        "SELECT * FROM Win32_BIOS").Get()

            AddDetail("BIOS Fabricante", WmiStr(bios, "Manufacturer"))
            AddDetail("Versão BIOS", WmiStr(bios, "SMBIOSBIOSVersion"))
            AddDetail("Data BIOS", WmiStr(bios, "ReleaseDate"))
        Next

    End Sub
    Public Sub ExportReport(completo As Boolean)
        ' Criar SaveFileDialog
        Dim sfd As New SaveFileDialog()
        sfd.Filter = "Arquivo HTML|*.html"
        sfd.FileName = "Relatorio.html"

        If sfd.ShowDialog() <> DialogResult.OK Then Exit Sub

        Dim sb As New StringBuilder()

        ' ===== Cabeçalho HTML com Bootstrap e Bootstrap Icons =====
        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html lang=""pt-br"">")
        sb.AppendLine("<head>")
        sb.AppendLine("<meta charset=""UTF-8"">")
        sb.AppendLine("<meta name=""viewport"" content=""width=device-width, initial-scale=1"">")
        sb.AppendLine("<link href=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css"" rel=""stylesheet"">")
        sb.AppendLine("<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css"">")
        sb.AppendLine("<title>Relatório de Sistema</title>")
        sb.AppendLine("</head>")
        sb.AppendLine("<body class=""p-3"">")
        sb.AppendLine("<div class=""container"">")

        ' ===== METADADOS =====
        sb.AppendLine(BuildHtmlMetadata())

        sb.AppendLine("<h2 class=""mb-3"">Relatório de Sistema</h2>")

        ' ===== Percorrer nodes =====
        If completo Then
            For Each node As TreeNode In TVdispositivos.Nodes
                AppendNodeHtmlFull(node, sb)
            Next
        Else
            If TVdispositivos.SelectedNode IsNot Nothing Then
                AppendNodeHtmlFull(TVdispositivos.SelectedNode, sb)
            End If
        End If

        sb.AppendLine("</div>")
        sb.AppendLine("</body>")
        sb.AppendLine("</html>")

        ' Salvar arquivo
        File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8)

        Dim resp = MessageBox.Show(
      $"Relatório salvo em:{vbCrLf}{sfd.FileName}{vbCrLf}{vbCrLf}Deseja abrir agora?",
      "Exportação concluída",
      MessageBoxButtons.YesNo,
      MessageBoxIcon.Information
  )

        If resp = DialogResult.Yes Then
            Process.Start(New ProcessStartInfo With {
                .FileName = sfd.FileName,
                .UseShellExecute = True
            })
        End If

    End Sub


    Private Function BuildHtmlMetadata() As String
        Dim sb As New StringBuilder()

        sb.AppendLine("<div class='card mb-4'>")
        sb.AppendLine("<div class='card-header bg-dark text-white'>")
        sb.AppendLine("<strong>Relatório Técnico do Sistema</strong>")
        sb.AppendLine("</div>")
        sb.AppendLine("<div class='card-body'>")

        sb.AppendLine("<ul class='list-unstyled mb-0'>")
        sb.AppendLine($"<li><strong>Computador:</strong> {Environment.MachineName}</li>")
        sb.AppendLine($"<li><strong>Usuário:</strong> {Environment.UserName}</li>")
        sb.AppendLine($"<li><strong>Sistema Operacional:</strong> {My.Computer.Info.OSFullName}</li>")
        sb.AppendLine($"<li><strong>Arquitetura:</strong> {If(Environment.Is64BitOperatingSystem, "64-bit", "32-bit")}</li>")
        sb.AppendLine($"<li><strong>Data/Hora:</strong> {DateTime.Now:dd/MM/yyyy HH:mm:ss}</li>")
        sb.AppendLine("</ul>")

        sb.AppendLine("</div>")
        sb.AppendLine("</div>")

        Return sb.ToString()
    End Function

    ' ===== Função recursiva para relatório completo com ícones =====
    Private Sub AppendNodeHtmlFull(node As TreeNode, sb As StringBuilder)
        ' ===== Escolher ícone para node =====
        Dim iconHtml As String = ""

        Select Case node.Tag
            Case "CPU"
                ' Detectar fabricante no LVdetalhes
                Dim fabricante As String = ""
                For Each item As ListViewItem In LVdetalhes.Items
                    If item.Text.ToLower().Contains("fabricante") Then
                        fabricante = item.SubItems(1).Text.ToLower()
                        Exit For
                    End If
                Next

                If fabricante.Contains("intel") Then
                    iconHtml = "<img src='https://images.seeklogo.com/logo-png/22/2/intel-logo-png_seeklogo-226413.png' width='32' height='32' class='me-2'>"
                ElseIf fabricante.Contains("amd") Then
                    iconHtml = "<img src='https://images.seeklogo.com/logo-png/0/2/amd-logo-png_seeklogo-7779.png' width='32' height='32' class='me-2'>"
                Else
                    iconHtml = "<i class='bi bi-cpu me-2' style='font-size:32px'></i>"
                End If

            Case "RAM", "RAM_SLOT"
                iconHtml = "<i class='bi bi-memory me-2' style='font-size:32px'></i>"
            Case "GPU", "GPU_ITEM"
                iconHtml = "<i class='bi bi-gpu-card me-2' style='font-size:32px'></i>"
            Case "MB"
                iconHtml = "<i class='bi bi-motherboard me-2' style='font-size:32px'></i>"
            Case "REDE"
                iconHtml = "<i class='bi bi-wifi me-2' style='font-size:32px'></i>"
            Case "BIOS"
                iconHtml = "<i class='bi bi-chip me-2' style='font-size:32px'></i>"
            Case "TPM"
                iconHtml = "<i class='bi bi-lock me-2' style='font-size:32px'></i>"
            Case "STORAGE"
                iconHtml = "<i class='bi bi-hdd me-2' style='font-size:32px'></i>"
            Case Else
                iconHtml = "<i class='bi bi-file-earmark-text me-2' style='font-size:32px'></i>"
        End Select

        sb.AppendLine("<div class=""mb-3"">")
        sb.AppendLine("<h4>" & iconHtml & node.Text & "</h4>")

        ' ===== ListView detalhes se selecionado =====
        If LVdetalhes.Items.Count > 0 AndAlso node.IsSelected Then
            sb.AppendLine("<table class=""table table-bordered table-striped"">")
            sb.AppendLine("<thead class=""table-dark""><tr><th>Propriedade</th><th>Valor</th></tr></thead>")
            sb.AppendLine("<tbody>")
            For Each item As ListViewItem In LVdetalhes.Items
                sb.AppendLine($"<tr><td>{item.Text}</td><td>{item.SubItems(1).Text}</td></tr>")
            Next
            sb.AppendLine("</tbody>")
            sb.AppendLine("</table>")
        End If

        ' ===== Recursão para filhos =====
        If node.Nodes.Count > 0 Then
            For Each child As TreeNode In node.Nodes
                LVdetalhes.Items.Clear()
                TVdispositivos.SelectedNode = child
                tvDispositivos_AfterSelect(Me, New TreeViewEventArgs(child))
                AppendNodeHtmlFull(child, sb)
            Next
        End If

        sb.AppendLine("</div>")
    End Sub



    Private Sub RelatórioCompletoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RelatórioCompletoToolStripMenuItem.Click
        ExportReport(True)

    End Sub

    Private Sub SomenteTelaSelecionadaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SomenteTelaSelecionadaToolStripMenuItem.Click
        ExportReport(False)

    End Sub

    Private Sub GerenciadorDeRecursosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GerenciadorDeRecursosToolStripMenuItem.Click
        SysRes.Show()

    End Sub

    Private Sub GerenciadorDeProcessosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GerenciadorDeProcessosToolStripMenuItem.Click
        tskmgr.Show()

    End Sub

    Private Sub RestaurarDriversToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestaurarDriversToolStripMenuItem.Click
        Dim frm As New Global.WindowsApp2.Form1()
        frm.Show()
    End Sub

    Private Sub RestaurarDriversToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RestaurarDriversToolStripMenuItem1.Click
        Dim frm As New Global.restoreapp.Form1()
        frm.Show()
    End Sub

    '' Nova parte adicionado 30/12/25
    Private Function GetBusFromPNP(pnpId As String) As String
        If String.IsNullOrEmpty(pnpId) Then Return "Outros"

        If pnpId.StartsWith("PCI\", StringComparison.OrdinalIgnoreCase) Then Return "PCI / PCIe"
        If pnpId.StartsWith("USB\", StringComparison.OrdinalIgnoreCase) Then Return "USB"
        If pnpId.StartsWith("ACPI\", StringComparison.OrdinalIgnoreCase) Then Return "ACPI"
        If pnpId.StartsWith("HID\", StringComparison.OrdinalIgnoreCase) Then Return "HID"
        If pnpId.StartsWith("SCSI\", StringComparison.OrdinalIgnoreCase) Then Return "SCSI"
        If pnpId.StartsWith("IDE\", StringComparison.OrdinalIgnoreCase) Then Return "IDE / SATA"
        If pnpId.StartsWith("NVME\", StringComparison.OrdinalIgnoreCase) Then Return "NVMe"

        Return "Outros"
    End Function
    Private Function IsRealPnPDevice(pnpId As String) As Boolean
        If String.IsNullOrEmpty(pnpId) Then Return False

        Dim invalidPrefixes = {
        "ROOT\",
        "SWD\",
        "HTREE\",
        "LEGACY\",
        "VMBUS\"
    }

        For Each prefix In invalidPrefixes
            If pnpId.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If
        Next

        ' Dispositivos reais sempre passam por um barramento físico
        Return True
    End Function
    Private Function GetBusIconKey(bus As String) As String

        Select Case bus
            Case "USB"
                Return "usb"
            Case "PCI / PCIe"
                Return "pci"
            Case "ACPI"
                Return "acpi"
            Case "HID"
                Return "hid"
            Case "SCSI"
                Return "scsi"
            Case "IDE / SATA"
                Return "ide"
            Case "NVMe"
                Return "nvme"
            Case Else
                Return "bus"
        End Select

    End Function



    Private Function GetStateIconKey(state As String) As String
        Select Case state
            Case "NO_DRIVER"
                Return "state_nodriver"

            Case "ERROR"
                Return "state_error"

            Case Else
                ' OK → não força ícone de estado
                Return Nothing
        End Select
    End Function
    Private Function GetDeviceState(dev As ManagementObject) As String

        Dim errCode As Integer = WmiInt(dev, "ConfigManagerErrorCode")
        Dim service As String = WmiStr(dev, "Service")

        ' sem driver
        If String.IsNullOrEmpty(service) OrElse errCode = 28 Then
            Return "NO_DRIVER"
        End If

        ' erro
        If errCode <> 0 Then
            Return "ERROR"
        End If

        Return "OK"

    End Function

End Class