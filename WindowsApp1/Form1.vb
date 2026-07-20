' ============================================================
'  Este projeto é open-source — código-fonte completo disponível
'  publicamente no GitHub. Contribuições, issues e PRs são bem-vindos.
'  Ao editar, prefira nomes/comentários claros e evite hardcode de
'  caminhos, credenciais ou dados específicos de uma única máquina.
' ============================================================
Imports System.Globalization
Imports System.IO
Imports System.Management
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Drawing.Icon
Imports System.Diagnostics
Imports System.Linq
Imports System.Drawing
Imports System.Net
Imports System.Xml.Linq




Public Class Form1

    Private Enum SHSTOCKICONID
        SIID_DRIVEFIXED = &H8
        SIID_DRIVEREMOVABLE = &H9
        SIID_DRIVENET = &HA
        SIID_DRIVERAM = &HB
        SIID_FOLDER = &H3          ' ADICIONADO: Ícone de Pasta genérica
        SIID_DOCASSOC = &H1        ' ADICIONADO: Usado para relatórios/documentos
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

    ' ===== Ícones do ListView de detalhes =====
    Private lvIconList As ImageList
    ' Categoria "corrente" usada pelo AddDetail para escolher o ícone da linha,
    ' setada no início de cada Show* antes das chamadas a AddDetail.
    Private currentIconKey As String = "info"

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure SHSTOCKICONINFO
        Public cbSize As UInteger
        Public hIcon As IntPtr
        Public iSysImageIndex As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public szPath As String
    End Structure

    <DllImport("user32.dll")>
    Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
    End Function

    ' FIX: faltava o DllImport correto (estava comentado / a função ficava sem atributo nenhum)
    <DllImport("shell32.dll")>
    Private Shared Function SHGetStockIconInfo(
    ByVal siid As SHSTOCKICONID,
    ByVal uFlags As UInteger,
    ByRef psii As SHSTOCKICONINFO
) As Integer
    End Function

    ' ===== P/Invoke para anexar/criar um console quando o app roda em modo CLI =====
    Private Const ATTACH_PARENT_PROCESS As Integer = -1

    <DllImport("kernel32.dll")>
    Private Shared Function AttachConsole(dwProcessId As Integer) As Boolean
    End Function

    <DllImport("kernel32.dll")>
    Private Shared Function AllocConsole() As Boolean
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
        Dim ramTotal As Long = 0
        Using searcher As New ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")
            Using results As ManagementObjectCollection = searcher.Get()
                Using cs As ManagementObject = results.Cast(Of ManagementObject)().FirstOrDefault()
                    ramTotal = If(cs IsNot Nothing, WmiLng(cs, "TotalPhysicalMemory"), 0)
                End Using
            End Using
        End Using

        ' RAM disponível
        Dim ramDisp As Long = CLng(ramCounter.NextValue())
        Dim ramUsada As Long = ramTotal - ramDisp
        Dim ramPercent As Double = If(ramTotal > 0, ramUsada / ramTotal * 100, 0)

        ' Processador (primeiro CPU)
        Dim cpuNome As String = ""
        Using searcher As New ManagementObjectSearcher("SELECT Name FROM Win32_Processor"), results As ManagementObjectCollection = searcher.Get()
            For Each cpu As ManagementObject In results
                Using cpu
                    cpuNome = WmiStr(cpu, "Name")
                End Using
                Exit For
            Next
        End Using

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

    ' ===== Monta o ImageList usando ícones nativos do Shell do Windows =====
    ' ===== Monta o ImageList usando ícones nativos do Shell do Windows =====
    Private Sub InitListViewIcons()
        lvIconList = New ImageList()
        lvIconList.ImageSize = New Size(16, 16)
        lvIconList.ColorDepth = ColorDepth.Depth32Bit

        ' Ícones base
        AddStockIconToList("cpu", SHSTOCKICONID.SIID_PROCESSOR)
        AddStockIconToList("ram", SHSTOCKICONID.SIID_MEMORY)
        AddStockIconToList("gpu", SHSTOCKICONID.SIID_COMPUTER)
        AddStockIconToList("mb", SHSTOCKICONID.SIID_COMPUTER)
        AddStockIconToList("os", SHSTOCKICONID.SIID_COMPUTER)
        AddStockIconToList("network", SHSTOCKICONID.SIID_DRIVENET)
        AddStockIconToList("storage", SHSTOCKICONID.SIID_DRIVEFIXED)
        AddStockIconToList("pnp", SHSTOCKICONID.SIID_USB)
        AddStockIconToList("bus", SHSTOCKICONID.SIID_USB)
        AddStockIconToList("info", SHSTOCKICONID.SIID_INFO)
        AddStockIconToList("warning", SHSTOCKICONID.SIID_WARNING)
        AddStockIconToList("folder", SHSTOCKICONID.SIID_FOLDER)

        ' CORREÇÃO CRÍTICA PARA O TREEVIEW: Criação das chaves em falta associando a ícones nativos equivalentes
        AddStockIconToList("report", SHSTOCKICONID.SIID_DOCASSOC)    ' Documento/Relatório
        AddStockIconToList("chip", SHSTOCKICONID.SIID_MEMORY)       ' Chips dos slots de RAM
        AddStockIconToList("nic", SHSTOCKICONID.SIID_DRIVENET)       ' Placas de rede individuais

        ' Vincula a lista de imagens a ambos os controlos
        LVdetalhes.SmallImageList = ImageList1
        TVdispositivos.ImageList = ImageList1
    End Sub

    Private Sub AddStockIconToList(key As String, siid As SHSTOCKICONID)
        Try
            Dim ico = GetStockIcon(siid, True)
            If ico IsNot Nothing AndAlso Not lvIconList.Images.ContainsKey(key) Then
                lvIconList.Images.Add(key, ico)
            End If
        Catch
            ' Se o ícone stock não puder ser obtido (SO antigo, tema custom etc.),
            ' a linha simplesmente fica sem ícone — não deve travar a UI.
        End Try
    End Sub

    ' ===== Define o título do formulário com Processador + Nome da Máquina =====
    Private Sub SetFormTitle()
        Dim cpuNome As String = "Processador desconhecido"

        Try
            Using searcher As New ManagementObjectSearcher("SELECT Name FROM Win32_Processor"), results As ManagementObjectCollection = searcher.Get()
                For Each cpu As ManagementObject In results
                    Using cpu
                        cpuNome = WmiStr(cpu, "Name", cpuNome).Trim()
                    End Using
                    Exit For
                Next
            End Using
        Catch
            ' Mantém o texto padrão caso a consulta WMI falhe
        End Try

        Me.Text = $"{cpuNome} — {Environment.MachineName}"
    End Sub




    ' ===== MODO CLI: se o app receber parâmetros, roda headless (sem exibir a janela) =====
    Protected Overrides Sub OnLoad(e As EventArgs)
        Dim cliArgs = My.Application.CommandLineArgs.ToArray()

        If cliArgs.Length > 0 Then
            ' Evita qualquer flash de janela antes de processar o comando
            Me.Opacity = 0
            Me.ShowInTaskbar = False
            Me.WindowState = FormWindowState.Minimized

            Try
                RunCliMode(cliArgs)
            Finally
                ' FIX: garante que todo o buffer de saída seja escoado para o console antes
                ' do Environment.Exit encerrar o processo abruptamente (sem rodar finalizers),
                ' mesmo que RunCliMode lance uma exceção não tratada no meio do caminho.
                Console.Out.Flush()
                Console.Error.Flush()
            End Try

            Environment.Exit(0)
            Return
        End If

        MyBase.OnLoad(e)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LVdetalhes.View = View.Details
        LVdetalhes.Columns.Add("Propriedade", 220)
        LVdetalhes.Columns.Add("Valor", 450)
        InitListViewIcons()
        SetFormTitle()
        BuildTree()
        TVdispositivos.SelectedNode = TVdispositivos.Nodes(0)
        InitStatusBar()
        CarregarMenuFerramentas()
        ReplicarNoToolStrip()


    End Sub

    ' FIX: libera os PerformanceCounter, o Timer e o ImageList ao fechar o formulário
    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        statusTimer?.Stop()
        statusTimer?.Dispose()
        cpuCounter?.Dispose()
        ramCounter?.Dispose()
        lvIconList?.Dispose()
        MyBase.OnFormClosing(e)
    End Sub

    ' ================================================================
    ' =====================  MODO CLI (console)  ====================
    ' ================================================================

    ' Anexa ao console do processo pai (cmd/PowerShell) ou cria um novo,
    ' e redireciona Console.Out/Error para ele.
    Private Sub EnsureConsole()
        If Not AttachConsole(ATTACH_PARENT_PROCESS) Then
            AllocConsole()
        End If

        Dim stdOut As New StreamWriter(Console.OpenStandardOutput())
        stdOut.AutoFlush = True
        Console.SetOut(stdOut)

        Dim stdErr As New StreamWriter(Console.OpenStandardError())
        stdErr.AutoFlush = True
        Console.SetError(stdErr)
    End Sub

    Private Sub RunCliMode(args As String())
        EnsureConsole()

        Dim asJson = args.Any(Function(a) a.Equals("--json", StringComparison.OrdinalIgnoreCase))
        Dim wantsHelp = args.Any(Function(a) a.Equals("--help", StringComparison.OrdinalIgnoreCase) OrElse a = "-h" OrElse a = "/?")

        If wantsHelp Then
            PrintUsage()
            Return
        End If

        ' ----- Modo relatório -----
        Dim temReport = args.Any(Function(a) a.Equals("--report", StringComparison.OrdinalIgnoreCase) OrElse a.Equals("-report", StringComparison.OrdinalIgnoreCase))
        If temReport Then
            RunCliReport(args, asJson)
            Return
        End If

        ' ----- Modo consulta pontual -----
        ' FIX: aceita o comando com "--", "-" ou sem prefixo nenhum (ex: --network, -network, network)
        Dim categoria = args.FirstOrDefault(Function(a) Not a.Equals("--json", StringComparison.OrdinalIgnoreCase))

        If categoria Is Nothing Then
            Console.Error.WriteLine("Nenhuma categoria reconhecida. Use --help para ver as opções.")
            PrintUsage()
            Return
        End If

        Dim chave = categoria.TrimStart("-"c).ToLowerInvariant()
        Dim tituloCategoria As String

        Select Case chave
            Case "cpu"
                ShowCPU() : tituloCategoria = "Processador"
            Case "ram"
                ShowRAMResumo() : tituloCategoria = "Memória RAM"
            Case "gpu"
                ShowGPUResumo() : tituloCategoria = "Placa de Vídeo"
            Case "mb", "motherboard", "placa-mae"
                ShowMotherboard() : tituloCategoria = "Placa-mãe"
            Case "rede", "network"
                ShowNetworkResumo() : tituloCategoria = "Rede"
            Case "pnp"
                ShowPnpResumo() : tituloCategoria = "Dispositivos Plug and Play"
            Case "os", "sistema"
                ShowSistemaResumo() : tituloCategoria = "Sistema Operacional"
            Case "dispositivos"
                ShowDispositivosResumo() : tituloCategoria = "Dispositivos"
            Case "storage", "armazenamento", "discos"
                ShowStorageResumo() : tituloCategoria = "Armazenamento"
            Case "battery", "bateria"
                ShowBattery() : tituloCategoria = "Bateria"
            Case "resumo", "all", "systeminfo"
                ShowResumo() : tituloCategoria = "Resumo do Sistema"
            Case "biosinfo", "bios"
                ShowBiosInfo() : tituloCategoria = "BIOS"
            Case "checktpm", "tpm"
                ShowTpmInfo() : tituloCategoria = "TPM"
            Case "checkvirtualization", "virtualizacao", "virtualization"
                ShowVirtualizationInfo() : tituloCategoria = "Virtualização"
            Case Else
                Console.Error.WriteLine($"Categoria desconhecida: {categoria}")
                PrintUsage()
                Return
        End Select

        If asJson Then
            Console.WriteLine(BuildJsonFromListView(tituloCategoria))
        Else
            PrintTextTable(tituloCategoria)
        End If
    End Sub

    Private Sub RunCliReport(args As String(), asJson As Boolean)
        BuildTree()

        Dim outPath As String = Nothing
        Dim idxOut = Array.IndexOf(args, "--out")
        If idxOut >= 0 AndAlso idxOut + 1 < args.Length Then
            outPath = args(idxOut + 1)
        End If

        If asJson Then
            Dim json = BuildJsonReport()
            Dim destino = If(outPath, Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "Relatorio.json"))
            Try
                File.WriteAllText(destino, json, Encoding.UTF8)
                Console.WriteLine($"Relatório JSON salvo em: {destino}")
            Catch ex As Exception
                Console.Error.WriteLine($"Erro ao salvar relatório JSON: {ex.Message}")
            End Try
        Else
            ' Reaproveita a exportação HTML já existente (isAuto detecta "--report"/"-report" e não abre diálogos)
            ExportReport(True, False, outPath)
        End If
    End Sub

    Private Sub PrintUsage()
        Console.WriteLine("Uso: <app>.exe [opção] [--json] [--out <arquivo>]")
        Console.WriteLine("(o prefixo -- é opcional: --cpu, -cpu e cpu funcionam igual)")
        Console.WriteLine()
        Console.WriteLine("Consulta pontual (imprime no console):")
        Console.WriteLine("  --resumo, --systeminfo    Resumo geral do sistema")
        Console.WriteLine("  --cpu                     Informações do processador")
        Console.WriteLine("  --ram                     Resumo de memória RAM")
        Console.WriteLine("  --gpu                     Resumo de placa(s) de vídeo")
        Console.WriteLine("  --mb                      Informações da placa-mãe")
        Console.WriteLine("  --biosinfo                Informações do BIOS/UEFI")
        Console.WriteLine("  --checktpm                Verifica presença/estado do TPM")
        Console.WriteLine("  --checkvirtualization     Verifica VT-x/AMD-V, SLAT e Hyper-V")
        Console.WriteLine("  --rede, --network         Resumo dos adaptadores de rede")
        Console.WriteLine("  --pnp                     Resumo de dispositivos Plug and Play")
        Console.WriteLine("  --os                      Resumo do sistema operacional")
        Console.WriteLine("  --dispositivos            Resumo geral de hardware")
        Console.WriteLine("  --storage, --armazenamento  Resumo dos discos")
        Console.WriteLine("  --battery, --bateria      Status e saúde da bateria (se houver)")
        Console.WriteLine()
        Console.WriteLine("Relatório completo:")
        Console.WriteLine("  --report                    Gera o relatório completo (HTML) em Documentos")
        Console.WriteLine("  --report --out <arquivo>    Salva o relatório no caminho informado")
        Console.WriteLine("  --report --json             Gera o relatório em JSON em vez de HTML")
        Console.WriteLine()
        Console.WriteLine("Outras opções:")
        Console.WriteLine("  --json            Formata a saída (consulta ou relatório) como JSON")
        Console.WriteLine("  --help, -h, /?    Mostra esta ajuda")
    End Sub

    Private Sub PrintTextTable(titulo As String)
        Console.WriteLine()
        Console.WriteLine($"=== {titulo.ToUpperInvariant()} ===")
        Console.WriteLine()

        If LVdetalhes.Items.Count = 0 Then
            Console.WriteLine("(nenhuma informação encontrada)")
            Return
        End If

        Dim maxNome = LVdetalhes.Items.Cast(Of ListViewItem)().Max(Function(i) i.Text.Length)
        maxNome = Math.Max(maxNome, "Propriedade".Length)

        Console.WriteLine("Propriedade".PadRight(maxNome) & "  Valor")
        Console.WriteLine(New String("-"c, maxNome) & "  " & New String("-"c, 40))

        For Each item As ListViewItem In LVdetalhes.Items
            Console.WriteLine(item.Text.PadRight(maxNome) & "  " & item.SubItems(1).Text)
        Next

        Console.WriteLine()
    End Sub

    Private Function BuildJsonFromListView(categoria As String) As String
        Dim sb As New StringBuilder()
        sb.Append("{")
        sb.Append($"""categoria"":""{JsonEscape(categoria)}"",")
        sb.Append("""itens"":[")

        Dim first = True
        For Each item As ListViewItem In LVdetalhes.Items
            If Not first Then sb.Append(",")
            first = False
            sb.Append("{")
            sb.Append($"""propriedade"":""{JsonEscape(item.Text)}"",")
            sb.Append($"""valor"":""{JsonEscape(item.SubItems(1).Text)}""")
            sb.Append("}")
        Next

        sb.Append("]}")
        Return sb.ToString()
    End Function

    Private Function BuildJsonReport() As String
        Dim sb As New StringBuilder()
        sb.Append("{")
        sb.Append($"""computador"":""{JsonEscape(Environment.MachineName)}"",")
        sb.Append($"""usuario"":""{JsonEscape(Environment.UserName)}"",")
        sb.Append($"""dataHora"":""{JsonEscape(DateTime.Now.ToString("s"))}"",")
        sb.Append("""categorias"":[")

        Dim first = True
        For Each node As TreeNode In TVdispositivos.Nodes
            If Not first Then sb.Append(",")
            first = False
            sb.Append(BuildJsonNode(node))
        Next

        sb.Append("]}")
        Return sb.ToString()
    End Function

    ' Espelha a mesma recursão de AppendNodeHtmlFull, só que gerando JSON em vez de HTML
    Private Function BuildJsonNode(node As TreeNode) As String
        LVdetalhes.Items.Clear()
        TVdispositivos.SelectedNode = node
        tvDispositivos_AfterSelect(Me, New TreeViewEventArgs(node))

        Dim sb As New StringBuilder()
        sb.Append("{")
        sb.Append($"""nome"":""{JsonEscape(node.Text)}"",")
        sb.Append("""itens"":[")

        Dim first = True
        For Each item As ListViewItem In LVdetalhes.Items
            If Not first Then sb.Append(",")
            first = False
            sb.Append("{")
            sb.Append($"""propriedade"":""{JsonEscape(item.Text)}"",")
            sb.Append($"""valor"":""{JsonEscape(item.SubItems(1).Text)}""")
            sb.Append("}")
        Next
        sb.Append("]")

        If node.Nodes.Count > 0 Then
            sb.Append(",""filhos"":[")
            Dim firstChild = True
            For Each child As TreeNode In node.Nodes
                If Not firstChild Then sb.Append(",")
                firstChild = False
                sb.Append(BuildJsonNode(child))
            Next
            sb.Append("]")
        End If

        sb.Append("}")
        Return sb.ToString()
    End Function

    Private Function JsonEscape(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "").Replace(vbLf, " ")
    End Function

    ' ================================================================
    ' =================  FIM DO MODO CLI (console)  =================
    ' ================================================================

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
        rootSistema.ImageKey = "folder"
        rootSistema.SelectedImageKey = "folder"

        ' CORREÇÃO: Definindo explicitamente as chaves para não pegar ícones aleatórios por índice padrão
        Dim nodeOS = rootSistema.Nodes.Add("Sistema Operacional")
        nodeOS.Tag = "OS" : nodeOS.ImageKey = "os" : nodeOS.SelectedImageKey = "os"

        Dim nodeProg = rootSistema.Nodes.Add("Programas Instalados")
        nodeProg.Tag = "PROGRAMAS" : nodeProg.ImageKey = "os" : nodeProg.SelectedImageKey = "os"

        Dim nodeDrivers = rootSistema.Nodes.Add("Drivers Instalados")
        nodeDrivers.Tag = "DRIVERS" : nodeDrivers.ImageKey = "pnp" : nodeDrivers.SelectedImageKey = "pnp"

        Dim nodeUpdates = rootSistema.Nodes.Add("Atualizações")
        nodeUpdates.Tag = "UPDATES" : nodeUpdates.ImageKey = "info" : nodeUpdates.SelectedImageKey = "info"

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

        Using searcher As New ManagementObjectSearcher("SELECT BankLabel FROM Win32_PhysicalMemory"), results As ManagementObjectCollection = searcher.Get()
            For Each m As ManagementObject In results
                Using m
                    Dim banco = WmiStr(m, "BankLabel", "Slot desconhecido")
                    Dim slotNode = ramNode.Nodes.Add(banco)
                    slotNode.Tag = "RAM_SLOT"
                    slotNode.ImageKey = "chip"       ' Vinculado ao chip de memória adicionado
                    slotNode.SelectedImageKey = "chip"
                End Using
            Next
        End Using

        ' ===== PLACA-MÃE =====
        Dim mbNode = rootDisp.Nodes.Add("Placa-mãe")
        mbNode.Tag = "MB"
        mbNode.ImageKey = "mb"
        mbNode.SelectedImageKey = "mb"

        Dim biosNode = mbNode.Nodes.Add("BIOS / UEFI")
        biosNode.Tag = "BIOS_INFO"
        biosNode.ImageKey = "mb"
        biosNode.SelectedImageKey = "mb"

        Dim tpmNode = mbNode.Nodes.Add("TPM")
        tpmNode.Tag = "TPM_INFO"
        tpmNode.ImageKey = "info"
        tpmNode.SelectedImageKey = "info"

        Dim virtNode = mbNode.Nodes.Add("Virtualização")
        virtNode.Tag = "VIRT_INFO"
        virtNode.ImageKey = "cpu"
        virtNode.SelectedImageKey = "cpu"

        ' ===== GPU / VÍDEO =====
        Dim gpuNode = rootDisp.Nodes.Add("Placa de Vídeo")
        gpuNode.Tag = "GPU"
        gpuNode.ImageKey = "gpu"
        gpuNode.SelectedImageKey = "gpu"

        Using searcher As New ManagementObjectSearcher("SELECT Name FROM Win32_VideoController"), results As ManagementObjectCollection = searcher.Get()
            For Each gpu As ManagementObject In results
                Using gpu
                    Dim nome = WmiStr(gpu, "Name", "GPU desconhecida")
                    Dim n = gpuNode.Nodes.Add(nome)
                    n.Tag = "GPU_ITEM"
                    n.ImageKey = "gpu"
                    n.SelectedImageKey = "gpu"
                End Using
            Next
        End Using

        ' ===== REDE =====
        Dim redeNode = rootDisp.Nodes.Add("Rede")
        redeNode.Tag = "REDE"
        redeNode.ImageKey = "network"
        redeNode.SelectedImageKey = "network"

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapter"), results As ManagementObjectCollection = searcher.Get()
            For Each nic As ManagementObject In results
                Using nic
                    If Not IsWantedNetworkAdapter(nic) Then Continue For

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
                    node.ImageKey = "nic"            ' Vinculado ao ícone individual corrigido
                    node.SelectedImageKey = "nic"
                End Using
            Next
        End Using

        ' ===== ARMAZENAMENTO =====
        Dim storageNode = rootDisp.Nodes.Add("Armazenamento")
        storageNode.Tag = "STORAGE"
        storageNode.ImageKey = "storage"
        storageNode.SelectedImageKey = "storage"

        Using searcher As New ManagementObjectSearcher("SELECT DeviceID, Model FROM Win32_DiskDrive"), results As ManagementObjectCollection = searcher.Get()
            For Each d As ManagementObject In results
                Using d
                    Dim modelo = WmiStr(d, "Model", "Disco desconhecido")
                    Dim n = storageNode.Nodes.Add(modelo)
                    n.Tag = WmiStr(d, "DeviceID")   ' ex: \\.\PHYSICALDRIVE0
                    n.ImageKey = "storage"
                    n.SelectedImageKey = "storage"
                End Using
            Next
        End Using

        ' ===== BATERIA (só cria o nó se o equipamento realmente tiver bateria) =====
        Dim temBateria As Boolean = False
        Using searcher As New ManagementObjectSearcher("SELECT DeviceID FROM Win32_Battery"), results As ManagementObjectCollection = searcher.Get()
            For Each b As ManagementObject In results
                Using b
                    temBateria = True
                End Using
                Exit For
            Next
        End Using

        If temBateria Then
            Dim batNode = rootDisp.Nodes.Add("Bateria")
            batNode.Tag = "BATTERY"
            batNode.ImageKey = "info"
            batNode.SelectedImageKey = "info"
        End If

        ' ===== DISPOSITIVOS PnP =====
        Dim pnpRoot = rootDisp.Nodes.Add("Dispositivos Plug and Play")
        pnpRoot.Tag = "PNP"
        pnpRoot.ImageKey = "pnp"
        pnpRoot.SelectedImageKey = "pnp"

        Dim barramentos As New Dictionary(Of String, TreeNode)

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Present = TRUE AND PNPDeviceID IS NOT NULL"), results As ManagementObjectCollection = searcher.Get()
            For Each dev As ManagementObject In results
                Using dev
                    Dim pnpId = WmiStr(dev, "PNPDeviceID")
                    If Not IsRealPnPDevice(pnpId) Then Continue For

                    Dim nome = WmiStr(dev, "Name", "Dispositivo desconhecido")
                    If String.IsNullOrEmpty(nome) Then Continue For

                    Dim barramento = GetBusFromPNP(pnpId)

                    If Not barramentos.ContainsKey(barramento) Then
                        Dim busNode = pnpRoot.Nodes.Add(barramento)
                        busNode.Tag = "PNP_BUS"

                        Dim iconKey = GetBusIconKey(barramento)
                        busNode.ImageKey = iconKey
                        busNode.SelectedImageKey = iconKey

                        barramentos(barramento) = busNode
                    End If

                    Dim estado = GetDeviceState(dev)
                    Dim devNode = barramentos(barramento).Nodes.Add(nome)
                    devNode.Tag = pnpId

                    Dim baseIcon = GetBusIconKey(barramento)
                    devNode.ImageKey = baseIcon
                    devNode.SelectedImageKey = baseIcon

                    Dim stateIcon = GetStateIconKey(estado)
                    If Not String.IsNullOrEmpty(stateIcon) Then
                        devNode.ImageKey = stateIcon
                        devNode.SelectedImageKey = stateIcon
                    End If
                End Using
            Next
        End Using

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

            Case "DISPOSITIVOS"
                ShowDispositivosResumo()

            Case "PNP"
                ShowPnpResumo()

            Case "PNP_BUS"
                ShowPnpBusResumo(e.Node.Text)

            Case "CPU"
                ShowCPU()

            Case "MB"
                ShowMotherboard()

            Case "BIOS_INFO"
                ShowBiosInfo()

            Case "TPM_INFO"
                ShowTpmInfo()

            Case "VIRT_INFO"
                ShowVirtualizationInfo()

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

            Case "STORAGE"
                ShowStorageResumo()

            Case "BATTERY"
                ShowBattery()

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

                ' Armazenamento → clique em disco específico
                If CStr(e.Node.Parent?.Tag) = "STORAGE" Then
                    ShowDisk(e.Node.Tag.ToString())
                End If
        End Select
    End Sub
    Private Sub ShowSistemaResumo()
        LVdetalhes.Items.Clear()
        currentIconKey = "os"

        ' ===== SISTEMA OPERACIONAL =====
        Using searcher As New ManagementObjectSearcher("SELECT Caption, Version, BuildNumber FROM Win32_OperatingSystem"), results As ManagementObjectCollection = searcher.Get()
            For Each os As ManagementObject In results
                Using os
                    AddDetail("Sistema Operacional", $"{WmiStr(os, "Caption")} {WmiStr(os, "Version")} (Build {WmiStr(os, "BuildNumber")})")
                End Using
                Exit For
            Next
        End Using

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
        Using searcher As New ManagementObjectSearcher("SELECT DeviceName FROM Win32_PnPSignedDriver"), results As ManagementObjectCollection = searcher.Get()
            For Each drv As ManagementObject In results
                Using drv
                    totalDrivers += 1
                End Using
            Next
        End Using

        AddDetail("Drivers Instalados", $"{totalDrivers} drivers encontrados")

        ' ===== ATUALIZAÇÕES =====
        Dim totalUpdates As Integer = 0
        Using searcher As New ManagementObjectSearcher("SELECT HotFixID FROM Win32_QuickFixEngineering"), results As ManagementObjectCollection = searcher.Get()
            For Each update As ManagementObject In results
                Using update
                    totalUpdates += 1
                End Using
            Next
        End Using

        AddDetail("Atualizações Aplicadas", $"{totalUpdates} atualizações instaladas")
    End Sub

    Private Sub ShowOSDetails()
        currentIconKey = "os"
        Using searcher As New ManagementObjectSearcher(
        "SELECT Caption, Version, BuildNumber, InstallDate FROM Win32_OperatingSystem"), results As ManagementObjectCollection = searcher.Get()
            For Each os As ManagementObject In results
                Using os
                    Dim dataInstalacao = WmiStr(os, "InstallDate")
                    Dim dataFmt = If(dataInstalacao.Length >= 8,
                                 $"{dataInstalacao.Substring(0, 4)}-{dataInstalacao.Substring(4, 2)}-{dataInstalacao.Substring(6, 2)}",
                                 "Desconhecida")

                    AddDetail("Sistema Operacional", $"{WmiStr(os, "Caption")} {WmiStr(os, "Version")} (Build {WmiStr(os, "BuildNumber")})")
                    AddDetail("Caminho do Sistema", Environment.GetFolderPath(Environment.SpecialFolder.Windows))
                    AddDetail("Data de Instalação", dataFmt)
                End Using
                Exit For
            Next
        End Using
    End Sub
    Private Sub ShowInstalledPrograms()
        currentIconKey = "os"
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
                            ' FIX: subKey pode vir Nothing (chave removida/inacessível entre a listagem e a abertura)
                            If subKey IsNot Nothing Then
                                Dim nome = subKey.GetValue("DisplayName")
                                Dim versao = subKey.GetValue("DisplayVersion")
                                If nome IsNot Nothing Then
                                    AddDetail("Programa", $"{nome} {If(versao, "")}")
                                End If
                            End If
                        End Using
                    Next
                End If
            End Using
        Next
    End Sub
    Private Sub ShowInstalledDrivers()
        currentIconKey = "pnp"
        Using searcher As New ManagementObjectSearcher(
        "SELECT DeviceName, DriverVersion, Manufacturer, DriverDate FROM Win32_PnPSignedDriver"), results As ManagementObjectCollection = searcher.Get()
            For Each drv As ManagementObject In results
                Using drv
                    AddDetail("Driver",
                          $"{WmiStr(drv, "DeviceName")} — {WmiStr(drv, "Manufacturer")} v{WmiStr(drv, "DriverVersion")} ({WmiStr(drv, "DriverDate")})")
                End Using
            Next
        End Using
    End Sub
    Private Sub ShowSystemUpdates()
        currentIconKey = "os"
        Using searcher As New ManagementObjectSearcher("SELECT HotFixID, InstalledOn FROM Win32_QuickFixEngineering"), results As ManagementObjectCollection = searcher.Get()
            For Each update As ManagementObject In results
                Using update
                    AddDetail("Atualização",
                          $"{WmiStr(update, "HotFixID")} — {WmiStr(update, "InstalledOn")}")
                End Using
            Next
        End Using
    End Sub

    Private Sub ShowNetworkResumo()
        LVdetalhes.Items.Clear()
        currentIconKey = "network"

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapter"), results As ManagementObjectCollection = searcher.Get()
            For Each nic As ManagementObject In results
                Using nic
                    ' FIX: mesmo filtro aplicado na árvore, para o resumo bater com o que é exibido nela
                    If Not IsWantedNetworkAdapter(nic) Then Continue For

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
                End Using
            Next
        End Using
    End Sub

    ' ===== Filtro: adaptadores físicos reais + VMs conhecidas, excluindo miniports do Windows =====
    Private Function IsWantedNetworkAdapter(nic As ManagementObject) As Boolean

        Dim nome = WmiStr(nic, "Name").ToLowerInvariant()
        Dim fabricante = WmiStr(nic, "Manufacturer").ToLowerInvariant()

        ' Blacklist: adaptadores virtuais/software do próprio Windows que não representam hardware real
        Dim blacklist = {
            "wan miniport", "isatap", "teredo", "6to4", "loopback",
            "kernel debug", "microsoft wi-fi direct virtual adapter",
            "microsoft network adapter multiplexor", "microsoft failover cluster virtual adapter",
            "direct parallel", "qos packet scheduler", "bluetooth device (personal area network)",
            "remote ndis", "microsoft wan miniport"
        }
        For Each termo In blacklist
            If nome.Contains(termo) Then Return False
        Next

        ' Whitelist: adaptadores criados por hipervisores/máquinas virtuais conhecidas
        Dim vmWhitelist = {"vmware", "virtualbox", "oracle", "hyper-v virtual ethernet adapter"}
        For Each termo In vmWhitelist
            If nome.Contains(termo) OrElse fabricante.Contains(termo) Then Return True
        Next

        ' Caso contrário, só é aceito se o Windows reportar como adaptador físico de fato
        Return WmiBool(nic, "PhysicalAdapter")

    End Function

    Private Sub ShowNetworkAdapter(deviceID As String)
        LVdetalhes.Items.Clear()
        currentIconKey = "network"

        ' FIX: escapa aspas simples para evitar quebrar a query WQL
        Dim safeId = deviceID.Replace("'", "''")
        Dim query = $"SELECT * FROM Win32_NetworkAdapter WHERE DeviceID='{safeId}'"

        Using searcher As New ManagementObjectSearcher(query), results As ManagementObjectCollection = searcher.Get()
            For Each nic As ManagementObject In results
                Using nic
                    Dim nome = WmiStr(nic, "Name")
                    Dim status As String = If(WmiBool(nic, "NetEnabled"), "Ativo", "Desativado")
                    Dim tipo = WmiStr(nic, "AdapterType")
                    Dim mac = WmiStr(nic, "MACAddress")
                    Dim speed = WmiLng(nic, "Speed")

                    ' IPs → consulta separada via Win32_NetworkAdapterConfiguration
                    Dim ipList As New List(Of String)
                    Using searcher2 As New ManagementObjectSearcher(
                    $"SELECT * FROM Win32_NetworkAdapterConfiguration WHERE Index={WmiInt(nic, "Index")} AND IPEnabled=True"), resultsCfg As ManagementObjectCollection = searcher2.Get()
                        For Each cfg As ManagementObject In resultsCfg
                            Using cfg
                                Dim ips = TryCast(cfg("IPAddress"), String())
                                If ips IsNot Nothing Then ipList.AddRange(ips)
                            End Using
                        Next
                    End Using

                    AddDetail("Nome", nome)
                    AddDetail("Status", status)
                    AddDetail("Tipo", tipo)
                    AddDetail("MAC", mac)
                    AddDetail("Velocidade", If(speed > 0, $"{speed / 1_000_000} Mbps", "Desconhecida"))
                    AddDetail("IPs", If(ipList.Count > 0, String.Join(" | ", ipList), "Nenhum"))
                End Using
            Next
        End Using
    End Sub

    Private Sub ShowGPUResumo()
        currentIconKey = "gpu"

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_VideoController"), results As ManagementObjectCollection = searcher.Get()
            For Each gpu As ManagementObject In results
                Using gpu
                    AddDetail("Nome", WmiStr(gpu, "Name"))
                    AddDetail("Fabricante", WmiStr(gpu, "AdapterCompatibility"))
                    AddDetail("Memória de Vídeo", FormatBytes(WmiLng(gpu, "AdapterRAM")))
                    AddDetail("Driver", WmiStr(gpu, "DriverVersion"))
                    AddDetail("Status", WmiStr(gpu, "Status"))
                End Using
            Next
        End Using

    End Sub
    Private Sub ShowGPU(nomeGpu As String)
        currentIconKey = "gpu"

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_VideoController"), results As ManagementObjectCollection = searcher.Get()
            For Each gpu As ManagementObject In results
                Using gpu
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
                End Using
            Next
        End Using

    End Sub
    Private Function WmiLng(obj As ManagementBaseObject, prop As String) As Long
        If obj Is Nothing Then Return 0
        If obj(prop) Is Nothing Then Return 0
        Return CLng(obj(prop))
    End Function

    Private Function WmiInt(obj As ManagementBaseObject, prop As String) As Integer
        If obj Is Nothing Then Return 0
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
        currentIconKey = "pnp"

        ' FIX: escapa aspas simples além da barra invertida
        Dim safeId = deviceId.Replace("\", "\\").Replace("'", "''")

        Using searcher As New ManagementObjectSearcher(
        "SELECT * FROM Win32_PnPSignedDriver WHERE DeviceID='" & safeId & "'"), results As ManagementObjectCollection = searcher.Get()
            For Each d As ManagementObject In results
                Using d
                    AddDetail("Driver", WmiStr(d, "DriverName"))
                    AddDetail("Versão", WmiStr(d, "DriverVersion"))
                    AddDetail("Fabricante", WmiStr(d, "DriverProviderName"))
                    AddDetail("INF", WmiStr(d, "InfName"))
                    AddDetail("Data", WmiStr(d, "DriverDate"))
                    AddDetail("Arquivo", WmiStr(d, "DriverName"))
                    AddDetail("Diretório", WmiStr(d, "DriverPath"))
                End Using
                Exit For
            Next
        End Using

    End Sub



    Private Sub ShowResumo()
        LVdetalhes.Items.Clear()
        currentIconKey = "info"

        ' ===== CPU =====
        Dim vt As String = "?"
        Dim slat As String = "?"
        Dim hyperv As String = "?"

        Using searcher As New ManagementObjectSearcher(
        "SELECT Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, " &
        "VirtualizationFirmwareEnabled, SecondLevelAddressTranslationExtensions " &
        "FROM Win32_Processor"), results As ManagementObjectCollection = searcher.Get()
            For Each cpu As ManagementObject In results
                Using cpu
                    Dim nome = WmiStr(cpu, "Name")
                    Dim cores = WmiInt(cpu, "NumberOfCores")
                    Dim threads = WmiInt(cpu, "NumberOfLogicalProcessors")
                    Dim freq = WmiInt(cpu, "MaxClockSpeed") / 1000.0

                    AddDetail("Processador", $"{nome} ({cores}/{threads} @ {freq:0.00} GHz)")

                    vt = If(WmiBool(cpu, "VirtualizationFirmwareEnabled"), "✓", "✗")
                    slat = If(WmiBool(cpu, "SecondLevelAddressTranslationExtensions"), "✓", "✗")
                End Using
                Exit For
            Next
        End Using

        ' ---- Hyper-V ----
        Using searcher As New ManagementObjectSearcher("SELECT HypervisorPresent FROM Win32_ComputerSystem"), results As ManagementObjectCollection = searcher.Get()
            For Each cs As ManagementObject In results
                Using cs
                    hyperv = If(WmiBool(cs, "HypervisorPresent"), "✓", "✗")
                End Using
                Exit For
            Next
        End Using

        AddDetail("Virtualização", $"VT-x / AMD-V {vt} | SLAT {slat} | Hyper-V {hyperv}")

        ' ===== INSTRUÇÕES DO PROCESSADOR =====
        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_Processor"), results As ManagementObjectCollection = searcher.Get()
            For Each cpu As ManagementObject In results
                Using cpu
                    Dim inst As New List(Of String)
                    If WmiBool(cpu, "PAEEnabled") Then inst.Add("PAE")
                    If WmiBool(cpu, "SecondLevelAddressTranslationExtensions") Then inst.Add("SLAT")
                    If WmiBool(cpu, "VirtualizationFirmwareEnabled") Then inst.Add("VT-x/AMD-V")
                    AddDetail("Instruções CPU", If(inst.Count > 0, String.Join(", ", inst), "Não reportadas"))
                End Using
                Exit For
            Next
        End Using

        ' ===== MEMÓRIA =====
        Dim totalRam As Long = 0
        Dim freqMax As Integer = 0
        Dim tipoMem As String = ""
        Dim slotsOcupados As Integer = 0
        Dim slotsTotal As Integer = 0

        Using searcher As New ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem"), results As ManagementObjectCollection = searcher.Get()
            For Each cs As ManagementObject In results
                Using cs
                    totalRam = WmiLng(cs, "TotalPhysicalMemory")
                End Using
            Next
        End Using

        Using searcher As New ManagementObjectSearcher("SELECT Speed, MemoryType FROM Win32_PhysicalMemory"), results As ManagementObjectCollection = searcher.Get()
            For Each m As ManagementObject In results
                Using m
                    freqMax = Math.Max(freqMax, WmiInt(m, "Speed"))
                    If tipoMem = "" Then tipoMem = MemoryType(WmiInt(m, "MemoryType"))
                    slotsOcupados += 1
                End Using
            Next
        End Using

        Using searcher As New ManagementObjectSearcher("SELECT MemoryDevices FROM Win32_PhysicalMemoryArray"), results As ManagementObjectCollection = searcher.Get()
            For Each arr As ManagementObject In results
                Using arr
                    slotsTotal = WmiInt(arr, "MemoryDevices")
                End Using
            Next
        End Using

        AddDetail("Memória", $"{FormatSize(totalRam)} ({freqMax} MHz | {tipoMem} | {slotsOcupados}/{slotsTotal})")

        ' ===== SISTEMA OPERACIONAL =====
        Using searcher As New ManagementObjectSearcher(
        "SELECT Caption, Version, ServicePackMajorVersion, BuildNumber, InstallDate FROM Win32_OperatingSystem"), results As ManagementObjectCollection = searcher.Get()
            For Each os As ManagementObject In results
                Using os
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
                End Using
                Exit For
            Next
        End Using

        ' ===== ARMAZENAMENTO =====
        Using searcher As New ManagementObjectSearcher("SELECT Model, Size, MediaType, InterfaceType FROM Win32_DiskDrive"), results As ManagementObjectCollection = searcher.Get()
            For Each d As ManagementObject In results
                Using d
                    Dim tipo = DiskType(d)
                    AddDetail($"Armazenamento {tipo}", $"{WmiStr(d, "Model")} — {FormatSize(WmiLng(d, "Size"))}")
                End Using
            Next
        End Using

        ' ===== GPU =====
        Dim gpus As New List(Of String)
        Using searcher As New ManagementObjectSearcher("SELECT Name, AdapterRAM FROM Win32_VideoController"), results As ManagementObjectCollection = searcher.Get()
            For Each gpu As ManagementObject In results
                Using gpu
                    Dim nome = WmiStr(gpu, "Name")
                    Dim vram = FormatBytes(WmiLng(gpu, "AdapterRAM"))
                    If nome <> "" Then gpus.Add($"{nome} ({vram})")
                End Using
            Next
        End Using
        If gpus.Count > 0 Then AddDetail("Placa(s) de Vídeo", String.Join(" | ", gpus))

        ' ===== PLACA-MÃE =====
        Using searcher As New ManagementObjectSearcher("SELECT Manufacturer, Product, SerialNumber FROM Win32_BaseBoard"), results As ManagementObjectCollection = searcher.Get()
            For Each b As ManagementObject In results
                Using b
                    Dim marca = WmiStr(b, "Manufacturer")
                    Dim modelo = WmiStr(b, "Product")
                    Dim serie = WmiStr(b, "SerialNumber")
                    Dim info = $"{marca} {modelo}".Trim()
                    If serie <> "" AndAlso serie.ToLower() <> "default string" Then info &= $" (S/N: {serie})"
                    AddDetail("Placa-mãe", info)
                End Using
                Exit For
            Next
        End Using

        ' ===== BIOS =====
        Using searcher As New ManagementObjectSearcher("SELECT Manufacturer, SMBIOSBIOSVersion, ReleaseDate FROM Win32_BIOS"), results As ManagementObjectCollection = searcher.Get()
            For Each bios As ManagementObject In results
                Using bios
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
                End Using
                Exit For
            Next
        End Using

        ' ===== TPM =====
        Try
            Dim scope As New ManagementScope("root\CIMV2\Security\MicrosoftTpm")
            scope.Connect()
            Dim query As New ObjectQuery("SELECT * FROM Win32_Tpm")
            Using searcher As New ManagementObjectSearcher(scope, query), results As ManagementObjectCollection = searcher.Get()
                For Each tpm As ManagementObject In results
                    Using tpm
                        Dim versao = WmiStr(tpm, "SpecVersion")
                        AddDetail("TPM", $"Presente (v{versao})")
                    End Using
                    Exit For
                Next
            End Using
        Catch
            AddDetail("TPM", "Não presente")
        End Try

        ' ===== SECURE BOOT =====
        Try
            Dim scope As New ManagementScope("root\Microsoft\Windows\HardwareManagement")
            scope.Connect()
            Dim query As New ObjectQuery("SELECT * FROM MSFT_SecureBoot")
            Using searcher As New ManagementObjectSearcher(scope, query), results As ManagementObjectCollection = searcher.Get()
                For Each sb As ManagementObject In results
                    Using sb
                        Dim ativo As Boolean = False
                        If sb.Properties("SecureBootEnabled") IsNot Nothing AndAlso sb("SecureBootEnabled") IsNot Nothing Then
                            ativo = CBool(sb("SecureBootEnabled"))
                        End If
                        AddDetail("Secure Boot", If(ativo, "Ativo ✓", "Desativado ✗"))
                    End Using
                    Exit For
                Next
            End Using
        Catch
            AddDetail("Secure Boot", "Não suportado / Legacy BIOS")
        End Try

        ' ===== MODO DE BOOT =====
        AddDetail("Modo de Boot", If(IsUEFI(), "UEFI", "Legacy"))

        ' ===== PLACAS DE REDE =====
        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapter WHERE NetEnabled=True"), results As ManagementObjectCollection = searcher.Get()
            For Each nic As ManagementObject In results
                Using nic
                    Dim nome = WmiStr(nic, "Name")
                    Dim mac = WmiStr(nic, "MACAddress")
                    Dim ips As New List(Of String)

                    Using searcher2 As New ManagementObjectSearcher(
                    $"SELECT * FROM Win32_NetworkAdapterConfiguration WHERE MACAddress='{mac.Replace("'", "''")}'"), resultsCfg As ManagementObjectCollection = searcher2.Get()
                        For Each cfg As ManagementObject In resultsCfg
                            Using cfg
                                Dim ipArr = cfg("IPAddress")
                                If ipArr IsNot Nothing Then
                                    For Each ip In CType(ipArr, String())
                                        ips.Add(ip)
                                    Next
                                End If
                            End Using
                        Next
                    End Using

                    AddDetail($"Rede: {nome}", $"MAC: {mac} | IP: {String.Join(", ", ips)}")
                End Using
            Next
        End Using

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
        currentIconKey = "cpu"

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_Processor"), results As ManagementObjectCollection = searcher.Get()
            For Each cpu As ManagementObject In results
                Using cpu
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
                End Using

                Exit For ' normalmente só existe um CPU
            Next
        End Using

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
        currentIconKey = "ram"

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMemory"), results As ManagementObjectCollection = searcher.Get()
            For Each m As ManagementObject In results
                Using m
                    If WmiStr(m, "BankLabel") = bank Then
                        ' FIX: usar os helpers seguros em vez de acesso direto (evita NullReferenceException)
                        AddDetail("Slot", bank)
                        AddDetail("Capacidade", FormatSize(WmiLng(m, "Capacity")))
                        AddDetail("Tipo", MemoryType(WmiInt(m, "SMBIOSMemoryType")))
                        AddDetail("Frequência", WmiStr(m, "Speed") & " MHz")
                        AddDetail("Fabricante", WmiStr(m, "Manufacturer"))
                        AddDetail("Serial", WmiStr(m, "SerialNumber"))
                    End If
                End Using
            Next
        End Using

    End Sub

    Private Sub ShowRAMResumo()
        currentIconKey = "ram"

        Dim totalSlots As Integer = 0
        Dim slotsUsados As Integer = 0
        Dim totalMem As Long = 0
        Dim freq As Integer = 0

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMemory"), results As ManagementObjectCollection = searcher.Get()
            For Each m As ManagementObject In results
                Using m
                    totalSlots += 1
                    slotsUsados += 1
                    totalMem += WmiLng(m, "Capacity")
                    freq = Math.Max(freq, WmiInt(m, "Speed"))
                End Using
            Next
        End Using

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
        currentIconKey = "pnp"

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity"), results As ManagementObjectCollection = searcher.Get()
            For Each d As ManagementObject In results
                Using d
                    If WmiStr(d, "DeviceID") = deviceId Then
                        For Each prop In d.Properties
                            If prop.Value IsNot Nothing Then
                                AddDetail(prop.Name, prop.Value.ToString())
                            End If
                        Next
                    End If
                End Using
            Next
        End Using

    End Sub
    ' ===== Resumo do nó raiz "Armazenamento" =====
    Private Sub ShowStorageResumo()
        currentIconKey = "storage"

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive"), results As ManagementObjectCollection = searcher.Get()
            For Each d As ManagementObject In results
                Using d
                    Dim tipo = DiskType(d)
                    Dim status = WmiStr(d, "Status")
                    AddDetail($"Disco ({tipo})", $"{WmiStr(d, "Model")} — {FormatSize(WmiLng(d, "Size"))} — {status}", status <> "OK")
                End Using
            Next
        End Using
    End Sub

    Private Sub ShowDisk(deviceId As String)
        currentIconKey = "storage"

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive"), results As ManagementObjectCollection = searcher.Get()
            For Each d As ManagementObject In results
                Using d
                    If WmiStr(d, "DeviceID") <> deviceId Then Continue For

                    ' FIX: acesso seguro em vez de d("Status").ToString() direto
                    Dim status = WmiStr(d, "Status")
                    Dim alerta = status <> "OK"

                    AddDetail("Modelo", WmiStr(d, "Model"))
                    AddDetail("Fabricante", WmiStr(d, "Manufacturer"))
                    AddDetail("Interface", WmiStr(d, "InterfaceType"))
                    AddDetail("Tipo de Mídia", DiskType(d))
                    AddDetail("Capacidade", FormatSize(WmiLng(d, "Size")))
                    AddDetail("Número de Partições", WmiStr(d, "Partitions"))
                    AddDetail("Número de Série", WmiStr(d, "SerialNumber").Trim())
                    AddDetail("Firmware", WmiStr(d, "FirmwareRevision"))
                    AddDetail("Status", status, alerta)

                    ' ===== SMART detalhado (melhor esforço — nem todo driver/controlador expõe) =====
                    AddSmartDetails(WmiStr(d, "PNPDeviceID"))

                    Exit For
                End Using
            Next
        End Using

    End Sub

    ' ===== SMART detalhado via root\WMI (MSStorageDriver_FailurePredict*) =====
    ' Observação: exige que o driver de armazenamento exponha essas classes (comum em discos
    ' SATA/AHCI mais antigos; muitos NVMe e controladores RAID/RST não expõem por essa via) e,
    ' em alguns sistemas, executar o aplicativo como Administrador para conseguir ler os dados.
    Private Sub AddSmartDetails(pnpId As String)
        If String.IsNullOrEmpty(pnpId) Then
            AddDetail("SMART", "Indisponível (PNPDeviceID não encontrado)", True)
            Return
        End If

        Try
            Dim scope As New ManagementScope("root\WMI")
            scope.Connect()

            ' ----- Status geral de previsão de falha -----
            Dim encontrouStatus = False
            Using searcher As New ManagementObjectSearcher(scope, New ObjectQuery("SELECT * FROM MSStorageDriver_FailurePredictStatus")),
                  results As ManagementObjectCollection = searcher.Get()
                For Each s As ManagementObject In results
                    Using s
                        If Not SmartInstanceMatches(WmiStr(s, "InstanceName"), pnpId) Then Continue For

                        encontrouStatus = True
                        Dim predizFalha = WmiBool(s, "PredictFailure")
                        AddDetail("SMART — Previsão de Falha", If(predizFalha, "ALERTA: falha iminente", "Saudável"), predizFalha)
                    End Using
                Next
            End Using

            If Not encontrouStatus Then
                AddDetail("SMART", "Não exposto pelo driver/controlador deste disco", True)
                Return
            End If

            ' ----- Atributos SMART brutos (valores/raw por atributo) -----
            Using searcher As New ManagementObjectSearcher(scope, New ObjectQuery("SELECT * FROM MSStorageDriver_FailurePredictData")),
                  results As ManagementObjectCollection = searcher.Get()
                For Each d As ManagementObject In results
                    Using d
                        If Not SmartInstanceMatches(WmiStr(d, "InstanceName"), pnpId) Then Continue For

                        Dim vendorData = TryCast(d("VendorSpecific"), Byte())
                        If vendorData IsNot Nothing Then
                            AppendSmartAttributes(vendorData)
                        End If
                    End Using
                Next
            End Using

        Catch ex As Exception
            AddDetail("SMART", "Não disponível: " & ex.Message, True)
        End Try
    End Sub

    ' O InstanceName do driver de armazenamento normalmente contém (ou é igual a) o PNPDeviceID do disco
    Private Function SmartInstanceMatches(instanceName As String, pnpId As String) As Boolean
        If String.IsNullOrEmpty(instanceName) OrElse String.IsNullOrEmpty(pnpId) Then Return False
        Return instanceName.IndexOf(pnpId, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               pnpId.IndexOf(instanceName.TrimEnd("_"c), StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    ' Decodifica o bloco bruto de atributos SMART (formato ATA padrão: 12 bytes por atributo a partir do offset 2)
    Private Sub AppendSmartAttributes(data As Byte())
        Dim nomes As New Dictionary(Of Integer, String) From {
            {1, "Taxa de Erro de Leitura"},
            {5, "Setores Realocados"},
            {9, "Horas Ligado"},
            {12, "Ciclos de Liga/Desliga"},
            {194, "Temperatura"},
            {197, "Setores Pendentes"},
            {198, "Erros Irrecuperáveis"}
        }

        Dim offset = 2
        While offset + 12 <= data.Length
            Dim attrId = CInt(data(offset))

            If attrId <> 0 Then
                Dim valorAtual = CInt(data(offset + 3))

                Dim rawValue As Long = 0
                For i = 0 To 5
                    rawValue = rawValue Or (CLng(data(offset + 5 + i)) << (i * 8))
                Next

                Dim nome = If(nomes.ContainsKey(attrId), nomes(attrId), $"Atributo #{attrId}")
                Dim alerta = (attrId = 5 OrElse attrId = 197 OrElse attrId = 198) AndAlso rawValue > 0

                AddDetail($"SMART: {nome}", $"Valor: {valorAtual} | Raw: {rawValue}", alerta)
            End If

            offset += 12
        End While
    End Sub

    ' ===== Bateria =====
    Private Sub ShowBattery()
        currentIconKey = "info"

        Dim achouBateria = False

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_Battery"), results As ManagementObjectCollection = searcher.Get()
            For Each b As ManagementObject In results
                Using b
                    achouBateria = True

                    Dim statusCode = WmiInt(b, "BatteryStatus")
                    Dim statusTxt = BatteryStatusText(statusCode)
                    Dim cargaPercent = WmiInt(b, "EstimatedChargeRemaining")
                    Dim alertaCarga = cargaPercent > 0 AndAlso cargaPercent <= 15
                    Dim alertaStatus = statusCode = 4 OrElse statusCode = 5 ' Baixa / Crítica

                    AddDetail("Nome", WmiStr(b, "Name"))
                    AddDetail("Status", statusTxt, alertaStatus)
                    AddDetail("Carga Estimada", $"{cargaPercent}%", alertaCarga)

                    Dim tempoRestante = WmiInt(b, "EstimatedRunTime")
                    AddDetail("Tempo Restante Estimado",
                          If(tempoRestante > 0 AndAlso tempoRestante < 71582, $"{tempoRestante} minutos", "Calculando / Carregando"))

                    AddDetail("Voltagem de Design", $"{WmiInt(b, "DesignVoltage")} mV")
                End Using
                Exit For
            Next
        End Using

        If Not achouBateria Then
            AddDetail("Bateria", "Nenhuma bateria detectada neste equipamento (desktop ou não reportada pelo Windows)")
            Return
        End If

        ' ===== Saúde da bateria (capacidade de projeto vs. capacidade máxima atual) via "powercfg /batteryreport" =====
        AppendBatteryHealthFromReport()
    End Sub

    Private Function BatteryStatusText(code As Integer) As String
        Select Case code
            Case 1 : Return "Descarregando"
            Case 2 : Return "Conectada na tomada (sem bateria instalada)"
            Case 3 : Return "Totalmente carregada"
            Case 4 : Return "Baixa"
            Case 5 : Return "Crítica"
            Case 6 : Return "Carregando"
            Case 7 : Return "Carregando — carga alta"
            Case 8 : Return "Carregando — carga baixa"
            Case 9 : Return "Carregando — carga crítica"
            Case 10 : Return "Desconhecido"
            Case 11 : Return "Em espera parcial"
            Case Else : Return "Desconhecido"
        End Select
    End Function

    ' Gera o relatório nativo do Windows (powercfg /batteryreport) num arquivo temporário,
    ' extrai capacidade de projeto x capacidade máxima atual, e apaga o arquivo em seguida.
    Private Sub AppendBatteryHealthFromReport()
        Dim tempPath = Path.Combine(Path.GetTempPath(), $"battery_report_{Guid.NewGuid():N}.xml")

        Try
            Dim psi As New ProcessStartInfo("powercfg", $"/batteryreport /xml /output ""{tempPath}""") With {
                .UseShellExecute = False,
                .CreateNoWindow = True,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True
            }

            Using p = Process.Start(psi)
                p.WaitForExit(5000)
            End Using

            If Not File.Exists(tempPath) Then
                AddDetail("Saúde da Bateria", "Não foi possível gerar o relatório (powercfg)")
                Return
            End If

            Dim xml = XDocument.Load(tempPath)
            Dim ns = xml.Root.GetDefaultNamespace()

            Dim designCap = xml.Descendants(ns + "DesignCapacity").FirstOrDefault()?.Value
            Dim fullCap = xml.Descendants(ns + "FullChargeCapacity").FirstOrDefault()?.Value

            If designCap IsNot Nothing AndAlso fullCap IsNot Nothing Then
                Dim design As Double
                Dim full As Double
                If Double.TryParse(designCap, NumberStyles.Any, CultureInfo.InvariantCulture, design) AndAlso
                   Double.TryParse(fullCap, NumberStyles.Any, CultureInfo.InvariantCulture, full) AndAlso
                   design > 0 Then

                    Dim saudePercent = full / design * 100

                    AddDetail("Capacidade de Projeto", $"{design:N0} mWh")
                    AddDetail("Capacidade Máxima Atual", $"{full:N0} mWh")
                    AddDetail("Saúde da Bateria", $"{saudePercent:0.0}%", saudePercent < 80)
                End If
            End If

        Catch ex As Exception
            AddDetail("Saúde da Bateria", "Não disponível: " & ex.Message)
        Finally
            Try
                If File.Exists(tempPath) Then File.Delete(tempPath)
            Catch
                ' Se não conseguir apagar o temporário, ignora — não é crítico
            End Try
        End Try
    End Sub
    Private Sub AddDetail(propriedade As String, valor As String, Optional alerta As Boolean = False)
        Dim item As New ListViewItem(propriedade)
        item.SubItems.Add(valor)

        ' CORREÇÃO DEFINITIVA PARA O LISTVIEW: Associa a chave do ícone atual à linha
        item.ImageKey = currentIconKey

        If alerta Then
            item.ForeColor = Color.Red
            item.ImageKey = "warning"
        End If

        LVdetalhes.Items.Add(item)
    End Sub

    Private Function FormatSize(bytes As Long) As String
        Return (bytes / 1024 / 1024 / 1024).ToString("0.00") & " GB"
    End Function

    ' ===== Recursos embutidos (CSS do relatório) =====
    Private Function GetEmbeddedResourceText(resourceNameEnding As String) As String
        Dim asm = Reflection.Assembly.GetExecutingAssembly()

        Dim resourceName = asm.GetManifestResourceNames().
            FirstOrDefault(Function(n) n.EndsWith(resourceNameEnding, StringComparison.OrdinalIgnoreCase))

        If resourceName Is Nothing Then Return ""

        Using stream = asm.GetManifestResourceStream(resourceName)
            If stream Is Nothing Then Return ""
            Using reader As New StreamReader(stream)
                Return reader.ReadToEnd()
            End Using
        End Using
    End Function

    ' ===== Ícones SVG genéricos para o relatório HTML (sem depender de internet/marcas de terceiros) =====
    Private Function GetIconSvg(tag As String) As String
        Select Case tag
            Case "CPU", "VIRT_INFO"
                Return "<svg width='32' height='32' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.5' class='me-2'>" &
                       "<rect x='6' y='6' width='12' height='12' rx='1'/>" &
                       "<path d='M9 2v3M12 2v3M15 2v3M9 19v3M12 19v3M15 19v3M2 9h3M2 12h3M2 15h3M19 9h3M19 12h3M19 15h3'/>" &
                       "</svg>"

            Case "RAM", "RAM_SLOT"
                Return "<svg width='32' height='32' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.5' class='me-2'>" &
                       "<rect x='3' y='7' width='18' height='10' rx='1'/>" &
                       "<path d='M6 17v2M9 17v2M12 17v2M15 17v2M18 17v2'/>" &
                       "</svg>"

            Case "GPU", "GPU_ITEM"
                Return "<svg width='32' height='32' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.5' class='me-2'>" &
                       "<rect x='2' y='6' width='20' height='9' rx='1'/>" &
                       "<circle cx='7' cy='10.5' r='2'/>" &
                       "<path d='M2 18h6'/>" &
                       "</svg>"

            Case "MB"
                Return "<svg width='32' height='32' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.5' class='me-2'>" &
                       "<rect x='4' y='4' width='16' height='16' rx='1'/>" &
                       "<circle cx='9' cy='9' r='1.5'/><circle cx='15' cy='9' r='1.5'/>" &
                       "<path d='M4 14h16M9 4v-2M15 4v-2M9 22v-2M15 22v-2'/>" &
                       "</svg>"

            Case "REDE"
                Return "<svg width='32' height='32' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.5' class='me-2'>" &
                       "<path d='M4 9a12 12 0 0 1 16 0'/><path d='M7 12.5a7.5 7.5 0 0 1 10 0'/>" &
                       "<path d='M10.5 16a3 3 0 0 1 3 0'/><circle cx='12' cy='19' r='1'/>" &
                       "</svg>"

            Case "BIOS", "BIOS_INFO"
                Return "<svg width='32' height='32' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.5' class='me-2'>" &
                       "<rect x='5' y='5' width='14' height='14' rx='1'/>" &
                       "<rect x='9' y='9' width='6' height='6'/>" &
                       "<path d='M9 2v3M15 2v3M9 19v3M15 19v3M2 9h3M2 15h3M19 9h3M19 15h3'/>" &
                       "</svg>"

            Case "TPM", "TPM_INFO"
                Return "<svg width='32' height='32' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.5' class='me-2'>" &
                       "<rect x='5' y='11' width='14' height='9' rx='1'/>" &
                       "<path d='M8 11V7a4 4 0 0 1 8 0v4'/>" &
                       "</svg>"

            Case "STORAGE"
                Return "<svg width='32' height='32' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.5' class='me-2'>" &
                       "<ellipse cx='12' cy='6' rx='8' ry='3'/>" &
                       "<path d='M4 6v12c0 1.7 3.6 3 8 3s8-1.3 8-3V6'/>" &
                       "<path d='M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3'/>" &
                       "</svg>"

            Case "BATTERY"
                Return "<svg width='32' height='32' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.5' class='me-2'>" &
                       "<rect x='2' y='7' width='18' height='10' rx='1.5'/>" &
                       "<path d='M22 10v4'/>" &
                       "<path d='M6 9v6'/>" &
                       "</svg>"

            Case Else
                Return "<svg width='32' height='32' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.5' class='me-2'>" &
                       "<path d='M6 2h9l3 3v17H6z'/><path d='M15 2v3h3'/>" &
                       "</svg>"
        End Select
    End Function



    Private Sub ShowMotherboard()

        LVdetalhes.Items.Clear()
        currentIconKey = "mb"

        ' ===== PLACA-MÃE =====
        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard"), results As ManagementObjectCollection = searcher.Get()
            For Each b As ManagementObject In results
                Using b
                    AddDetail("Fabricante", WmiStr(b, "Manufacturer"))
                    AddDetail("Modelo", WmiStr(b, "Product"))
                    AddDetail("Número de Série", WmiStr(b, "SerialNumber"))
                    AddDetail("Asset Tag", WmiStr(b, "Tag"))
                End Using
            Next
        End Using

        ' ===== BIOS =====
        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_BIOS"), results As ManagementObjectCollection = searcher.Get()
            For Each bios As ManagementObject In results
                Using bios
                    AddDetail("BIOS Fabricante", WmiStr(bios, "Manufacturer"))
                    AddDetail("Versão BIOS", WmiStr(bios, "SMBIOSBIOSVersion"))
                    AddDetail("Data BIOS", WmiStr(bios, "ReleaseDate"))
                End Using
            Next
        End Using

    End Sub

    ' ===== Informações do BIOS/UEFI (consulta dedicada, sem misturar com placa-mãe) =====
    Private Sub ShowBiosInfo()
        LVdetalhes.Items.Clear()
        currentIconKey = "mb"

        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_BIOS"), results As ManagementObjectCollection = searcher.Get()
            For Each bios As ManagementObject In results
                Using bios
                    AddDetail("Fabricante", WmiStr(bios, "Manufacturer"))
                    AddDetail("Versão (SMBIOS)", WmiStr(bios, "SMBIOSBIOSVersion"))
                    AddDetail("Versão do BIOS", WmiStr(bios, "Version"))

                    Dim dataRaw = WmiStr(bios, "ReleaseDate")
                    Dim dataFmt = If(dataRaw.Length >= 8,
                                  $"{dataRaw.Substring(0, 4)}-{dataRaw.Substring(4, 2)}-{dataRaw.Substring(6, 2)}",
                                  "Desconhecida")
                    AddDetail("Data de Lançamento", dataFmt)
                    AddDetail("Número de Série", WmiStr(bios, "SerialNumber"))
                End Using
                Exit For
            Next
        End Using

        AddDetail("Modo de Boot", If(IsUEFI(), "UEFI", "Legacy"))
    End Sub

    ' ===== Verificação dedicada de TPM (presença, versão, ativação) =====
    Private Sub ShowTpmInfo()
        LVdetalhes.Items.Clear()
        currentIconKey = "info"

        Try
            Dim scope As New ManagementScope("root\CIMV2\Security\MicrosoftTpm")
            scope.Connect()
            Dim query As New ObjectQuery("SELECT * FROM Win32_Tpm")

            Dim encontrado = False
            Using searcher As New ManagementObjectSearcher(scope, query), results As ManagementObjectCollection = searcher.Get()
                For Each tpm As ManagementObject In results
                    Using tpm
                        encontrado = True
                        AddDetail("TPM Presente", "Sim")
                        AddDetail("Versão da Especificação", WmiStr(tpm, "SpecVersion", "Desconhecida"))
                        AddDetail("Fabricante", WmiStr(tpm, "ManufacturerIdTxt", "Desconhecido"))
                        AddDetail("Versão do Fabricante", WmiStr(tpm, "ManufacturerVersion", "Desconhecida"))

                        Dim ativado = False
                        Try
                            If tpm("IsActivated_InitialValue") IsNot Nothing Then
                                ativado = CBool(tpm("IsActivated_InitialValue"))
                            End If
                        Catch
                        End Try
                        AddDetail("Ativado", If(ativado, "Sim", "Não"), Not ativado)

                        Dim habilitado = False
                        Try
                            If tpm("IsEnabled_InitialValue") IsNot Nothing Then
                                habilitado = CBool(tpm("IsEnabled_InitialValue"))
                            End If
                        Catch
                        End Try
                        AddDetail("Habilitado", If(habilitado, "Sim", "Não"), Not habilitado)
                    End Using
                    Exit For
                Next
            End Using

            If Not encontrado Then
                AddDetail("TPM Presente", "Não", True)
            End If

        Catch
            AddDetail("TPM Presente", "Não / Não suportado neste hardware", True)
        End Try
    End Sub

    ' ===== Verificação dedicada de virtualização (VT-x/AMD-V, SLAT, Hyper-V) =====
    Private Sub ShowVirtualizationInfo()
        LVdetalhes.Items.Clear()
        currentIconKey = "cpu"

        Dim vt As String = "Não suportado"
        Dim slat As String = "Não suportado"
        Dim hyperv As String = "Nenhum hipervisor ativo"

        Using searcher As New ManagementObjectSearcher(
        "SELECT VirtualizationFirmwareEnabled, SecondLevelAddressTranslationExtensions, VMMonitorModeExtensions FROM Win32_Processor"), results As ManagementObjectCollection = searcher.Get()
            For Each cpu As ManagementObject In results
                Using cpu
                    Dim suportaVmx = WmiBool(cpu, "VMMonitorModeExtensions")
                    Dim ativoNoBios = WmiBool(cpu, "VirtualizationFirmwareEnabled")

                    vt = If(ativoNoBios, "Ativo no BIOS",
                         If(suportaVmx, "Suportado pelo CPU, mas desativado no BIOS", "Não suportado"))
                    slat = If(WmiBool(cpu, "SecondLevelAddressTranslationExtensions"), "Suportado", "Não suportado")
                End Using
                Exit For
            Next
        End Using

        Using searcher As New ManagementObjectSearcher("SELECT HypervisorPresent FROM Win32_ComputerSystem"), results As ManagementObjectCollection = searcher.Get()
            For Each cs As ManagementObject In results
                Using cs
                    hyperv = If(WmiBool(cs, "HypervisorPresent"), "Um hipervisor está ativo (Hyper-V ou outro)", "Nenhum hipervisor ativo")
                End Using
                Exit For
            Next
        End Using

        AddDetail("Virtualização (VT-x / AMD-V)", vt, Not vt.Contains("Ativo"))
        AddDetail("SLAT (EPT/RVI)", slat, slat = "Não suportado")
        AddDetail("Hyper-V / Hipervisor", hyperv)
    End Sub

    ' ===== Resumo do nó raiz "Dispositivos" =====
    Private Sub ShowDispositivosResumo()
        LVdetalhes.Items.Clear()
        currentIconKey = "info"

        ' CPU
        Using searcher As New ManagementObjectSearcher("SELECT Name, NumberOfCores, NumberOfLogicalProcessors FROM Win32_Processor"), results As ManagementObjectCollection = searcher.Get()
            For Each cpu As ManagementObject In results
                Using cpu
                    AddDetail("Processador", $"{WmiStr(cpu, "Name")} ({WmiInt(cpu, "NumberOfCores")}C/{WmiInt(cpu, "NumberOfLogicalProcessors")}T)")
                End Using
                Exit For
            Next
        End Using

        ' RAM
        Using searcher As New ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem"), results As ManagementObjectCollection = searcher.Get()
            For Each cs As ManagementObject In results
                Using cs
                    AddDetail("Memória RAM", FormatSize(WmiLng(cs, "TotalPhysicalMemory")))
                End Using
                Exit For
            Next
        End Using

        ' Placa-mãe
        Using searcher As New ManagementObjectSearcher("SELECT Manufacturer, Product FROM Win32_BaseBoard"), results As ManagementObjectCollection = searcher.Get()
            For Each b As ManagementObject In results
                Using b
                    AddDetail("Placa-mãe", $"{WmiStr(b, "Manufacturer")} {WmiStr(b, "Product")}".Trim())
                End Using
                Exit For
            Next
        End Using

        ' GPU(s)
        Dim gpus As New List(Of String)
        Using searcher As New ManagementObjectSearcher("SELECT Name FROM Win32_VideoController"), results As ManagementObjectCollection = searcher.Get()
            For Each gpu As ManagementObject In results
                Using gpu
                    Dim nome = WmiStr(gpu, "Name")
                    If nome <> "" Then gpus.Add(nome)
                End Using
            Next
        End Using
        AddDetail("Placa(s) de Vídeo", If(gpus.Count > 0, String.Join(" | ", gpus), "Não detectada"))

        ' Adaptadores de rede (aplicando o mesmo filtro da árvore)
        Dim totalNic As Integer = 0
        Dim ativosNic As Integer = 0
        Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapter"), results As ManagementObjectCollection = searcher.Get()
            For Each nic As ManagementObject In results
                Using nic
                    If Not IsWantedNetworkAdapter(nic) Then Continue For
                    totalNic += 1
                    If WmiBool(nic, "NetEnabled") Then ativosNic += 1
                End Using
            Next
        End Using
        AddDetail("Adaptadores de Rede", $"{totalNic} detectado(s) — {ativosNic} ativo(s)")

        ' Dispositivos PnP (contagem geral, reaproveita a mesma lógica da árvore)
        Dim totalPnp As Integer = 0
        Using searcher As New ManagementObjectSearcher(
        "SELECT * FROM Win32_PnPEntity WHERE Present = TRUE AND PNPDeviceID IS NOT NULL"), results As ManagementObjectCollection = searcher.Get()
            For Each dev As ManagementObject In results
                Using dev
                    Dim pnpId = WmiStr(dev, "PNPDeviceID")
                    If Not IsRealPnPDevice(pnpId) Then Continue For
                    If String.IsNullOrEmpty(WmiStr(dev, "Name")) Then Continue For
                    totalPnp += 1
                End Using
            Next
        End Using
        AddDetail("Dispositivos Plug and Play", $"{totalPnp} dispositivo(s) detectado(s)")
    End Sub

    ' ===== Resumo do nó raiz "Dispositivos Plug and Play" =====
    Private Sub ShowPnpResumo()
        LVdetalhes.Items.Clear()
        currentIconKey = "pnp"

        Dim porBarramento As New Dictionary(Of String, Integer)
        Dim totalOk As Integer = 0
        Dim totalSemDriver As Integer = 0
        Dim totalErro As Integer = 0

        Using searcher As New ManagementObjectSearcher(
        "SELECT * FROM Win32_PnPEntity WHERE Present = TRUE AND PNPDeviceID IS NOT NULL"), results As ManagementObjectCollection = searcher.Get()
            For Each dev As ManagementObject In results
                Using dev
                    Dim pnpId = WmiStr(dev, "PNPDeviceID")
                    If Not IsRealPnPDevice(pnpId) Then Continue For
                    If String.IsNullOrEmpty(WmiStr(dev, "Name")) Then Continue For

                    Dim barramento = GetBusFromPNP(pnpId)
                    If Not porBarramento.ContainsKey(barramento) Then porBarramento(barramento) = 0
                    porBarramento(barramento) += 1

                    Select Case GetDeviceState(dev)
                        Case "NO_DRIVER" : totalSemDriver += 1
                        Case "ERROR" : totalErro += 1
                        Case Else : totalOk += 1
                    End Select
                End Using
            Next
        End Using

        Dim totalGeral = porBarramento.Values.Sum()
        AddDetail("Total de Dispositivos", totalGeral.ToString())
        AddDetail("Funcionando OK", totalOk.ToString())
        AddDetail("Sem Driver", totalSemDriver.ToString(), totalSemDriver > 0)
        AddDetail("Com Erro", totalErro.ToString(), totalErro > 0)

        For Each kv In porBarramento.OrderByDescending(Function(x) x.Value)
            AddDetail($"Barramento: {kv.Key}", $"{kv.Value} dispositivo(s)")
        Next
    End Sub

    ' ===== Resumo de um barramento específico (ex: USB, PCI/PCIe...) =====
    Private Sub ShowPnpBusResumo(barramento As String)
        LVdetalhes.Items.Clear()
        currentIconKey = "bus"

        Using searcher As New ManagementObjectSearcher(
        "SELECT * FROM Win32_PnPEntity WHERE Present = TRUE AND PNPDeviceID IS NOT NULL"), results As ManagementObjectCollection = searcher.Get()
            For Each dev As ManagementObject In results
                Using dev
                    Dim pnpId = WmiStr(dev, "PNPDeviceID")
                    If Not IsRealPnPDevice(pnpId) Then Continue For

                    Dim nome = WmiStr(dev, "Name")
                    If String.IsNullOrEmpty(nome) Then Continue For

                    If GetBusFromPNP(pnpId) <> barramento Then Continue For

                    Dim estado = GetDeviceState(dev)
                    Dim alerta = estado <> "OK"
                    AddDetail(nome, If(estado = "OK", "Funcionando", If(estado = "NO_DRIVER", "Sem driver", "Erro")), alerta)
                End Using
            Next
        End Using
    End Sub

    Public Sub ExportReport(completo As Boolean, Optional imprimir As Boolean = False, Optional customPath As String = Nothing)
        ' Detecta se o aplicativo foi iniciado com o parâmetro de relatório automático
        Dim isAuto As Boolean = My.Application.CommandLineArgs.Any(
            Function(a) a.Equals("-report", StringComparison.OrdinalIgnoreCase) OrElse
                        a.Equals("--report", StringComparison.OrdinalIgnoreCase))
        Dim filePath As String = ""

        ' --- 1. TRATAMENTO DE AVISOS ---
        If imprimir AndAlso completo AndAlso Not isAuto Then
            Dim respAviso = MessageBox.Show(
            "O relatório COMPLETO pode consumir várias folhas de papel." & vbCrLf & vbCrLf &
            "Sim = Imprimir completo | Não = Apenas resumo",
            "Aviso de Impressão", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)

            If respAviso = DialogResult.Cancel Then Exit Sub
            If respAviso = DialogResult.No Then completo = False
        End If

        ' --- 2. DEFINIÇÃO DO CAMINHO DO ARQUIVO ---
        If customPath IsNot Nothing Then
            ' FIX: suporte a --out <arquivo> no modo CLI
            filePath = customPath
        ElseIf isAuto Then
            ' Salva direto em Documentos sem perguntar ao usuário
            filePath = Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "Relatorio_Auto.html")
        ElseIf imprimir Then
            filePath = Path.Combine(Path.GetTempPath(), "RelatorioSistema.html")
        Else
            ' Abre o diálogo de salvar apenas se não for automático
            Dim sfd As New SaveFileDialog()
            sfd.Filter = "Arquivo HTML|*.html"
            sfd.FileName = "Relatorio.html"

            If sfd.ShowDialog() <> DialogResult.OK Then Exit Sub
            filePath = sfd.FileName
        End If

        ' --- 3. GERAÇÃO DO CONTEÚDO HTML ---
        Dim sb As New StringBuilder()
        sb.AppendLine("<!DOCTYPE html><html lang=""pt-br""><head>")
        sb.AppendLine("<meta charset=""UTF-8"">")
        sb.AppendLine("<style>")
        sb.AppendLine(GetEmbeddedResourceText("bootstrap.min.css"))
        sb.AppendLine("</style>")
        sb.AppendLine("<title>Relatório de Sistema</title></head><body class=""p-3"">")
        sb.AppendLine("<div class=""container"">")

        sb.AppendLine(BuildHtmlMetadata())
        sb.AppendLine("<h2 class=""mb-3"">Relatório de Sistema</h2>")

        If completo Then
            For Each node As TreeNode In TVdispositivos.Nodes
                AppendNodeHtmlFull(node, sb)
            Next
        Else
            If TVdispositivos.SelectedNode IsNot Nothing Then
                AppendNodeHtmlFull(TVdispositivos.SelectedNode, sb)
            End If
        End If

        sb.AppendLine("</div></body></html>")

        ' --- 4. SALVAMENTO E FINALIZAÇÃO ---
        Try
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8)

            ' Só interage com o usuário se NÃO for modo automático
            If Not isAuto Then
                If imprimir Then
                    Process.Start(New ProcessStartInfo With {.FileName = filePath, .UseShellExecute = True})
                Else
                    Dim resp = MessageBox.Show($"Relatório salvo em:{vbCrLf}{filePath}{vbCrLf}{vbCrLf}Deseja abrir agora?",
                                         "Exportação concluída", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    If resp = DialogResult.Yes Then
                        Process.Start(New ProcessStartInfo With {.FileName = filePath, .UseShellExecute = True})
                    End If
                End If
            Else
                ' FIX: feedback no console quando rodando via CLI (não faz nada se não houver console anexado)
                Console.WriteLine($"Relatório salvo em: {filePath}")
            End If
        Catch ex As Exception
            If Not isAuto Then
                MessageBox.Show("Erro ao salvar: " & ex.Message)
            Else
                Console.Error.WriteLine("Erro ao salvar: " & ex.Message)
            End If
        End Try
    End Sub


    Private Function BuildHtmlMetadata() As String
        Dim sb As New StringBuilder()

        sb.AppendLine("<div class='card mb-4'>")
        sb.AppendLine("<div class='card-header bg-dark text-white'>")
        sb.AppendLine("<strong>Relatório Técnico do Sistema</strong>")
        sb.AppendLine("</div>")
        sb.AppendLine("<div class='card-body'>")

        sb.AppendLine("<ul class='list-unstyled mb-0'>")
        ' FIX: HtmlEncode em todo valor que vai para o HTML (evita quebrar o layout com < > & em nomes/valores)
        sb.AppendLine($"<li><strong>Computador:</strong> {WebUtility.HtmlEncode(Environment.MachineName)}</li>")
        sb.AppendLine($"<li><strong>Usuário:</strong> {WebUtility.HtmlEncode(Environment.UserName)}</li>")
        sb.AppendLine($"<li><strong>Sistema Operacional:</strong> {WebUtility.HtmlEncode(My.Computer.Info.OSFullName)}</li>")
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
                iconHtml = GetIconSvg("CPU")
            Case "RAM", "RAM_SLOT"
                iconHtml = GetIconSvg("RAM")
            Case "GPU", "GPU_ITEM"
                iconHtml = GetIconSvg("GPU")
            Case "MB"
                iconHtml = GetIconSvg("MB")
            Case "REDE"
                iconHtml = GetIconSvg("REDE")
            Case "BIOS", "BIOS_INFO"
                iconHtml = GetIconSvg("BIOS")
            Case "TPM", "TPM_INFO"
                iconHtml = GetIconSvg("TPM")
            Case "VIRT_INFO"
                iconHtml = GetIconSvg("VIRT_INFO")
            Case "STORAGE"
                iconHtml = GetIconSvg("STORAGE")
            Case "BATTERY"
                iconHtml = GetIconSvg("BATTERY")
            Case Else
                iconHtml = GetIconSvg("")
        End Select

        sb.AppendLine("<div class=""mb-3"">")
        ' FIX: HtmlEncode no texto do nó (nomes de dispositivo podem ter caracteres especiais)
        sb.AppendLine("<h4>" & iconHtml & WebUtility.HtmlEncode(node.Text) & "</h4>")

        ' ===== ListView detalhes se selecionado =====
        If LVdetalhes.Items.Count > 0 AndAlso node.IsSelected Then
            sb.AppendLine("<table class=""table table-bordered table-striped"">")
            sb.AppendLine("<thead class=""table-dark""><tr><th>Propriedade</th><th>Valor</th></tr></thead>")
            sb.AppendLine("<tbody>")
            For Each item As ListViewItem In LVdetalhes.Items
                ' FIX: HtmlEncode nos valores da tabela (evita quebrar o HTML com <, >, & vindos do WMI/registro)
                sb.AppendLine($"<tr><td>{WebUtility.HtmlEncode(item.Text)}</td><td>{WebUtility.HtmlEncode(item.SubItems(1).Text)}</td></tr>")
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
        ExportReport(True, False)

    End Sub

    Private Sub SomenteTelaSelecionadaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SomenteTelaSelecionadaToolStripMenuItem.Click
        ExportReport(False, False)

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

    Private Sub RelatórioCompletoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RelatórioCompletoToolStripMenuItem1.Click
        ExportReport(True, True)
    End Sub

    Private Sub SomentePáginaSelecionadaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SomentePáginaSelecionadaToolStripMenuItem.Click
        ExportReport(False, True)
    End Sub

    Private Sub CompletoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CompletoToolStripMenuItem.Click
        ExportReport(True, False)
    End Sub

    Private Sub TelaSelecionadaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TelaSelecionadaToolStripMenuItem.Click
        ExportReport(False, False)
    End Sub

    Private Sub SairToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SairToolStripMenuItem.Click
        Application.Exit()

    End Sub
    ' Método para ser chamado via linha de comando
    Public Sub GerarRelatorioAutomatico()
        ' Forçamos a construção da árvore internamente para coletar os dados
        BuildTree()
        ExportReport(True, False)
    End Sub




    Private Sub CarregarMenuFerramentas()
        ' 1. REMOVE APENAS OS ITENS DINÂMICOS ANTIGOS
        For i As Integer = ToolStripMenuItem1.DropDownItems.Count - 1 To 0 Step -1
            Dim item As ToolStripItem = ToolStripMenuItem1.DropDownItems(i)

            If item.Name.StartsWith("dyn_") Then
                ToolStripMenuItem1.DropDownItems.RemoveAt(i)
            End If
        Next

        Dim pastaTools As String = Path.Combine(Application.StartupPath, "tools")

        If Directory.Exists(pastaTools) Then
            Dim executaveis() As String = Directory.GetFiles(pastaTools, "*.exe")

            If executaveis.Length > 0 Then
                Dim separador As New ToolStripSeparator()
                separador.Name = "dyn_separador"
                ToolStripMenuItem1.DropDownItems.Add(separador)
            End If

            ' 2. ADICIONA OS NOVOS ITENS COM ÍCONE
            For Each caminhoArquivo As String In executaveis
                Dim nomeItem As String = Path.GetFileNameWithoutExtension(caminhoArquivo)
                Dim novoItem As New ToolStripMenuItem(nomeItem)

                novoItem.Name = "dyn_" & nomeItem
                novoItem.Tag = caminhoArquivo

                ' --- EXTRAÇÃO DO ÍCONE ---
                Try
                    ' Extrai o ícone associado ao arquivo executável
                    Using iconeApp As Icon = Icon.ExtractAssociatedIcon(caminhoArquivo)
                        If iconeApp IsNot Nothing Then
                            ' Converte o ícone para Bitmap e atribui ao menu
                            novoItem.Image = iconeApp.ToBitmap()
                        End If
                    End Using
                Catch
                    ' Se falhar ao extrair o ícone de algum executável específico, 
                    ' o item apenas fica sem imagem, mas o app não quebra.
                End Try
                ' -------------------------

                AddHandler novoItem.Click, AddressOf ItemMenu_Click
                ToolStripMenuItem1.DropDownItems.Add(novoItem)
            Next
        End If
    End Sub

    ' Este é o método único que vai gerenciar o clique de QUALQUER executável adicionado
    Private Sub ItemMenu_Click(sender As Object, e As EventArgs)
        ' Descobre qual item foi clicado
        Dim itemClicado As ToolStripMenuItem = CType(sender, ToolStripMenuItem)

        ' Recupera o caminho do arquivo que salvamos na Tag
        Dim caminhoExecutavel As String = itemClicado.Tag.ToString()

        Try
            ' Executa o arquivo
            Process.Start(caminhoExecutavel)
        Catch ex As Exception
            MessageBox.Show("Não foi possível iniciar a ferramenta: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub ReplicarNoToolStrip()
        ' 1. REMOVE APENAS OS BOTÕES DINÂMICOS ANTIGOS DO TOOLSTRIP
        For i As Integer = ToolStrip1.Items.Count - 1 To 0 Step -1
            Dim item As ToolStripItem = ToolStrip1.Items(i)

            If item.Name IsNot Nothing AndAlso item.Name.StartsWith("dyn_") Then
                ToolStrip1.Items.RemoveAt(i)
            End If
        Next

        ' Opcional: Adiciona um separador se houver itens a replicar
        If ToolStripMenuItem1.DropDownItems.Count > 0 Then
            Dim sep As New ToolStripSeparator()
            sep.Name = "dyn_separador_ts"
            ToolStrip1.Items.Add(sep)
        End If

        ' 2. VARRE O MENU E CRIA OS BOTÕES APENAS COM ÍCONE
        For Each itemMenu As ToolStripItem In ToolStripMenuItem1.DropDownItems

            ' Ignora separadores
            If TypeOf itemMenu Is ToolStripSeparator Then Continue For

            Dim novoBotao As New ToolStripButton()

            ' Configura para exibir APENAS o ícone
            novoBotao.DisplayStyle = ToolStripItemDisplayStyle.Image
            novoBotao.Image = itemMenu.Image

            ' Coloca o texto apenas como ToolTip (aquela legendinha ao passar o mouse por cima)
            novoBotao.ToolTipText = itemMenu.Text
            novoBotao.Text = itemMenu.Text

            ' TRUQUE MÁGICO: Guarda o objeto do menu inteiro na Tag do botão
            novoBotao.Tag = itemMenu

            ' Marca o nome para a limpeza posterior
            novoBotao.Name = "dyn_btn_" & itemMenu.Name

            ' Vincula ao novo evento de clique unificado
            AddHandler novoBotao.Click, AddressOf BotaoToolStrip_Click

            ToolStrip1.Items.Add(novoBotao)
        Next
    End Sub

    ' Evento de clique que espelha o comportamento do menu original
    Private Sub BotaoToolStrip_Click(sender As Object, e As EventArgs)
        Dim botaoClicado As ToolStripButton = CType(sender, ToolStripButton)

        ' Recupera o item de menu original associado a este botão
        If botaoClicado.Tag IsNot Nothing AndAlso TypeOf botaoClicado.Tag Is ToolStripItem Then
            Dim itemMenuOriginal As ToolStripItem = CType(botaoClicado.Tag, ToolStripItem)

            ' Dispara o evento de clique do menu original exatamente como se o usuário tivesse ido lá e clicado nele
            itemMenuOriginal.PerformClick()
        End If
    End Sub

    Private Sub BarraDeFerramentasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BarraDeFerramentasToolStripMenuItem.Click
        ToolStrip1.Visible = BarraDeFerramentasToolStripMenuItem.Checked

    End Sub

    Private Sub ListaLateraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListaLateraToolStripMenuItem.Click
        SplitContainer1.Panel1.Visible = BarraDeFerramentasToolStripMenuItem.Checked

    End Sub
End Class