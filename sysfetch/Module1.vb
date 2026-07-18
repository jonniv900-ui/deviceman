Imports System.Management
Imports System.Diagnostics
Imports Microsoft.Win32
Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports System.IO
Imports System.Runtime.InteropServices

' =====================================================================
' NATIVE INTEROP - Buffer de tela via WriteConsoleOutput (Win32)
' Permite montar o frame inteiro em memória e jogar na tela com UMA
' única chamada, eliminando o flicker causado por múltiplos
' Console.Write / SetCursorPosition e Console.Clear() por frame.
' =====================================================================

<StructLayout(LayoutKind.Explicit)>
Public Structure CHAR_INFO
    <FieldOffset(0)> Public UnicodeChar As UInt16
    <FieldOffset(2)> Public Attributes As UInt16
End Structure

<StructLayout(LayoutKind.Sequential)>
Public Structure COORD
    Public X As Short
    Public Y As Short
    Public Sub New(x As Short, y As Short)
        Me.X = x
        Me.Y = y
    End Sub
End Structure

<StructLayout(LayoutKind.Sequential)>
Public Structure SMALL_RECT
    Public Left As Short
    Public Top As Short
    Public Right As Short
    Public Bottom As Short
End Structure

Module NativeMethods
    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Unicode, EntryPoint:="CreateFileW")>
    Public Function CreateFileW(lpFileName As String, dwDesiredAccess As UInteger, dwShareMode As UInteger,
                                 lpSecurityAttributes As IntPtr, dwCreationDisposition As UInteger,
                                 dwFlagsAndAttributes As UInteger, hTemplateFile As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True, EntryPoint:="WriteConsoleOutputW")>
    Public Function WriteConsoleOutputW(hConsoleOutput As IntPtr, lpBuffer As CHAR_INFO(), dwBufferSize As COORD,
                                         dwBufferCoord As COORD, ByRef lpWriteRegion As SMALL_RECT) As Boolean
    End Function
End Module

''' <summary>
''' Buffer de tela em memória. Todo o desenho acontece aqui (PutChar / WriteText);
''' só quando Render() é chamado é que o conteúdo vai para o console, em uma
''' única chamada de API nativa - sem flicker, sem múltiplas idas ao terminal.
''' </summary>
Public Class ScreenBuffer
    Private Const GENERIC_READ As UInteger = &H80000000UI
    Private Const GENERIC_WRITE As UInteger = &H40000000UI
    Private Const FILE_SHARE_READ As UInteger = &H1UI
    Private Const FILE_SHARE_WRITE As UInteger = &H2UI
    Private Const OPEN_EXISTING As UInteger = 3UI
    Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80UI

    Private ReadOnly _width As Integer
    Private ReadOnly _height As Integer
    Private ReadOnly _handle As IntPtr
    Private ReadOnly _cells As CHAR_INFO()

    Public Sub New(width As Integer, height As Integer)
        _width = width
        _height = height
        _cells = New CHAR_INFO(_width * _height - 1) {}
        _handle = CreateFileW("CONOUT$", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE,
                               IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero)
        Clear(ConsoleColor.DarkBlue)
    End Sub

    Public ReadOnly Property Width As Integer
        Get
            Return _width
        End Get
    End Property

    Public ReadOnly Property Height As Integer
        Get
            Return _height
        End Get
    End Property

    Private Function Idx(x As Integer, y As Integer) As Integer
        Return y * _width + x
    End Function

    Private Function MakeAttr(fg As ConsoleColor, bg As ConsoleColor) As UInt16
        Return CUShort(CInt(fg) Or (CInt(bg) << 4))
    End Function

    Public Sub Clear(Optional bg As ConsoleColor = ConsoleColor.DarkBlue)
        Dim attr = MakeAttr(ConsoleColor.White, bg)
        For i = 0 To _cells.Length - 1
            _cells(i).UnicodeChar = CUShort(AscW(" "c))
            _cells(i).Attributes = attr
        Next
    End Sub

    Public Sub PutChar(x As Integer, y As Integer, ch As Char, fg As ConsoleColor, bg As ConsoleColor)
        If x < 0 OrElse x >= _width OrElse y < 0 OrElse y >= _height Then Exit Sub
        Dim i = Idx(x, y)
        _cells(i).UnicodeChar = CUShort(AscW(ch))
        _cells(i).Attributes = MakeAttr(fg, bg)
    End Sub

    Public Sub WriteText(x As Integer, y As Integer, text As String, fg As ConsoleColor, Optional bg As ConsoleColor = ConsoleColor.DarkBlue)
        For i = 0 To text.Length - 1
            PutChar(x + i, y, text(i), fg, bg)
        Next
    End Sub

    ''' <summary>Joga o buffer inteiro na tela em uma única chamada nativa (sem flicker).</summary>
    Public Sub Render()
        Dim size As New COORD(CShort(_width), CShort(_height))
        Dim origin As New COORD(0, 0)
        Dim region As New SMALL_RECT With {.Left = 0, .Top = 0, .Right = CShort(_width - 1), .Bottom = CShort(_height - 1)}
        WriteConsoleOutputW(_handle, _cells, size, origin, region)
    End Sub
End Class

' =====================================================================
' APLICAÇÃO
' =====================================================================
Module Module1
    Dim cpuTotal As PerformanceCounter
    Dim coreCounters As New List(Of PerformanceCounter)

    ' --- Cache de informações estáticas (coletadas 1x, nunca mudam durante a execução) ---
    Dim cachedOS As String = "N/A"
    Dim cachedKernel As String = "N/A"
    Dim cachedCPU As String = "N/A"
    Dim cachedMB As String = "N/A"
    Dim cachedGPU As String = "N/A"
    Dim cachedVRAM As String = "N/A"
    Dim cachedRamSlots As New List(Of String)
    Dim cachedIsWinPE As Boolean = False

    Sub Main()
        Try
            Console.Title = "SYSTEM INSPECTOR - v4.0"
            Console.SetWindowSize(100, 32)
            Console.SetBufferSize(100, 32)
            Console.CursorVisible = False
        Catch : End Try

        ShowSplash()
        InitCounters()
        CacheStaticInfo() ' WMI custoso (CPU/GPU/VRAM/RAM) é lido 1x só, fora do loop

        Dim screen As New ScreenBuffer(100, 32)

        While True
            screen.Clear(ConsoleColor.DarkBlue)
            DrawBox(screen, 0, 0, 99, 28, $" {Environment.MachineName} - System Diagnostic [{If(cachedIsWinPE, "WinPE", "Windows")}] ")
            DrawWindowsLogo(screen, 4, 2)
            DrawStaticLabels(screen)
            DrawMenuFooter(screen, "  D Dispositivos | M Monitor Full | F Ferramentas | ESC Sair ")

            While Not Console.KeyAvailable
                UpdateMainDynamicValues(screen)
                screen.Render()
                Threading.Thread.Sleep(400)
            End While

            Dim key = Console.ReadKey(True).Key
            Select Case key
                Case ConsoleKey.D : ShowPnPDevices(screen)
                Case ConsoleKey.M : ShowFullMonitor(screen)
                Case ConsoleKey.F : ShowTools(screen)
                Case ConsoleKey.Escape : Exit While
            End Select
        End While
    End Sub

    ''' <summary>Coleta uma única vez os dados de hardware que não mudam durante a execução.
    ''' Antes, GPU/VRAM eram consultados via WMI a cada 500ms - isso sozinho já causava
    ''' engasgos perceptíveis, já que cada query WMI pode levar dezenas de ms.</summary>
    Sub CacheStaticInfo()
        cachedIsWinPE = IsWinPE()
        cachedOS = Truncate(My.Computer.Info.OSFullName, 50)
        cachedKernel = Environment.OSVersion.VersionString
        cachedCPU = Truncate(GetWMI("Win32_Processor", "Name"), 60)
        cachedMB = GetWMI("Win32_BaseBoard", "Product") & " | BIOS: " & GetWMI("Win32_BIOS", "Version")
        cachedGPU = Truncate(GetWMI("Win32_VideoController", "Name"), 50)
        cachedVRAM = GetVRAM()

        cachedRamSlots.Clear()
        Try
            Using s As New ManagementObjectSearcher("SELECT Capacity, Speed, DeviceLocator FROM Win32_PhysicalMemory")
                For Each obj In s.Get()
                    Dim cap = (Val(obj("Capacity")) / 1024 ^ 3).ToString("F0")
                    cachedRamSlots.Add($"• {obj("DeviceLocator")}: {cap}GB @ {obj("Speed")}MHz")
                Next
            End Using
        Catch : End Try
    End Sub

    Sub ShowTools(sb As ScreenBuffer)
        Dim toolsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools")

        If Not Directory.Exists(toolsPath) Then
            Try : Directory.CreateDirectory(toolsPath) : Catch : End Try
        End If

        Dim files = Directory.GetFiles(toolsPath, "*.*").Where(Function(f) _
            f.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) Or
            f.EndsWith(".bat", StringComparison.OrdinalIgnoreCase) Or
            f.EndsWith(".ps1", StringComparison.OrdinalIgnoreCase)).ToList()

        ' Descrições calculadas 1x (antes eram recalculadas a cada redesenho da lista)
        Dim descriptions As New List(Of String)
        For Each f In files
            Dim d = ""
            Try
                If f.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) Then
                    Dim info = FileVersionInfo.GetVersionInfo(f)
                    d = If(Not String.IsNullOrEmpty(info.FileDescription), $" ({info.FileDescription})", "")
                ElseIf f.EndsWith(".bat", StringComparison.OrdinalIgnoreCase) Then
                    d = " [Batch Script]"
                ElseIf f.EndsWith(".ps1", StringComparison.OrdinalIgnoreCase) Then
                    d = " [PowerShell]"
                End If
            Catch : End Try
            descriptions.Add(d)
        Next

        Dim idx = 0
        Dim startIdx = 0

        While True
            sb.Clear(ConsoleColor.DarkBlue)

            Dim titulo = $" CAIXA DE FERRAMENTAS ({files.Count} itens) "
            DrawBox(sb, 5, 4, 90, 22, titulo)

            If files.Count = 0 Then
                sb.WriteText(10, 12, "Nenhuma ferramenta encontrada na pasta \tools.", ConsoleColor.Red, ConsoleColor.DarkBlue)
            Else
                For i = 0 To 14
                    Dim fileIdx = startIdx + i
                    If fileIdx < files.Count Then
                        Dim fileName = Path.GetFileName(files(fileIdx))
                        Dim line = Truncate(fileName & descriptions(fileIdx), 80).PadRight(83)
                        If fileIdx = idx Then
                            sb.WriteText(8, 6 + i, "> " & line, ConsoleColor.Black, ConsoleColor.Cyan)
                        Else
                            sb.WriteText(8, 6 + i, "  " & line, ConsoleColor.White, ConsoleColor.DarkBlue)
                        End If
                    End If
                Next
            End If

            DrawMenuFooter(sb, " [SETAS] Navegar | [ENTER] Lançar | [R] Reiniciar | [S] Desligar | [ESC] Voltar ")
            sb.Render()

            Dim k = Console.ReadKey(True)
            If k.Key = ConsoleKey.Escape Then Exit While

            If k.Key = ConsoleKey.R Then
                If ConfirmAction(sb, "REINICIAR O SISTEMA?") Then RebootMachine()
            End If
            If k.Key = ConsoleKey.S Then
                If ConfirmAction(sb, "DESLIGAR O COMPUTADOR?") Then ShutdownMachine()
            End If

            If k.Key = ConsoleKey.DownArrow And idx < files.Count - 1 Then
                idx += 1
                If idx >= startIdx + 15 Then startIdx += 1
            End If
            If k.Key = ConsoleKey.UpArrow And idx > 0 Then
                idx -= 1
                If idx < startIdx Then startIdx -= 1
            End If

            If k.Key = ConsoleKey.Enter AndAlso files.Count > 0 Then
                Try
                    Process.Start(New ProcessStartInfo(files(idx)) With {.UseShellExecute = True})
                Catch ex As Exception
                    DrawBoxColor(sb, 15, 10, 70, 7, " FALHA NO LANÇAMENTO ", ConsoleColor.DarkRed, ConsoleColor.White)
                    sb.WriteText(17, 12, "Erro: " & Truncate(ex.Message, 60), ConsoleColor.White, ConsoleColor.DarkRed)
                    sb.Render()
                    Console.ReadKey(True)
                End Try
            End If
        End While
    End Sub

    ''' <summary>Detecta se o processo está rodando dentro de um ambiente WinPE.
    ''' O instalador do WinPE cria a chave HKLM\SYSTEM\CurrentControlSet\Control\MiniNT,
    ''' que não existe em uma instalação normal do Windows - é o jeito padrão de checar isso.</summary>
    Function IsWinPE() As Boolean
        Try
            Using key = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Control\MiniNT")
                Return key IsNot Nothing
            End Using
        Catch
            Return False
        End Try
    End Function

    ''' <summary>Reinicia a máquina usando o comando correto para o ambiente detectado:
    ''' wpeutil no WinPE (shutdown.exe geralmente não funciona lá) ou shutdown.exe no Windows normal.</summary>
    Sub RebootMachine()
        Try
            If IsWinPE() Then
                Process.Start("wpeutil", "reboot")
            Else
                Process.Start("shutdown", "/r /t 0")
            End If
        Catch : End Try
    End Sub

    ''' <summary>Desliga a máquina usando o comando correto para o ambiente detectado.</summary>
    Sub ShutdownMachine()
        Try
            If IsWinPE() Then
                Process.Start("wpeutil", "shutdown")
            Else
                Process.Start("shutdown", "/s /t 0")
            End If
        Catch : End Try
    End Sub
    Function ConfirmAction(sb As ScreenBuffer, msg As String) As Boolean
        DrawBoxColor(sb, 30, 11, 40, 5, " CONFIRMAÇÃO ", ConsoleColor.Red, ConsoleColor.White)
        sb.WriteText(32, 13, $"{msg} (S/N)", ConsoleColor.White, ConsoleColor.Red)
        sb.Render()
        Dim res = Console.ReadKey(True).Key
        Return (res = ConsoleKey.S)
    End Function

    Sub ShowSplash()
        Console.BackgroundColor = ConsoleColor.Black : Console.Clear()
        Dim cX = 30, cY = 12
        DrawBoxColorLegacy(cX, cY, 40, 6, " INICIALIZANDO ", ConsoleColor.Black, ConsoleColor.White)

        Dim tarefas = {"Drivers", "WMI", "Network", "CPU"}
        For i = 0 To tarefas.Length - 1
            Console.SetCursorPosition(cX + 2, cY + 2)
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write($"Carregando {tarefas(i)}...".PadRight(20))
            Console.SetCursorPosition(cX + 2, cY + 4)
            DrawSimpleBarLegacy(i + 1, tarefas.Length, ConsoleColor.Yellow, 36)
            Threading.Thread.Sleep(600)
        Next
    End Sub

    Sub DrawStaticLabels(sb As ScreenBuffer)
        Dim infoX = 25 : Dim y = 2
        sb.WriteText(infoX, y, "OS:      " & cachedOS, ConsoleColor.White, ConsoleColor.DarkBlue)
        sb.WriteText(infoX, y + 1, "KERNEL:  " & cachedKernel, ConsoleColor.White, ConsoleColor.DarkBlue)
        sb.WriteText(infoX, y + 2, "REDE:    ", ConsoleColor.White, ConsoleColor.DarkBlue)

        DrawBox(sb, infoX - 1, y + 4, 72, 5, " PROCESSADOR E PLACA-MAE ")
        DrawBox(sb, infoX - 1, y + 10, 72, 4, " SUBSISTEMA DE VIDEO ")
        DrawBox(sb, infoX - 1, y + 15, 72, 4, " MONITORAMENTO DE CARGA ")

        sb.WriteText(infoX + 1, y + 6, "CPU: " & cachedCPU, ConsoleColor.Cyan, ConsoleColor.DarkBlue)
        sb.WriteText(infoX + 1, y + 7, "M/B: " & cachedMB, ConsoleColor.Cyan, ConsoleColor.DarkBlue)
        sb.WriteText(infoX + 1, y + 11, "GPU: " & cachedGPU, ConsoleColor.Cyan, ConsoleColor.DarkBlue)
        sb.WriteText(infoX + 1, y + 12, "VRAM: " & cachedVRAM, ConsoleColor.Cyan, ConsoleColor.DarkBlue)
    End Sub

    ''' <summary>Atualiza SÓ o que realmente muda a cada tick: rede, % de CPU e RAM.
    ''' OS/Kernel/CPU/GPU/VRAM saíram daqui - vêm do cache, sem custo de WMI.</summary>
    Sub UpdateMainDynamicValues(sb As ScreenBuffer)
        Dim infoX = 34 : Dim y = 2

        sb.WriteText(infoX, y + 2, GetActiveNetwork().PadRight(60), ConsoleColor.Cyan, ConsoleColor.DarkBlue)

        Dim rT = My.Computer.Info.TotalPhysicalMemory / (1024 ^ 3)
        Dim rU = (My.Computer.Info.TotalPhysicalMemory - My.Computer.Info.AvailablePhysicalMemory) / (1024 ^ 3)
        Dim cU = If(cpuTotal IsNot Nothing, cpuTotal.NextValue(), 0)

        sb.WriteText(infoX + 1, y + 16, "CPU: ", ConsoleColor.White, ConsoleColor.DarkBlue)
        Dim nx1 = DrawSimpleBar(sb, infoX + 6, y + 16, cU, 100, ConsoleColor.Green, 30)
        sb.WriteText(nx1, y + 16, $" {cU:F0}%  ", ConsoleColor.White, ConsoleColor.DarkBlue)

        sb.WriteText(infoX + 1, y + 17, "RAM: ", ConsoleColor.White, ConsoleColor.DarkBlue)
        Dim nx2 = DrawSimpleBar(sb, infoX + 6, y + 17, rU, rT, ConsoleColor.Yellow, 30)
        sb.WriteText(nx2, y + 17, $" {rU:F1}/{rT:F1}GB  ", ConsoleColor.White, ConsoleColor.DarkBlue)
    End Sub

    Sub ShowFullMonitor(sb As ScreenBuffer)
        sb.Clear(ConsoleColor.DarkBlue)
        DrawBox(sb, 1, 1, 97, 26, " MONITOR DE RECURSOS DETALHADO ")
        DrawBox(sb, 3, 4, 45, 21, " PROCESSAMENTO (CORES) ")
        DrawBox(sb, 49, 4, 47, 8, " MEMORIA RAM ")
        DrawBox(sb, 49, 12, 47, 5, " VIDEO / GPU ")
        DrawBox(sb, 49, 17, 47, 8, " SLOTS FISICOS (RAM) ")

        sb.WriteText(4, 3, "CPU: " & cachedCPU, ConsoleColor.Yellow, ConsoleColor.DarkBlue)
        sb.WriteText(51, 14, "GPU: " & Truncate(cachedGPU, 35), ConsoleColor.White, ConsoleColor.DarkBlue)
        sb.WriteText(51, 15, "VRAM: " & cachedVRAM, ConsoleColor.White, ConsoleColor.DarkBlue)

        Dim sy = 19
        For Each linha In cachedRamSlots
            If sy <= 23 Then
                sb.WriteText(51, sy, linha, ConsoleColor.Gray, ConsoleColor.DarkBlue)
                sy += 1
            End If
        Next

        DrawMenuFooter(sb, "  ESC Voltar ao Menu Principal ")
        sb.Render()

        While Not (Console.KeyAvailable AndAlso Console.ReadKey(True).Key = ConsoleKey.Escape)
            For i = 0 To Math.Min(coreCounters.Count - 1, 15)
                Dim val = coreCounters(i).NextValue()
                sb.WriteText(5, 6 + i, $"C{i:00}: ", ConsoleColor.White, ConsoleColor.DarkBlue)
                Dim nextX = DrawSimpleBar(sb, 10, 6 + i, val, 100, ConsoleColor.Cyan, 15)
                sb.WriteText(nextX, 6 + i, $" {val:F0}%  ", ConsoleColor.White, ConsoleColor.DarkBlue)
            Next

            Dim ramT = My.Computer.Info.TotalPhysicalMemory / (1024 ^ 3)
            Dim ramL = My.Computer.Info.AvailablePhysicalMemory / (1024 ^ 3)
            Dim ramU = ramT - ramL

            sb.WriteText(51, 6, "USO: ", ConsoleColor.White, ConsoleColor.DarkBlue)
            DrawSimpleBar(sb, 56, 6, ramU, ramT, ConsoleColor.Yellow, 20)

            sb.WriteText(51, 8, $"TOTAL: {ramT:F1} GB".PadRight(20), ConsoleColor.White, ConsoleColor.DarkBlue)
            sb.WriteText(51, 9, $"USADA: {ramU:F1} GB".PadRight(20), ConsoleColor.White, ConsoleColor.DarkBlue)
            sb.WriteText(51, 10, $"LIVRE: {ramL:F1} GB".PadRight(20), ConsoleColor.White, ConsoleColor.DarkBlue)

            sb.Render()
            Threading.Thread.Sleep(700)
        End While
    End Sub

    ' --- MÉTODOS DE APOIO ---
    ''' <summary>Retorna o tipo de conexão (Ethernet/Wi-Fi/Outro) com base no NetworkInterfaceType.</summary>
    Function GetConnectionTypeLabel(ni As NetworkInterface) As String
        Select Case ni.NetworkInterfaceType
            Case NetworkInterfaceType.Wireless80211
                Return "Wi-Fi"
            Case NetworkInterfaceType.Ethernet, NetworkInterfaceType.GigabitEthernet, NetworkInterfaceType.FastEthernetT, NetworkInterfaceType.FastEthernetFx
                Return "Ethernet"
            Case NetworkInterfaceType.Ppp, NetworkInterfaceType.Tunnel
                Return "VPN/PPP"
            Case Else
                Return ni.NetworkInterfaceType.ToString()
        End Select
    End Function

    Function GetActiveNetwork() As String
        Try
            For Each ni In NetworkInterface.GetAllNetworkInterfaces()
                If ni.OperationalStatus = OperationalStatus.Up AndAlso Not ni.Description.ToLower().Contains("virtual") Then
                    Dim props = ni.GetIPProperties()
                    Dim ip = props.UnicastAddresses.FirstOrDefault(Function(x) x.Address.AddressFamily = AddressFamily.InterNetwork)
                    If ip IsNot Nothing Then
                        Dim tipo = GetConnectionTypeLabel(ni)
                        Return $"[{tipo}] {Truncate(ni.Description, 20)} [{ip.Address}]"
                    End If
                End If
            Next
        Catch : End Try
        Return "Sem Conexão"
    End Function

    Sub ShowPnPDevices(sb As ScreenBuffer)
        Dim devs As New List(Of ManagementObject)
        Using s As New ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity")
            For Each d In s.Get() : devs.Add(CType(d, ManagementObject)) : Next
        End Using
        Dim idx = 0
        While True
            sb.Clear(ConsoleColor.DarkBlue)
            DrawBox(sb, 2, 1, 95, 25, " GERENCIADOR DE HARDWARE ")
            For i = 0 To 19
                If idx + i < devs.Count Then
                    Dim name = If(devs(idx + i)("Name") IsNot Nothing, devs(idx + i)("Name").ToString(), "Desconhecido")
                    Dim line = Truncate(name, 85).PadRight(88)
                    If i = 0 Then
                        sb.WriteText(4, 3 + i, "> " & line, ConsoleColor.Black, ConsoleColor.Cyan)
                    Else
                        sb.WriteText(4, 3 + i, "  " & line, ConsoleColor.White, ConsoleColor.DarkBlue)
                    End If
                End If
            Next
            DrawMenuFooter(sb, " [SETAS] Navegar | [ENTER] Detalhes | [ESC] Sair ")
            sb.Render()

            Dim k = Console.ReadKey(True)
            If k.Key = ConsoleKey.Escape Then Exit While
            If k.Key = ConsoleKey.DownArrow And idx < devs.Count - 1 Then idx += 1
            If k.Key = ConsoleKey.UpArrow And idx > 0 Then idx -= 1
            If k.Key = ConsoleKey.Enter Then
                Dim d = devs(idx)
                DrawBoxColor(sb, 15, 6, 70, 14, " PROPRIEDADES DO HARDWARE ", ConsoleColor.Gray, ConsoleColor.Black)
                sb.WriteText(18, 9, "NOME: " & Truncate(d("Name"), 60), ConsoleColor.Black, ConsoleColor.Gray)
                sb.WriteText(18, 10, "FABRICANTE: " & If(d("Manufacturer")?.ToString(), "N/A"), ConsoleColor.Black, ConsoleColor.Gray)
                sb.WriteText(18, 11, "ID PNP: " & Truncate(d("PNPDeviceID"), 55), ConsoleColor.Black, ConsoleColor.Gray)
                sb.WriteText(18, 12, "STATUS: " & If(d("Status")?.ToString(), "N/A"), ConsoleColor.Black, ConsoleColor.Gray)
                sb.WriteText(15 + 18, 18, " Pressione qualquer tecla ", ConsoleColor.DarkRed, ConsoleColor.Gray)
                sb.Render()
                Console.ReadKey(True)
            End If
        End While
    End Sub

    Sub DrawWindowsLogo(sb As ScreenBuffer, x As Integer, y As Integer)
        Dim b = ChrW(&H2588)
        For i = 0 To 3
            sb.WriteText(x, y + i, New String(b, 7) & " " & New String(b, 7), ConsoleColor.Cyan, ConsoleColor.DarkBlue)
        Next
        For i = 5 To 8
            sb.WriteText(x, y + i, New String(b, 7) & " " & New String(b, 7), ConsoleColor.Cyan, ConsoleColor.DarkBlue)
        Next
    End Sub

    Sub DrawBoxColor(sb As ScreenBuffer, x As Integer, y As Integer, w As Integer, h As Integer, title As String, bg As ConsoleColor, fg As ConsoleColor)
        sb.WriteText(x, y, "╔" & New String("═"c, w - 2) & "╗", fg, bg)
        For i = 1 To h - 2
            sb.PutChar(x, y + i, "║"c, fg, bg)
            For fillX = x + 1 To x + w - 2
                sb.PutChar(fillX, y + i, " "c, fg, bg)
            Next
            sb.PutChar(x + w - 1, y + i, "║"c, fg, bg)
        Next
        sb.WriteText(x, y + h - 1, "╚" & New String("═"c, w - 2) & "╝", fg, bg)
        If title <> "" Then
            sb.WriteText(x + (w \ 2) - (title.Length \ 2), y, title, fg, bg)
        End If
    End Sub

    Sub DrawBox(sb As ScreenBuffer, x As Integer, y As Integer, w As Integer, h As Integer, title As String)
        DrawBoxColor(sb, x, y, w, h, title, ConsoleColor.DarkBlue, ConsoleColor.White)
    End Sub

    ''' <summary>Desenha a barra no buffer e retorna o X logo após o fim dela,
    ''' para o chamador poder continuar escrevendo texto em seguida.</summary>
    Function DrawSimpleBar(sb As ScreenBuffer, x As Integer, y As Integer, atual As Double, max As Double, cor As ConsoleColor, tam As Integer) As Integer
        Dim p As Integer = Math.Max(0, Math.Min(tam, CInt((atual / max) * tam)))
        For i = 0 To p - 1
            sb.PutChar(x + i, y, "█"c, cor, ConsoleColor.DarkBlue)
        Next
        For i = p To tam - 1
            sb.PutChar(x + i, y, "▒"c, ConsoleColor.DarkGray, ConsoleColor.DarkBlue)
        Next
        Return x + tam
    End Function

    Function GetWMI(cl As String, pr As String) As String
        Try
            Using s As New ManagementObjectSearcher($"SELECT {pr} FROM {cl}")
                For Each o In s.Get() : Return o(pr).ToString().Trim() : Next
            End Using
        Catch : End Try
        Return "N/A"
    End Function

    Function GetVRAM() As String
        Try
            Using s As New ManagementObjectSearcher("SELECT AdapterRAM FROM Win32_VideoController")
                For Each o In s.Get() : Return (Math.Abs(Val(o("AdapterRAM"))) / (1024 ^ 3)).ToString("F0") & " GB" : Next
            End Using
        Catch : End Try
        Return "N/A"
    End Function

    Sub InitCounters()
        Try
            cpuTotal = New PerformanceCounter("Processor", "% Processor Time", "_Total")
            cpuTotal.NextValue()
            For i = 0 To Math.Min(Environment.ProcessorCount, 16) - 1
                coreCounters.Add(New PerformanceCounter("Processor", "% Processor Time", i.ToString()))
            Next
        Catch : End Try
    End Sub

    Sub DrawMenuFooter(sb As ScreenBuffer, text As String)
        sb.WriteText(0, 29, text.PadRight(sb.Width), ConsoleColor.Black, ConsoleColor.Gray)
    End Sub

    Function Truncate(value As Object, max As Integer) As String
        If value Is Nothing Then Return "N/A"
        Dim s = value.ToString()
        Return If(s.Length <= max, s, s.Substring(0, max - 3) & "...")
    End Function

    ' --- Versões "legacy" usadas SÓ na splash screen (roda 1x, sem loop contínuo,
    ' então não sofre com flicker; mantidas simples com Console direto). ---
    Sub DrawBoxColorLegacy(x As Integer, y As Integer, w As Integer, h As Integer, title As String, bg As ConsoleColor, fg As ConsoleColor)
        Console.BackgroundColor = bg : Console.ForegroundColor = fg
        Console.SetCursorPosition(x, y) : Console.Write("╔" & New String("═"c, w - 2) & "╗")
        For i = 1 To h - 2 : Console.SetCursorPosition(x, y + i) : Console.Write("║" & New String(" "c, w - 2) & "║") : Next
        Console.SetCursorPosition(x, y + h - 1) : Console.Write("╚" & New String("═"c, w - 2) & "╝")
        If title <> "" Then
            Console.SetCursorPosition(x + (w \ 2) - (title.Length \ 2), y) : Console.Write(title)
        End If
    End Sub

    Sub DrawSimpleBarLegacy(atual As Double, max As Double, cor As ConsoleColor, tam As Integer)
        Dim p As Integer = Math.Max(0, Math.Min(tam, CInt((atual / max) * tam)))
        Console.ForegroundColor = cor : Console.Write(New String("█"c, p))
        Console.ForegroundColor = ConsoleColor.DarkGray : Console.Write(New String("▒"c, tam - p))
        Console.ForegroundColor = ConsoleColor.White
    End Sub
End Module