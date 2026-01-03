Imports System.Management
Imports System.Diagnostics
Imports Microsoft.Win32
Imports System.Net.NetworkInformation
Imports System.Net.Sockets

Module Module1
    Dim cpuTotal As PerformanceCounter
    Dim coreCounters As New List(Of PerformanceCounter)

    Sub Main()
        Try
            ' Configuração de Buffer maior que a Janela para evitar scroll/bipe
            Console.Title = "SYSTEM INSPECTOR - v3.5"
            Console.SetWindowSize(100, 31)
            Console.SetBufferSize(100, 35) ' Buffer maior evita o bipe de scroll
            Console.CursorVisible = False
        Catch : End Try

        ShowSplash()
        InitCounters()

        While True
            Console.BackgroundColor = ConsoleColor.DarkBlue
            Console.Clear()
            ' Reduzi a altura da caixa (27) para não encostar no final do buffer
            DrawBox(0, 0, 98, 27, $" {Environment.MachineName} - System Diagnostic ")
            DrawWindowsLogo(3, 2)
            DrawStaticLabels()
            DrawMenuFooter("  D Dispositivos | M Monitor Full | ESC Sair ")

            While Not Console.KeyAvailable
                UpdateMainDynamicValues()
                Threading.Thread.Sleep(500)
            End While

            Dim key = Console.ReadKey(True).Key
            Select Case key
                Case ConsoleKey.D : ShowPnPDevices()
                Case ConsoleKey.M : ShowFullMonitor()
                Case ConsoleKey.Escape : Exit While
            End Select
        End While
    End Sub

    ' --- TELAS ---

    Sub ShowFullMonitor()
        Console.BackgroundColor = ConsoleColor.DarkBlue : Console.Clear()

        ' Moldura principal
        DrawBox(1, 1, 96, 27, " MONITOR DE RECURSOS DETALHADO ")

        ' Áreas fixas
        Console.SetCursorPosition(4, 3) : Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write("PROCESSADOR: " & GetWMI("Win32_Processor", "Name"))

        Console.SetCursorPosition(51, 10) : Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write("[ SLOTS RAM DETALHADOS ]")

        DrawMenuFooter("  ESC Voltar ao Menu Principal ")

        While Not (Console.KeyAvailable AndAlso Console.ReadKey(True).Key = ConsoleKey.Escape)
            ' CPU Cores (Lado Esquerdo)
            For i = 0 To Math.Min(coreCounters.Count - 1, 15)
                Console.SetCursorPosition(4, 5 + i)
                Dim val = coreCounters(i).NextValue()
                Console.ForegroundColor = ConsoleColor.White : Console.Write($"C{i:00}: ")
                DrawSimpleBar(val, 100, ConsoleColor.Cyan, 12) : Console.Write($" {val:F0}%  ")
            Next

            ' RAM e GPU (Lado Direito)
            Dim ramT = My.Computer.Info.TotalPhysicalMemory / (1024 ^ 3)
            Dim ramU = (My.Computer.Info.TotalPhysicalMemory - My.Computer.Info.AvailablePhysicalMemory) / (1024 ^ 3)
            Console.SetCursorPosition(51, 5) : Console.ForegroundColor = ConsoleColor.White
            Console.Write("RAM USO: ") : DrawSimpleBar(ramU, ramT, ConsoleColor.Yellow, 15)
            Console.Write($" {ramU:F1}/{ramT:F1}GB ")

            Console.SetCursorPosition(51, 7) : Console.Write("GPU: " & GetWMI("Win32_VideoController", "Name").PadRight(35))
            Console.SetCursorPosition(51, 8) : Console.Write("VRAM: " & GetVRAM().PadRight(20))

            ' SLOTS DE RAM (Abaixo da GPU)
            Dim sy = 11
            Using s As New ManagementObjectSearcher("SELECT Capacity, Speed, DeviceLocator FROM Win32_PhysicalMemory")
                For Each obj In s.Get()
                    If sy < 25 Then
                        Console.SetCursorPosition(51, sy) : Console.ForegroundColor = ConsoleColor.Gray
                        Dim cap = (Val(obj("Capacity")) / 1024 ^ 3).ToString("F0")
                        Console.Write($"• {obj("DeviceLocator")}: {cap}GB @ {obj("Speed")}MHz".PadRight(40))
                        sy += 1
                    End If
                Next
            End Using

            Threading.Thread.Sleep(700)
        End While
    End Sub

    ' --- LÓGICA DE REDE (REAL IP) ---

    Function GetActiveNetwork() As String
        Try
            ' Filtro rigoroso para interfaces físicas e ativas
            Dim ni = NetworkInterface.GetAllNetworkInterfaces().
                FirstOrDefault(Function(n) n.OperationalStatus = OperationalStatus.Up AndAlso
                n.NetworkInterfaceType <> NetworkInterfaceType.Loopback AndAlso
                Not n.Description.ToLower().Contains("virtual") AndAlso
                Not n.Description.ToLower().Contains("pseudo"))

            If ni IsNot Nothing Then
                Dim ip = ni.GetIPProperties().UnicastAddresses.
                    FirstOrDefault(Function(ua) ua.Address.AddressFamily = AddressFamily.InterNetwork)
                If ip IsNot Nothing Then
                    Return $"{ni.Description} [{ip.Address}]"
                End If
            End If
        Catch : End Try
        Return "Buscando conexão real..."
    End Function

    ' --- VISUAIS DA TELA PRINCIPAL ---

    Sub UpdateMainDynamicValues()
        Dim infoX = 25 : Dim y = 2
        Console.ForegroundColor = ConsoleColor.White

        ' Informações de Sistema (Recuperadas)
        Console.SetCursorPosition(infoX + 9, y) : Console.Write(My.Computer.Info.OSFullName.PadRight(55))
        Console.SetCursorPosition(infoX + 9, y + 1) : Console.Write(Environment.OSVersion.VersionString.PadRight(55))
        Console.SetCursorPosition(infoX + 9, y + 2) : Console.Write(GetActiveNetwork().PadRight(55))

        ' GPU
        Console.SetCursorPosition(infoX + 5, y + 5) : Console.Write(GetWMI("Win32_VideoController", "Name").PadRight(55))
        Console.SetCursorPosition(infoX + 6, y + 6) : Console.Write(GetVRAM().PadRight(20))

        ' Performance
        Dim ramT = My.Computer.Info.TotalPhysicalMemory / (1024 ^ 3)
        Dim ramU = (My.Computer.Info.TotalPhysicalMemory - My.Computer.Info.AvailablePhysicalMemory) / (1024 ^ 3)
        Dim cpuU = If(cpuTotal IsNot Nothing, cpuTotal.NextValue(), 0)

        Console.SetCursorPosition(infoX + 5, y + 14) : DrawSimpleBar(cpuU, 100, ConsoleColor.Green, 20) : Console.Write($" {cpuU:F0}%  ")
        Console.SetCursorPosition(infoX + 5, y + 15) : DrawSimpleBar(ramU, ramT, ConsoleColor.Yellow, 20) : Console.Write($" {ramU:F1}/{ramT:F1}GB  ")
    End Sub

    Sub DrawStaticLabels()
        Dim infoX = 25 : Dim y = 2
        Console.ForegroundColor = ConsoleColor.White
        Console.SetCursorPosition(infoX, y) : Console.Write("OS:      ")
        Console.SetCursorPosition(infoX, y + 1) : Console.Write("KERNEL:  ")
        Console.SetCursorPosition(infoX, y + 2) : Console.Write("REDE:    ")

        Console.SetCursorPosition(infoX, y + 4) : Console.ForegroundColor = ConsoleColor.Yellow : Console.Write("[ PLACA DE VÍDEO ]")
        Console.ForegroundColor = ConsoleColor.White
        Console.SetCursorPosition(infoX, y + 5) : Console.Write("GPU: ")
        Console.SetCursorPosition(infoX, y + 6) : Console.Write("VRAM: ")

        Console.SetCursorPosition(infoX, y + 13) : Console.ForegroundColor = ConsoleColor.Yellow : Console.Write("[ RECURSOS DO SISTEMA ]")
        Console.ForegroundColor = ConsoleColor.White
        Console.SetCursorPosition(infoX, y + 14) : Console.Write("CPU: ")
        Console.SetCursorPosition(infoX, y + 15) : Console.Write("RAM: ")
    End Sub

    Sub DrawMenuFooter(text As String)
        ' Posicionado na linha 28, longe da borda 31 para evitar o bipe
        Console.SetCursorPosition(0, 28)
        Console.BackgroundColor = ConsoleColor.Gray
        Console.ForegroundColor = ConsoleColor.Black
        Console.Write(text.PadRight(100))
        Console.BackgroundColor = ConsoleColor.DarkBlue
    End Sub

    ' --- FUNÇÕES DE BASE (Splash, Counters, WMI) ---

    Sub ShowSplash()
        Console.BackgroundColor = ConsoleColor.Black : Console.Clear()
        DrawBox(30, 10, 40, 8, " CARREGANDO SISTEMA ")
        Dim stp() As String = {"WMI Engine", "DirectX GPU", "Physical Network", "Performance"}
        For i = 0 To stp.Length - 1
            Console.SetCursorPosition(32, 12) : Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write("Status: " & stp(i).PadRight(20))
            Console.SetCursorPosition(32, 14)
            DrawSimpleBar(i + 1, stp.Length, ConsoleColor.Yellow, 36)
            Threading.Thread.Sleep(400)
        Next
    End Sub

    Sub DrawWindowsLogo(x As Integer, y As Integer)
        Dim b As String = ChrW(&H2588)
        Console.ForegroundColor = ConsoleColor.Cyan
        For i = 0 To 2 : Console.SetCursorPosition(x, y + i) : Console.Write(New String(b, 8) & "  " & New String(b, 8)) : Next
        Console.SetCursorPosition(x, y + 3) : Console.Write(New String(" "c, 18))
        For i = 4 To 6 : Console.SetCursorPosition(x, y + i) : Console.Write(New String(b, 8) & "  " & New String(b, 8)) : Next
    End Sub

    Sub InitCounters()
        Try
            coreCounters.Clear()
            cpuTotal = New PerformanceCounter("Processor", "% Processor Time", "_Total")
            cpuTotal.NextValue()
            Dim cCount = Math.Min(Environment.ProcessorCount, 16)
            For i = 0 To cCount - 1
                Dim p = New PerformanceCounter("Processor", "% Processor Time", i.ToString())
                p.NextValue()
                coreCounters.Add(p)
            Next
        Catch : End Try
    End Sub

    Function GetVRAM() As String
        Try
            Using s As New ManagementObjectSearcher("SELECT AdapterRAM FROM Win32_VideoController")
                For Each o In s.Get() : Return (Math.Abs(Val(o("AdapterRAM"))) / (1024 ^ 3)).ToString("F0") & " GB" : Next
            End Using
        Catch : End Try
        Return "N/A"
    End Function

    Sub ShowPnPDevices()
        Console.BackgroundColor = ConsoleColor.DarkBlue : Console.Clear()
        Dim devs As New List(Of String)
        Using s As New ManagementObjectSearcher("SELECT Name FROM Win32_PnPEntity")
            For Each d In s.Get() : If d("Name") IsNot Nothing Then devs.Add(d("Name").ToString())
            Next
        End Using
        Dim idx = 0
        While True
            Console.BackgroundColor = ConsoleColor.DarkBlue : Console.Clear()
            DrawBox(2, 1, 95, 25, " DISPOSITIVOS DETECTADOS ")
            For i = 0 To 19
                If idx + i < devs.Count Then
                    Console.SetCursorPosition(4, 3 + i)
                    Console.ForegroundColor = If(i = 0, ConsoleColor.Yellow, ConsoleColor.White)
                    Console.Write(If(i = 0, ">> ", "   ") & devs(idx + i).PadRight(85))
                End If
            Next
            DrawMenuFooter("  [SETAS] Navegar | [ESC] Sair ")
            Dim k = Console.ReadKey(True).Key
            If k = ConsoleKey.Escape Then Exit While
            If k = ConsoleKey.DownArrow And idx < devs.Count - 20 Then idx += 1
            If k = ConsoleKey.UpArrow And idx > 0 Then idx -= 1
        End While
    End Sub

    Sub DrawBox(x As Integer, y As Integer, w As Integer, h As Integer, title As String)
        Console.ForegroundColor = ConsoleColor.White
        Console.SetCursorPosition(x, y) : Console.Write("╔" & New String("═"c, w - 2) & "╗")
        For i = 1 To h - 2 : Console.SetCursorPosition(x, y + i) : Console.Write("║" & New String(" "c, w - 2) & "║") : Next
        Console.SetCursorPosition(x, y + h - 1) : Console.Write("╚" & New String("═"c, w - 2) & "╝")
        If title <> "" Then
            Console.SetCursorPosition(x + (w \ 2) - (title.Length \ 2), y)
            Console.ForegroundColor = ConsoleColor.Yellow : Console.Write(title)
        End If
    End Sub

    Sub DrawSimpleBar(atual As Double, max As Double, cor As ConsoleColor, tam As Integer)
        Dim p As Integer = Math.Max(0, Math.Min(tam, CInt((atual / max) * tam)))
        Console.ForegroundColor = cor : Console.Write(New String("█"c, p))
        Console.ForegroundColor = ConsoleColor.DarkGray : Console.Write(New String("▒"c, tam - p))
        Console.ForegroundColor = ConsoleColor.White
    End Sub

    Function GetWMI(cl As String, pr As String) As String
        Try
            Using s As New ManagementObjectSearcher($"SELECT {pr} FROM {cl}")
                For Each o In s.Get() : If o(pr) IsNot Nothing Then Return o(pr).ToString().Trim()
                Next
            End Using
        Catch : End Try
        Return "N/A"
    End Function
End Module