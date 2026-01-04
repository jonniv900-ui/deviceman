Imports System.Management
Imports System.Diagnostics
Imports Microsoft.Win32
Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports System.IO

Module Module1
    Dim cpuTotal As PerformanceCounter
    Dim coreCounters As New List(Of PerformanceCounter)

    Sub Main()
        Try
            Console.Title = "SYSTEM INSPECTOR - v3.9"
            Console.SetWindowSize(100, 32)
            Console.SetBufferSize(100, 32)
            Console.CursorVisible = False
        Catch : End Try

        ShowSplash()
        InitCounters()

        While True
            Console.BackgroundColor = ConsoleColor.DarkBlue : Console.Clear()
            DrawBox(0, 0, 99, 28, $" {Environment.MachineName} - System Diagnostic ")
            DrawWindowsLogo(4, 2)
            DrawStaticLabels()
            DrawMenuFooter("  D Dispositivos | M Monitor Full | F Ferramentas | ESC Sair ")
            While Not Console.KeyAvailable
                UpdateMainDynamicValues()
                Threading.Thread.Sleep(500)
            End While

            Dim key = Console.ReadKey(True).Key
            Select Case key
                Case ConsoleKey.D : ShowPnPDevices()
                Case ConsoleKey.M : ShowFullMonitor()
                Case ConsoleKey.F : ShowTools()
                Case ConsoleKey.Escape : Exit While
            End Select
        End While
    End Sub
    Sub ShowTools()
        Dim toolsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools")

        If Not Directory.Exists(toolsPath) Then
            Try : Directory.CreateDirectory(toolsPath) : Catch : End Try
        End If

        Dim files = Directory.GetFiles(toolsPath, "*.*").Where(Function(f) _
            f.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) Or
            f.EndsWith(".bat", StringComparison.OrdinalIgnoreCase) Or
            f.EndsWith(".ps1", StringComparison.OrdinalIgnoreCase)).ToList()

        Dim idx = 0
        Dim startIdx = 0

        While True
            Console.BackgroundColor = ConsoleColor.DarkBlue : Console.Clear()

            Dim titulo = $" CAIXA DE FERRAMENTAS ({files.Count} itens) "
            DrawBox(5, 4, 90, 22, titulo)

            If files.Count = 0 Then
                Console.SetCursorPosition(10, 12) : Console.ForegroundColor = ConsoleColor.Red
                Console.Write("Nenhuma ferramenta encontrada na pasta \tools.")
            Else
                For i = 0 To 14
                    Dim fileIdx = startIdx + i
                    If fileIdx < files.Count Then
                        Dim fullPath = files(fileIdx)
                        Dim fileName = Path.GetFileName(fullPath)
                        Dim description = ""

                        Try
                            If fileName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) Then
                                Dim info = FileVersionInfo.GetVersionInfo(fullPath)
                                description = If(Not String.IsNullOrEmpty(info.FileDescription), $" ({info.FileDescription})", "")
                            ElseIf fileName.EndsWith(".bat", StringComparison.OrdinalIgnoreCase) Then
                                description = " [Batch Script]"
                            ElseIf fileName.EndsWith(".ps1", StringComparison.OrdinalIgnoreCase) Then
                                description = " [PowerShell]"
                            End If
                        Catch : End Try

                        Console.SetCursorPosition(8, 6 + i)
                        If fileIdx = idx Then
                            Console.BackgroundColor = ConsoleColor.Cyan : Console.ForegroundColor = ConsoleColor.Black
                            Console.Write($"> {Truncate(fileName & description, 80).PadRight(83)}")
                            Console.BackgroundColor = ConsoleColor.DarkBlue
                        Else
                            Console.ForegroundColor = ConsoleColor.White
                            Console.Write($"  {Truncate(fileName & description, 80).PadRight(83)}")
                        End If
                    End If
                Next
            End If

            ' Rodapé atualizado com comandos de energia
            DrawMenuFooter(" [SETAS] Navegar | [ENTER] Lançar | [R] Reiniciar | [S] Desligar | [ESC] Voltar ")

            Dim k = Console.ReadKey(True)
            If k.Key = ConsoleKey.Escape Then Exit While

            ' Comandos de Energia (WinPE utiliza shutdown.exe padrão)
            If k.Key = ConsoleKey.R Then
                If ConfirmAction("REINICIAR O SISTEMA?") Then Process.Start("shutdown", "/r /t 0")
            End If
            If k.Key = ConsoleKey.S Then
                If ConfirmAction("DESLIGAR O COMPUTADOR?") Then Process.Start("shutdown", "/s /t 0")
            End If

            ' Navegação
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
                    DrawBoxColor(15, 10, 70, 7, " FALHA NO LANÇAMENTO ", ConsoleColor.DarkRed, ConsoleColor.White)
                    Console.SetCursorPosition(17, 12) : Console.Write("Erro: " & Truncate(ex.Message, 60))
                    Console.ReadKey(True)
                End Try
            End If
        End While
    End Sub

    ' Função auxiliar para evitar desligamentos acidentais
    Function ConfirmAction(msg As String) As Boolean
        DrawBoxColor(30, 11, 40, 5, " CONFIRMAÇÃO ", ConsoleColor.Red, ConsoleColor.White)
        Console.SetCursorPosition(32, 13) : Console.Write($"{msg} (S/N)")
        Dim res = Console.ReadKey(True).Key
        Return (res = ConsoleKey.S)
    End Function

    ' --- CORREÇÃO: SPLASH SCREEN ---
    Sub ShowSplash()
        Console.BackgroundColor = ConsoleColor.Black : Console.Clear()
        Dim cX = 30, cY = 12
        DrawBoxColor(cX, cY, 40, 6, " INICIALIZANDO ", ConsoleColor.Black, ConsoleColor.White)

        Dim tarefas = {"Drivers", "WMI", "Network", "CPU"}
        For i = 0 To tarefas.Length - 1
            Console.SetCursorPosition(cX + 2, cY + 2)
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write($"Carregando {tarefas(i)}...".PadRight(20))
            Console.SetCursorPosition(cX + 2, cY + 4)
            DrawSimpleBar(i + 1, tarefas.Length, ConsoleColor.Yellow, 36)
            Threading.Thread.Sleep(600)
        Next
    End Sub

    ' --- CORREÇÃO: LABELS (OS E KERNEL) ---
    Sub DrawStaticLabels()
        Dim infoX = 25 : Dim y = 2
        Console.ForegroundColor = ConsoleColor.White
        ' Labels que haviam sumido
        Console.SetCursorPosition(infoX, y) : Console.Write("OS:      ")
        Console.SetCursorPosition(infoX, y + 1) : Console.Write("KERNEL:  ")
        Console.SetCursorPosition(infoX, y + 2) : Console.Write("REDE:    ")

        DrawBox(infoX - 1, y + 4, 72, 5, " PROCESSADOR E PLACA-MAE ")
        DrawBox(infoX - 1, y + 10, 72, 4, " SUBSISTEMA DE VIDEO ")
        DrawBox(infoX - 1, y + 15, 72, 4, " MONITORAMENTO DE CARGA ")

        Console.SetCursorPosition(infoX + 1, y + 6) : Console.Write("CPU: " & Truncate(GetWMI("Win32_Processor", "Name"), 60))
        Console.SetCursorPosition(infoX + 1, y + 7) : Console.Write("M/B: " & GetWMI("Win32_BaseBoard", "Product") & " | BIOS: " & GetWMI("Win32_BIOS", "Version"))
    End Sub

    Sub UpdateMainDynamicValues()
        Dim infoX = 34 : Dim y = 2
        Console.ForegroundColor = ConsoleColor.Cyan

        ' Preenchimento dos dados de OS e Kernel
        Console.SetCursorPosition(infoX, y) : Console.Write(Truncate(My.Computer.Info.OSFullName, 50))
        Console.SetCursorPosition(infoX, y + 1) : Console.Write(Environment.OSVersion.VersionString.PadRight(50))
        Console.SetCursorPosition(infoX, y + 2) : Console.Write(GetActiveNetwork().PadRight(60))

        Console.SetCursorPosition(infoX - 1, y + 11) : Console.Write("GPU: " & Truncate(GetWMI("Win32_VideoController", "Name"), 50))
        Console.SetCursorPosition(infoX - 1, y + 12) : Console.Write("VRAM: " & GetVRAM())

        Dim rT = My.Computer.Info.TotalPhysicalMemory / (1024 ^ 3)
        Dim rU = (My.Computer.Info.TotalPhysicalMemory - My.Computer.Info.AvailablePhysicalMemory) / (1024 ^ 3)
        Dim cU = If(cpuTotal IsNot Nothing, cpuTotal.NextValue(), 0)

        Console.SetCursorPosition(infoX - 1, y + 16) : Console.ForegroundColor = ConsoleColor.White : Console.Write("CPU: ")
        DrawSimpleBar(cU, 100, ConsoleColor.Green, 30) : Console.Write($" {cU:F0}%  ")
        Console.SetCursorPosition(infoX - 1, y + 17) : Console.ForegroundColor = ConsoleColor.White : Console.Write("RAM: ")
        DrawSimpleBar(rU, rT, ConsoleColor.Yellow, 30) : Console.Write($" {rU:F1}/{rT:F1}GB  ")
    End Sub
    ' --- TELA DE MONITORAMENTO (CORREÇÃO FINAL: 4 SLOTS E LAYOUT) ---
    Sub ShowFullMonitor()
        Console.BackgroundColor = ConsoleColor.DarkBlue : Console.Clear()
        DrawBox(1, 1, 97, 26, " MONITOR DE RECURSOS DETALHADO ")
        DrawBox(3, 4, 45, 21, " PROCESSAMENTO (CORES) ")

        ' Moldura de RAM
        DrawBox(49, 4, 47, 8, " MEMORIA RAM ")

        ' Moldura de Vídeo (Y=12, Altura 5)
        DrawBox(49, 12, 47, 5, " VIDEO / GPU ")

        ' CORREÇÃO: Moldura de SLOTS (Y=17, Altura 8 para garantir as 4 linhas com folga)
        DrawBox(49, 17, 47, 8, " SLOTS FISICOS (RAM) ")

        ' Nome do Processador
        Console.SetCursorPosition(4, 3) : Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write("CPU: " & Truncate(GetWMI("Win32_Processor", "Name"), 80))

        ' Info de GPU
        Dim gpuName = GetWMI("Win32_VideoController", "Name")
        Dim vram = GetVRAM()
        Console.SetCursorPosition(51, 14) : Console.ForegroundColor = ConsoleColor.White
        Console.Write("GPU: " & Truncate(gpuName, 35))
        Console.SetCursorPosition(51, 15) : Console.Write("VRAM: " & vram)

        ' CORREÇÃO: Preenchimento de Slots - Agora garantindo espaço para 4
        Dim sy = 19
        Try
            Using s As New ManagementObjectSearcher("SELECT Capacity, Speed, DeviceLocator FROM Win32_PhysicalMemory")
                For Each obj In s.Get()
                    If sy <= 23 Then ' sy vai de 19, 20, 21, 22... cabe 4 ou 5 slots com folga
                        Console.SetCursorPosition(51, sy) : Console.ForegroundColor = ConsoleColor.Gray
                        Dim cap = (Val(obj("Capacity")) / 1024 ^ 3).ToString("F0")
                        Console.Write($"• {obj("DeviceLocator")}: {cap}GB @ {obj("Speed")}MHz")
                        sy += 1
                    End If
                Next
            End Using
        Catch : End Try

        DrawMenuFooter("  ESC Voltar ao Menu Principal ")

        While Not (Console.KeyAvailable AndAlso Console.ReadKey(True).Key = ConsoleKey.Escape)
            ' Atualização dos Cores
            For i = 0 To Math.Min(coreCounters.Count - 1, 15)
                Console.SetCursorPosition(5, 6 + i)
                Dim val = coreCounters(i).NextValue()
                Console.ForegroundColor = ConsoleColor.White : Console.Write($"C{i:00}: ")
                DrawSimpleBar(val, 100, ConsoleColor.Cyan, 15) : Console.Write($" {val:F0}%  ")
            Next

            ' RAM Dinâmica
            Dim ramT = My.Computer.Info.TotalPhysicalMemory / (1024 ^ 3)
            Dim ramL = My.Computer.Info.AvailablePhysicalMemory / (1024 ^ 3)
            Dim ramU = ramT - ramL

            Console.SetCursorPosition(51, 6) : Console.ForegroundColor = ConsoleColor.White
            Console.Write("USO: ") : DrawSimpleBar(ramU, ramT, ConsoleColor.Yellow, 20)

            Console.SetCursorPosition(51, 8) : Console.Write($"TOTAL: {ramT:F1} GB".PadRight(20))
            Console.SetCursorPosition(51, 9) : Console.Write($"USADA: {ramU:F1} GB".PadRight(20))
            Console.SetCursorPosition(51, 10) : Console.Write($"LIVRE: {ramL:F1} GB".PadRight(20))

            Threading.Thread.Sleep(700)
        End While
    End Sub

    ' --- MÉTODOS DE APOIO ---
    Function GetActiveNetwork() As String
        Try
            For Each ni In NetworkInterface.GetAllNetworkInterfaces()
                If ni.OperationalStatus = OperationalStatus.Up AndAlso Not ni.Description.ToLower().Contains("virtual") Then
                    Dim props = ni.GetIPProperties()
                    Dim ip = props.UnicastAddresses.FirstOrDefault(Function(x) x.Address.AddressFamily = AddressFamily.InterNetwork)
                    If ip IsNot Nothing Then Return $"{Truncate(ni.Description, 25)} [{ip.Address}]"
                End If
            Next
        Catch : End Try
        Return "Sem Conexão"
    End Function

    ' [As subs DrawWindowsLogo, ShowPnPDevices, DrawBox, DrawBoxColor, DrawSimpleBar, GetWMI, GetVRAM, InitCounters, DrawMenuFooter e Truncate permanecem iguais]
    ' Garante o funcionamento do Pop-up e Gráficos de barra.

    Sub ShowPnPDevices()
        Dim devs As New List(Of ManagementObject)
        Using s As New ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity")
            For Each d In s.Get() : devs.Add(CType(d, ManagementObject)) : Next
        End Using
        Dim idx = 0
        While True
            Console.BackgroundColor = ConsoleColor.DarkBlue : Console.Clear()
            DrawBox(2, 1, 95, 25, " GERENCIADOR DE HARDWARE ")
            For i = 0 To 19
                If idx + i < devs.Count Then
                    Console.SetCursorPosition(4, 3 + i)
                    Dim name = If(devs(idx + i)("Name") IsNot Nothing, devs(idx + i)("Name").ToString(), "Desconhecido")
                    If i = 0 Then
                        Console.BackgroundColor = ConsoleColor.Cyan : Console.ForegroundColor = ConsoleColor.Black
                        Console.Write("> " & Truncate(name, 85).PadRight(88))
                        Console.BackgroundColor = ConsoleColor.DarkBlue
                    Else
                        Console.ForegroundColor = ConsoleColor.White : Console.Write("  " & Truncate(name, 85).PadRight(88))
                    End If
                End If
            Next
            DrawMenuFooter(" [SETAS] Navegar | [ENTER] Detalhes | [ESC] Sair ")
            Dim k = Console.ReadKey(True)
            If k.Key = ConsoleKey.Escape Then Exit While
            If k.Key = ConsoleKey.DownArrow And idx < devs.Count - 1 Then idx += 1
            If k.Key = ConsoleKey.UpArrow And idx > 0 Then idx -= 1
            If k.Key = ConsoleKey.Enter Then
                Dim d = devs(idx)
                DrawBoxColor(15, 6, 70, 14, " PROPRIEDADES DO HARDWARE ", ConsoleColor.Gray, ConsoleColor.Black)
                Console.ForegroundColor = ConsoleColor.Black : Console.BackgroundColor = ConsoleColor.Gray
                Console.SetCursorPosition(18, 9) : Console.Write("NOME: " & Truncate(d("Name"), 60))
                Console.SetCursorPosition(18, 10) : Console.Write("FABRICANTE: " & d("Manufacturer"))
                Console.SetCursorPosition(18, 11) : Console.Write("ID PNP: " & Truncate(d("PNPDeviceID"), 55))
                Console.SetCursorPosition(18, 12) : Console.Write("STATUS: " & d("Status"))
                Console.SetCursorPosition(15 + 18, 18) : Console.ForegroundColor = ConsoleColor.DarkRed : Console.Write(" Pressione qualquer tecla ")
                Console.ReadKey(True)
            End If
        End While
    End Sub

    Sub DrawWindowsLogo(x As Integer, y As Integer)
        Dim b = ChrW(&H2588)
        Console.ForegroundColor = ConsoleColor.Cyan
        For i = 0 To 3
            Console.SetCursorPosition(x, y + i) : Console.Write(New String(b, 7) & " " & New String(b, 7))
        Next
        For i = 5 To 8
            Console.SetCursorPosition(x, y + i) : Console.Write(New String(b, 7) & " " & New String(b, 7))
        Next
    End Sub

    Sub DrawBoxColor(x As Integer, y As Integer, w As Integer, h As Integer, title As String, bg As ConsoleColor, fg As ConsoleColor)
        Console.BackgroundColor = bg : Console.ForegroundColor = fg
        Console.SetCursorPosition(x, y) : Console.Write("╔" & New String("═"c, w - 2) & "╗")
        For i = 1 To h - 2 : Console.SetCursorPosition(x, y + i) : Console.Write("║" & New String(" "c, w - 2) & "║") : Next
        Console.SetCursorPosition(x, y + h - 1) : Console.Write("╚" & New String("═"c, w - 2) & "╝")
        If title <> "" Then
            Console.SetCursorPosition(x + (w \ 2) - (title.Length \ 2), y) : Console.Write(title)
        End If
    End Sub

    Sub DrawBox(x As Integer, y As Integer, w As Integer, h As Integer, title As String)
        DrawBoxColor(x, y, w, h, title, ConsoleColor.DarkBlue, ConsoleColor.White)
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

    Sub DrawMenuFooter(text As String)
        Console.SetCursorPosition(0, 29)
        Console.BackgroundColor = ConsoleColor.Gray : Console.ForegroundColor = ConsoleColor.Black
        Console.Write(text.PadRight(100))
        Console.BackgroundColor = ConsoleColor.DarkBlue
    End Sub

    Function Truncate(value As Object, max As Integer) As String
        If value Is Nothing Then Return "N/A"
        Dim s = value.ToString()
        Return If(s.Length <= max, s, s.Substring(0, max - 3) & "...")
    End Function
End Module