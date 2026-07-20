Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Runtime.InteropServices

''' <summary>Uma linha de dados de processo já pronta para exibição.</summary>
Public Structure ProcessRow
    Public PID As Integer
    Public Name As String
    Public CpuPercent As Double
    Public MemoryMB As Double
    Public DiskKBs As Double
    Public Priority As String
    Public FilePath As String
End Structure

''' <summary>
''' Coleta dados de processos em execução de forma eficiente e sem ambiguidade
''' de instância (CPU e Disco calculados por PID via API do Windows, não por nome).
''' Deve ser chamado de uma thread de background (Task.Run); não toca em UI.
''' </summary>
Public Class ProcessSnapshotService

    Private Structure CpuSample
        Public CpuTime As TimeSpan
        Public Timestamp As DateTime
    End Structure

    Private Structure IoSample
        Public TotalBytes As ULong
        Public Timestamp As DateTime
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure IO_COUNTERS
        Public ReadOperationCount As ULong
        Public WriteOperationCount As ULong
        Public OtherOperationCount As ULong
        Public ReadTransferCount As ULong
        Public WriteTransferCount As ULong
        Public OtherTransferCount As ULong
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure MEMORYSTATUSEX
        Public dwLength As UInteger
        Public dwMemoryLoad As UInteger
        Public ullTotalPhys As ULong
        Public ullAvailPhys As ULong
        Public ullTotalPageFile As ULong
        Public ullAvailPageFile As ULong
        Public ullTotalVirtual As ULong
        Public ullAvailVirtual As ULong
        Public ullAvailExtendedVirtual As ULong
    End Structure

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function GetProcessIoCounters(hProcess As IntPtr, ByRef counters As IO_COUNTERS) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function GlobalMemoryStatusEx(ByRef lpBuffer As MEMORYSTATUSEX) As Boolean
    End Function

    Private cpuSamples As New Dictionary(Of Integer, CpuSample)
    Private ioSamples As New Dictionary(Of Integer, IoSample)

    ''' <summary>Monta a lista atual de processos com CPU%, memória e I/O de disco calculados por delta.</summary>
    Public Function GetSnapshot() As List(Of ProcessRow)
        Dim result As New List(Of ProcessRow)
        Dim now = DateTime.UtcNow
        Dim seenPids As New HashSet(Of Integer)

        For Each p As Process In Process.GetProcesses()
            Try
                Dim pid = p.Id
                seenPids.Add(pid)

                Dim cpuPercent = CalculateCpuPercent(p, pid, now)
                Dim memMB = SafeMemory(p)
                Dim diskKBs = CalculateDiskKBs(p, pid, now)
                Dim priorityStr = SafePriority(p)
                Dim filePath = SafeFilePath(p)

                result.Add(New ProcessRow With {
                    .PID = pid,
                    .Name = p.ProcessName,
                    .CpuPercent = cpuPercent,
                    .MemoryMB = memMB,
                    .DiskKBs = diskKBs,
                    .Priority = priorityStr,
                    .FilePath = filePath
                })
            Catch
                ' Processo terminou durante a leitura, ou sem acesso - ignora essa linha.
            End Try
        Next

        ' Evita crescimento infinito: remove amostras de PIDs que não existem mais.
        PruneDictionary(cpuSamples, seenPids)
        PruneDictionary(ioSamples, seenPids)

        Return result
    End Function

    ''' <summary>Memória física total/usada do sistema (não soma de processos, evita contagem duplicada de memória compartilhada).</summary>
    Public Function GetSystemMemoryInfo() As (UsedGB As Double, TotalGB As Double, Percent As Double)
        Dim status As New MEMORYSTATUSEX()
        status.dwLength = CUInt(Marshal.SizeOf(status))
        If GlobalMemoryStatusEx(status) Then
            Dim totalGB = status.ullTotalPhys / 1024.0 / 1024.0 / 1024.0
            Dim availGB = status.ullAvailPhys / 1024.0 / 1024.0 / 1024.0
            Return (totalGB - availGB, totalGB, status.dwMemoryLoad)
        End If
        Return (0, 0, 0)
    End Function

    ' ================= internos =================

    Private Function CalculateCpuPercent(p As Process, pid As Integer, now As DateTime) As Double
        Dim cpuPercent As Double = 0
        Try
            Dim currentCpuTime = p.TotalProcessorTime
            If cpuSamples.ContainsKey(pid) Then
                Dim prev = cpuSamples(pid)
                Dim elapsedMs = (now - prev.Timestamp).TotalMilliseconds
                If elapsedMs > 0 Then
                    Dim deltaCpuMs = (currentCpuTime - prev.CpuTime).TotalMilliseconds
                    cpuPercent = Math.Max(0, deltaCpuMs / (elapsedMs * Environment.ProcessorCount) * 100)
                End If
            End If
            cpuSamples(pid) = New CpuSample With {.CpuTime = currentCpuTime, .Timestamp = now}
        Catch
            ' Processo protegido do sistema: sem acesso ao tempo de CPU. Mantém 0.
        End Try
        Return cpuPercent
    End Function

    Private Function CalculateDiskKBs(p As Process, pid As Integer, now As DateTime) As Double
        Dim diskKBs As Double = 0
        Try
            Dim counters As New IO_COUNTERS()
            If GetProcessIoCounters(p.Handle, counters) Then
                Dim totalBytes = counters.ReadTransferCount + counters.WriteTransferCount
                If ioSamples.ContainsKey(pid) Then
                    Dim prevIo = ioSamples(pid)
                    Dim elapsedSec = (now - prevIo.Timestamp).TotalSeconds
                    If elapsedSec > 0 Then
                        Dim deltaBytes = CLng(totalBytes) - CLng(prevIo.TotalBytes)
                        diskKBs = Math.Max(0, deltaBytes / elapsedSec / 1024.0)
                    End If
                End If
                ioSamples(pid) = New IoSample With {.TotalBytes = totalBytes, .Timestamp = now}
            End If
        Catch
            ' Sem permissão para abrir o processo (comum em processos protegidos do sistema).
        End Try
        Return diskKBs
    End Function

    Private Function SafeMemory(p As Process) As Double
        Try
            Return p.WorkingSet64 / 1024.0 / 1024.0
        Catch
            Return 0
        End Try
    End Function

    Private Function SafePriority(p As Process) As String
        Try
            Select Case p.PriorityClass
                Case ProcessPriorityClass.Idle : Return "Baixa"
                Case ProcessPriorityClass.BelowNormal : Return "Abaixo do Normal"
                Case ProcessPriorityClass.Normal : Return "Normal"
                Case ProcessPriorityClass.AboveNormal : Return "Acima do Normal"
                Case ProcessPriorityClass.High : Return "Alta"
                Case ProcessPriorityClass.RealTime : Return "Tempo Real"
                Case Else : Return "N/D"
            End Select
        Catch
            Return "N/D"
        End Try
    End Function

    Private Function SafeFilePath(p As Process) As String
        Try
            Return p.MainModule.FileName
        Catch
            Return ""
        End Try
    End Function

    Private Sub PruneDictionary(Of T)(dict As Dictionary(Of Integer, T), seenPids As HashSet(Of Integer))
        For Each oldPid In dict.Keys.Where(Function(k) Not seenPids.Contains(k)).ToList()
            dict.Remove(oldPid)
        Next
    End Sub

End Class
