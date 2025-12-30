Imports System.Runtime.Serialization

<DataContract>
Public Class BackupMetadata
    <DataMember>
    Public Property BackupDate As DateTime

    <DataMember>
    Public Property ComputerName As String

    <DataMember>
    Public Property OS As String

    <DataMember>
    Public Property Drivers As List(Of DriverInfo)
End Class

<DataContract>
Public Class DriverInfo
    <DataMember>
    Public Property DeviceName As String

    <DataMember>
    Public Property Manufacturer As String

    <DataMember>
    Public Property Version As String

    <DataMember>
    Public Property InfName As String

    <DataMember>
    Public Property DeviceID As String

    <DataMember>
    Public Property ClassName As String

    <DataMember>
    Public Property Folder As String
End Class
