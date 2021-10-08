Imports System.Net
Imports System.Reflection
Imports System.Net.Sockets

Public Class AppConfig
    Private Shared _Application As AssemblyName = Nothing
    Private Shared _IP As IPAddress

    Public Shared ReadOnly Property Application() As AssemblyName
        Get
            If (AppConfig._Application Is Nothing) Then
                Dim entryAssembly As Assembly = Assembly.GetEntryAssembly
                If (Not entryAssembly Is Nothing) Then
                    AppConfig._Application = entryAssembly.GetName
                Else
                    AppConfig._Application = New AssemblyName
                End If
            End If
            Return AppConfig._Application
        End Get
    End Property
    Public Shared ReadOnly Property ApplicationName() As String
        Get
            Return AppConfig.Application.Name
        End Get
    End Property
    Public Shared ReadOnly Property ApplicationPath() As String
        Get
            Return My.Application.Info.DirectoryPath & "\"
        End Get
    End Property
    Public Shared ReadOnly Property AppLocal() As String
        Get
            Return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\" & AppConfig.ApplicationName & "\"
        End Get
    End Property
    Public Shared ReadOnly Property IP() As IPAddress
        Get
            If (AppConfig._IP Is Nothing) Then
                Dim hostEntry As IPHostEntry = Dns.GetHostEntry(Dns.GetHostName)
                AppConfig._IP = New IPAddress(0)
                Dim address As IPAddress
                For Each address In hostEntry.AddressList
                    If (address.AddressFamily = AddressFamily.InterNetwork) Then
                        AppConfig._IP = address
                    End If
                Next
            End If
            Return AppConfig._IP
        End Get
    End Property
End Class
