Imports System.Runtime.InteropServices
Imports System
Imports System.Collections.Generic
Imports System.IO.Ports
Imports System.Threading
Namespace Accent
    Friend Class Synergy
        Private Const NAK As Byte = 21
        Private Const SYN As Byte = 22
        Private Const VERSION As String = "1.0.1"
        Private serialPort As String = "COM1"
        Private baudRate As Integer = 9600
        Private port As SerialPort
        Private Shared seq As Byte = 33
        Private Shared isFirstCommandExecuted As Boolean = False

        Public Sub New(ByVal serialPort As String, ByVal baudRate As Integer)
            Me.serialPort = serialPort
            Me.baudRate = baudRate
            Me.port = New SerialPort(serialPort, baudRate, Parity.None, 8, StopBits.One)
            Me.port.ReadTimeout = 100
            Me.port.WriteTimeout = 100
        End Sub
        Public Function WriteCommand(ByVal command As Byte, ByVal data As String) As PrinterResult
            If Not Me.port.IsOpen Then
                Throw New PortNotOpenedException()
            End If
            If Not Synergy.isFirstCommandExecuted Then
                Synergy.isFirstCommandExecuted = True
                Me.SendReceive(74, "")
            End If
            Return Me.SendReceive(command, data)
        End Function
        Private Function SendReceive(ByVal command As Byte, ByVal data As String) As PrinterResult
            Dim printerResult As PrinterResult = Nothing
            Dim num As Integer = 0
            Dim flag As Boolean = False
            Dim newSequenceCode As Boolean = True
            While Not flag
                printerResult = Me.SendCommand(Me.getSequenceCode(newSequenceCode), command, data)
                Dim flag2 As Boolean = printerResult.ResultStatus = PrinterResultStatus.NAK_RECEIVED OrElse printerResult.ResultStatus = PrinterResultStatus.TIMEOUT_READING OrElse printerResult.ResultStatus = PrinterResultStatus.WRONG_COMMAND_RESPONSE
                If flag2 Then
                    newSequenceCode = False
                    num += 1
                    flag = (num > 2)
                Else
                    flag = True
                End If
            End While
            Return printerResult
        End Function
        Private Function SendCommand(ByVal seq As Byte, ByVal command As Byte, ByVal data As String) As PrinterResult
            Dim num As Integer = 36
            If data Is Nothing Then
                data = ""
            End If
            num += data.Length
            Dim array As Byte() = New Byte(10 + data.Length - 1) {}
            array(0) = 1
            array(1) = CByte((4 + data.Length + 32))
            array(2) = seq
            array(3) = command
            num += CInt(seq)
            num += CInt(command)
            For i As Integer = 0 To data.Length - 1
                'array(4 + i) = CByte(data(i))
                'num += CInt((CByte(data(i))))
                array(4 + i) = CByte(AscW(data(i)))
                num += CInt(CByte(AscW(data(i))))
            Next
            num += 5
            array(4 + data.Length) = 5
            array(4 + data.Length + 1) = CByte((num / 256 / 16 + 48))
            array(4 + data.Length + 2) = CByte((num / 256 Mod 16 + 48))
            array(4 + data.Length + 3) = CByte((num Mod 256 / 16 + 48))
            array(4 + data.Length + 4) = CByte((num Mod 256 Mod 16 + 48))
            array(4 + data.Length + 5) = 3
            Me.WriteToPort(array)
            Return Me.ReadFromPort(command)
        End Function
        Private Sub WriteToPort(ByVal bytes As Byte())
            Me.port.Write(bytes, 0, bytes.Length)
        End Sub
        Private Function ReadFromPort(ByVal command As Byte) As PrinterResult
            Dim printerResult As PrinterResult = New PrinterResult()
            Dim flag As Boolean = False
            Dim num As Integer = 0
            Dim flag2 As Boolean = False
            Dim num2 As Integer = 0
            Dim b As Byte = 0
            While Not flag
                Try
                    b = CByte(Me.port.ReadByte())
                Catch 'ex_27 As Object
                    b = 0
                End Try
                If b > 0 Then
                    num = 0
                    Dim b2 As Byte = b
                    If b2 <> 1 Then
                        Select Case b2
                            Case 21
                                flag = True
                                printerResult.ResultStatus = PrinterResultStatus.NAK_RECEIVED
                                Return printerResult
                            Case 22
                                Try
                                    Thread.Sleep(60)
                                    Continue While
                                Catch 'ex_6B As Object
                                    Continue While
                                End Try
                            Case Else
                                If flag2 Then
                                    If printerResult.Response.Count >= 4 AndAlso printerResult.Response(3) <> CInt(command) Then
                                        flag = True
                                        printerResult.Response = Nothing
                                        printerResult.ResultStatus = PrinterResultStatus.WRONG_COMMAND_RESPONSE
                                        Return printerResult
                                    End If
                                    printerResult.Response.Add(CInt(b))
                                    If b <> 3 Then
                                        Continue While
                                    End If
                                    flag = True
                                    printerResult.ResultStatus = PrinterResultStatus.OK
                                    Try
                                        printerResult.Data = New Byte(printerResult.Response.Count - 17 - 1) {}
                                        For i As Integer = 0 To printerResult.Response.Count - 17 - 1
                                            printerResult.Data(i) = CByte(printerResult.Response(i + 4))
                                        Next
                                        printerResult.Status = New Byte(6 - 1) {}
                                        For j As Integer = 0 To 6 - 1
                                            printerResult.Status(j) = CByte(printerResult.Response(j + (printerResult.Response.Count - 17) + 5))
                                        Next
                                        If Me.IsBitSet(printerResult.Status(0), 0) Then
                                            printerResult.ResultStatus = PrinterResultStatus.SYNTAX_ERROR
                                            Dim result As PrinterResult = printerResult
                                            Return result
                                        End If
                                        Continue While
                                    Catch 'ex_197 As Object
                                        printerResult.ResultStatus = PrinterResultStatus.INVALID_RESPONSE
                                        Dim result As PrinterResult = printerResult
                                        Return result
                                    End Try
                                End If
                                num2 += 1
                                If num2 >= 1000 Then
                                    flag = True
                                    printerResult.ResultStatus = PrinterResultStatus.INVALID_RESPONSE
                                    printerResult.Response = Nothing
                                    Return printerResult
                                End If
                                Continue While
                        End Select
                    End If
                    num2 = 0
                    flag2 = True
                    printerResult.Response = New List(Of Integer)()
                    printerResult.Response.Add(1)
                Else
                    Try
                        Thread.Sleep(50)
                    Catch ex_1CE As Exception
                    End Try
                    num += 1
                    If num >= 3 Then
                        printerResult.ResultStatus = PrinterResultStatus.TIMEOUT_READING
                        printerResult.Response = Nothing
                        Return printerResult
                    End If
                End If
            End While
            Return printerResult
        End Function
        Private Function getSequenceCode(ByVal newSequenceCode As Boolean) As Byte
            If newSequenceCode Then
                Synergy.seq += 1
            End If
            If Synergy.seq > 127 Then
                Synergy.seq = 33
            End If
            Return Synergy.seq
        End Function
        Public Sub OpenPort()
            Me.port.Open()
        End Sub
        Public Sub ClosePort()
            Me.port.Close()
        End Sub
        Private Function IsBitSet(ByVal b As Byte, ByVal pos As Integer) As Boolean
            Return (CInt(b) And 1 << pos) <> 0
        End Function
        Public Function GetVersion() As String
            Return "1.0.1"
        End Function
    End Class
    Friend Class PrinterResult
        Public ResultStatus As PrinterResultStatus
        Public Response As List(Of Integer)
        Public Data As Byte()
        Public Status As Byte()
    End Class
    Friend Class PortNotOpenedException
        Inherits Exception

        Public Sub New()
        End Sub

        Public Sub New(ByVal message As String)
            MyBase.New(message)
            'MyBase.[New](message)
        End Sub

        Public Sub New(ByVal format As String, <Out()> ByVal ParamArray args As Object())
            MyBase.New(String.Format(format, args))
            'MyBase.[New](String.Format(format, args))
        End Sub

        Public Sub New(ByVal message As String, ByVal innerException As Exception)
            MyBase.New(message, innerException)
            'MyBase.[New](message, innerException)
        End Sub

        'Public Sub New(ByVal format As String, ByVal innerException As Exception, <Out()> ByVal ParamArray args As Object())
        '    MyBase.New(String.Format(format, args), args)
        '    'MyBase.[New](String.Format(format, args), innerException)
        'End Sub
    End Class
    Friend Enum PrinterResultStatus
        UNKNOWN
        OK
        NAK_RECEIVED
        TIMEOUT_READING
        WRONG_COMMAND_RESPONSE
        GENERAL_ERROR
        SYNTAX_ERROR
        INVALID_RESPONSE
    End Enum
End Namespace