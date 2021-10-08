Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.IO.Ports
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Namespace DavidFiscal
    <ComVisible(True)> _
    Public NotInheritable Class David

#Region "Fields"
        Private _ComPortNumber As Integer
        Private _Items As New List(Of Article)
        Private AppPath As String '= AppConfig.AppLocal & "DAVID.EXE"
        Private INIPath As String
        Private TextFile As String

        Public Shared LastCommand As Integer
#End Region
#Region "Property"
        Property ComPortNumber() As Integer
            Get
                Return _ComPortNumber
            End Get
            Set(ByVal value As Integer)
                _ComPortNumber = value
            End Set
        End Property
        Property Stavki() As List(Of Article)
            Get
                Return _Items
            End Get
            Set(ByVal value As List(Of Article))
                _Items = New List(Of Article)
                _Items = value
            End Set
        End Property
#End Region
#Region "Enums"
        Public Enum PaidMode
            SoKredit = 3
            SoKarticka = 2
            VoGotovo = 1
        End Enum
#End Region
#Region "Ctor"
        Private Sub New()
            If IO.Directory.Exists(AppConfig.AppLocal) = False Then
                IO.Directory.CreateDirectory(AppConfig.AppLocal)
            End If
            Me.AppPath = AppConfig.AppLocal & "DAVID32.EXE"
            Me.INIPath = AppConfig.AppLocal & "FISKAL.INI"
            Me.TextFile = AppConfig.AppLocal & "smetka.txt"
            _ComPortNumber = 1
            Me._Items = New List(Of Article)
            
            David.LastCommand = 0
        End Sub
        Sub New(ByVal ComPort As Integer)
            Me.New()
            Me._ComPortNumber = ComPort
            Me.CreateIniFile()
            If IO.File.Exists(AppPath) = False Then
                'Ok nuk po ekzistojka duhet me formu ket.
                Me.CreateExecutable()
            End If
        End Sub
        Sub New(ByVal ComPort As Integer, ByVal stavki As List(Of Article))
            Me.New(ComPort)
            'Me._ComPortNumber = ComPort
            Me.Stavki = stavki
            'Me.CreateIniFile()
            'If IO.File.Exists(AppPath) = False Then
            '    'Ok nuk po ekzistojka duhet me formu ket.
            '    Me.CreateExecutable()
            'End If
        End Sub
#End Region
#Region "Public Methods"
        <ComVisible(True)> _
        Public Sub FiskalnaSmetka(Optional ByVal PaidType As PaidMode = PaidMode.VoGotovo)
            If Me.Stavki.Count = 0 Then
                Return
            End If
            Me.CreateFiskalnaPF550(PaidType)
            'PF550.LastCommand += 1
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub FiskalnaSmetka(ByVal Stavki As List(Of Article), Optional ByVal PaidType As PaidMode = PaidMode.VoGotovo)
            Me.Stavki = Stavki
            Me.FiskalnaSmetka(PaidType)
        End Sub
        <ComVisible(True)> _
        Public Sub StornaSmetka(Optional ByVal PaidType As PaidMode = PaidMode.VoGotovo)
            If Me.Stavki.Count = 0 Then
                Return
            End If
            Me.CreateStornaPF550(PaidType)
            'PF550.LastCommand += 1
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub StornaSmetka(ByVal Stavki As List(Of Article), Optional ByVal PaidType As PaidMode = PaidMode.VoGotovo)
            Me.Stavki = Stavki
            Me.StornaSmetka(PaidType)
        End Sub
        <ComVisible(True)> _
        Public Sub SluzbenVnes(ByVal Amount As Decimal)
            Using sw As StreamWriter = File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If David.LastCommand = 1 Then
                    seq = Chr(35)
                    David.LastCommand = 2
                Else
                    seq = Chr(36)
                    David.LastCommand = 1
                End If
                sw.Write(seq & Chr(70) & Amount & Chr(13) & Chr(10))
                sw.Close()
            End Using
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub PaperFeed(Optional ByVal Lines As Integer = 2)
            Using sw As StreamWriter = File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If David.LastCommand = 1 Then
                    seq = Chr(35)
                    David.LastCommand = 2
                Else
                    seq = Chr(36)
                    David.LastCommand = 1
                End If
                sw.Write(seq + Chr(44) + Lines.ToString + Chr(13) + Chr(10))
                sw.Close()
            End Using
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub PodesuvajCas()
            Using sw As StreamWriter = File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If David.LastCommand = 1 Then
                    seq = Chr(35)
                    David.LastCommand = 2
                Else
                    seq = Chr(36)
                    David.LastCommand = 1
                End If
                sw.Write(seq + Chr(61) + Date.Now.ToString("dd-MM-yy HH:MM:ss") + Chr(13) + Chr(10))
                'sw.Flush()
                sw.Close()
            End Using
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub ZatvoriDen()
            Using sw As StreamWriter = File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If David.LastCommand = 1 Then
                    seq = Chr(35)
                    David.LastCommand = 2
                Else
                    seq = Chr(36)
                    David.LastCommand = 1
                End If
                sw.Write(seq + Chr(69) + "1" + Chr(13) + Chr(10))
                sw.Close()
            End Using
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub Izvestaj_X()
            Using sw As StreamWriter = File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If David.LastCommand = 1 Then
                    seq = Chr(35)
                    David.LastCommand = 2
                Else
                    seq = Chr(36)
                    David.LastCommand = 1
                End If
                sw.Write(seq & Chr(69) + "3" + Chr(13) + Chr(10))
                sw.Close()
            End Using
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub PeriodicenIzvestaj(ByVal OdDatum As DateTime, ByVal DoDatum As DateTime)
            Using sw As StreamWriter = File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If David.LastCommand = 1 Then
                    seq = Chr(35)
                    David.LastCommand = 2
                Else
                    seq = Chr(36)
                    David.LastCommand = 1
                End If
                sw.Write(seq + Chr(79) + OdDatum.ToString("ddMMyy") & "," & DoDatum.ToString("ddMMyy") + Chr(13) + Chr(10))
                sw.Close()
            End Using
            Me.Run()
        End Sub
#End Region
#Region "Private Methods"
        Private Sub CreateStornaPF550(Optional ByVal PaidMode As PaidMode = PaidMode.VoGotovo)
            Dim sw As StreamWriter = File.CreateText(TextFile)
            sw.Write(Chr(32) + Chr(48) + "1,0001,1,R" + Chr(13) + Chr(10))
            Dim m As Integer = 0
            For Each tmp As Article In Me.Stavki
                Dim d As Char = Chr(tmp.VAT)
                'Select Case tmp.VAT
                '    Case VATgroup.C
                '        d = Chr(68)
                '    Case VATgroup.B
                '        d = Chr(66)
                '    Case VATgroup.A
                '        d = Chr(65)
                '    Case Else
                '        d = Chr(68)
                'End Select
                If m Mod 2 = 0 Then
                    sw.Write(Chr(35) & Chr(49) & tmp.Name & Chr(9) & d & tmp.Price.ToString("##.00") & "*" & tmp.Amount.ToString("#0.000") & Chr(13) & Chr(10))
                Else
                    sw.Write(Chr(36) & Chr(49) & tmp.Name & Chr(9) & d & tmp.Price.ToString("##.00") & "*" & tmp.Amount.ToString("#0.000") & Chr(13) & Chr(10))
                End If
                m += 1
            Next
            'sw.Write(Chr(37) & Chr(52) & Chr(9) & Chr(13) & Chr(10))
            Dim Paid As Char = "P"c
            Select Case PaidMode
                Case David.PaidMode.VoGotovo
                    Paid = "P"c
                Case David.PaidMode.SoKarticka
                    Paid = "C"c
                Case David.PaidMode.SoKredit
                    Paid = "D"c
            End Select

            sw.Write(Chr(32) & Chr(53) & Chr(9) & Paid & Chr(13) & Chr(10))
            sw.Write(Chr(37) & Chr(56))
            sw.Flush()
            sw.Close()
        End Sub
        Private Sub CreateFiskalnaPF550(Optional ByVal PaidMode As PaidMode = PaidMode.VoGotovo)
            Dim sw As StreamWriter = File.CreateText(TextFile)
            sw.Write(Chr(32) + Chr(48) + "1,0001,1" + Chr(13) + Chr(10))
            Dim m As Integer = 0
            For Each tmp As Article In Me.Stavki
                Dim d As Char = Chr(tmp.VAT)
                'Select Case tmp.VAT
                '    Case VATgroup.C
                '        d = Chr(68)
                '    Case VATgroup.B
                '        d = Chr(66)
                '    Case VATgroup.A
                '        d = Chr(65)
                '    Case Else
                '        d = Chr(68)
                'End Select
                If m Mod 2 = 0 Then
                    sw.Write(Chr(35) & Chr(49) & tmp.Name & Chr(9) & d & tmp.Price.ToString("##.00") & "*" & tmp.Amount.ToString("#0.000") & Chr(13) & Chr(10))
                Else
                    sw.Write(Chr(36) & Chr(49) & tmp.Name & Chr(9) & d & tmp.Price.ToString("##.00") & "*" & tmp.Amount.ToString("#0.000") & Chr(13) & Chr(10))
                End If
                m += 1
            Next
            'sw.Write(Chr(37) & Chr(52) & Chr(9) & Chr(13) & Chr(10))
            Dim Paid As Char = "P"c
            Select Case PaidMode
                Case David.PaidMode.VoGotovo
                    Paid = "P"c
                Case David.PaidMode.SoKarticka
                    Paid = "C"c
                Case David.PaidMode.SoKredit
                    Paid = "D"c
            End Select

            sw.Write(Chr(32) & Chr(53) & Chr(9) & Paid & Chr(13) & Chr(10))
            sw.Write(Chr(37) & Chr(56))
            sw.Flush()
            sw.Close()
        End Sub
        Private Sub CreateExecutable()
            Using createFile As New FileStream(Me.AppPath, FileMode.CreateNew, FileAccess.Write)
                createFile.Write(My.Resources.DAVID32, 0, My.Resources.DAVID32.Length)
            End Using
        End Sub
        Private Sub CreateIniFile()
            If IO.File.Exists(INIPath) = True Then
                IO.File.Delete(INIPath)
            End If
            Using sw As StreamWriter = File.CreateText(INIPath)
                'sw.WriteLine(";Ini fajlot treba da bide vo ist dir. so fiscal32.exe")
                'sw.WriteLine("")
                sw.WriteLine("[Setup]")
                sw.WriteLine("Port=COM" & Me.ComPortNumber)
                sw.WriteLine("Speed=5")
                sw.WriteLine("Bit=8")
                sw.WriteLine("Parity=0")
                sw.WriteLine("Stop=1")
                sw.WriteLine("Flow=0")
                sw.Close()
            End Using
        End Sub
        Private Sub Run()
            Shell(String.Format("{0} {1}", AppPath, TextFile), AppWinStyle.Hide)
        End Sub
        Public Function FormatNumber(ByVal Number As Decimal, ByVal DecimalDigits As Integer) As String
            If (Decimal.Compare(Number, Decimal.Zero) = 0) Then
                Dim str2 As String = "0."
                Dim num3 As Integer = DecimalDigits
                Dim i As Integer = 1
                Do While (i <= num3)
                    str2 = (str2 & "0")
                    i += 1
                Loop
                Return str2
            End If
            Dim num As Long = CLng(Math.Round(CDbl((Convert.ToDouble(Number) * Math.Pow(10, CDbl(DecimalDigits))))))
            Return (Strings.Left(Conversions.ToString(num), (Strings.Len(Conversions.ToString(num)) - DecimalDigits)) & "," & Strings.Right(Conversions.ToString(num), DecimalDigits))
        End Function
#End Region
    End Class
End Namespace