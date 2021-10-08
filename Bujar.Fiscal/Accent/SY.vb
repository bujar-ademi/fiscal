Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.IO.Ports
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Text
Namespace Accent
    <ComVisible(True)> _
    Public NotInheritable Class SY
#Region "Fields"
        Private _ComPortNumber As Integer
        Private _Items As New List(Of Article)
        'Private s As Synergy
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
        Property Items() As List(Of Article)
            Get
                Return _Items
            End Get
            Set(ByVal value As List(Of Article))
                _Items = New List(Of Article)
                _Items = value
            End Set
        End Property

        Public Enum TypeIn
            CashIn = 0
            CashOut = 1
        End Enum
#End Region
#Region "Ctor"
        Private Sub New()
            ComPortNumber = 1
            Items = New List(Of Article)
            's = New Synergy(String.Format("COM{0}", ComPortNumber), 9600)
        End Sub
        Sub New(ByVal COM As Integer)
            Me.New()
            ComPortNumber = COM
            's = New Synergy(String.Format("COM{0}", ComPortNumber), 9600)
        End Sub
        Sub New(ByVal Com As Integer, ByVal itms As List(Of Article))
            Me.New(Com)
            Items = itms
        End Sub
#End Region
#Region "Methods"
        Private Sub Log(ByVal msg As String)
            'Dim wr As StreamWriter = New StreamWriter(path, True, Text.Encoding.Unicode)
            Dim pathFile As String = AppConfig.AppLocal & "Fiscal_" & Date.Today.ToString("ddMMyyyy") & ".log"
            Dim fileExists As Boolean
            fileExists = My.Computer.FileSystem.FileExists(pathFile)
            If fileExists = False Then
                My.Computer.FileSystem.WriteAllText(pathFile, String.Empty, False)
            End If
            Using wr As New StreamWriter(pathFile, True, Encoding.UTF8)
                wr.WriteLine(String.Format("{0} - {1}", Date.Now.ToString("dd.MM.yyyy HH:m:s"), msg))
                wr.Close()
            End Using
        End Sub
        <ComVisible(True)> _
        Public Sub FiskalnaSmetka(Optional ByVal PaidType As PaidMode = PaidMode.VoGotovo)
            If Me.Items.Count = 0 Then
                Return
            End If
            Dim s As Synergy = New Synergy(String.Format("COM{0}", ComPortNumber), 9600)
            s.OpenPort()
            Dim result As PrinterResult
            '" 01	1		0	"
            result = s.WriteCommand(48, "01" & Chr(9) & "0" & Chr(9) & "0")
            Log(result.ResultStatus & ", response bytes: " & result.Response.Count)
            Dim Iznos As Decimal = 0
            For Each tmp As Article In Items
                Dim d As Char = Chr(tmp.VAT)
                Dim VAT As Int16 = 1
                Select Case tmp.VAT
                    Case VATgroup.А
                        VAT = 1
                    Case VATgroup.Б
                        VAT = 2
                    Case VATgroup.В
                        VAT = 3
                    Case VATgroup.Г
                        VAT = 4
                End Select

                Dim data As String = tmp.Name.ToUpper & Chr(9) & VAT & Chr(9) & FormatNumber(tmp.Price, 2) & Chr(9) & FormatNumber(tmp.Amount, 3)
                result = s.WriteCommand(49, data)
                Log(result.ResultStatus & ", response bytes: " & result.Response.Count)
                'Iznos += CDec(tmp.Price * tmp.Amount)
            Next

            result = s.WriteCommand(53, PaidType & Chr(9))
            Log(result.ResultStatus & ", response bytes: " & result.Response.Count)
            result = s.WriteCommand(56, "")
            Log(result.ResultStatus & ", response bytes: " & result.Response.Count)
            s.ClosePort()
        End Sub
        <ComVisible(True)> _
        Public Sub FiskalnaSmetka(ByVal itms As List(Of Article), Optional ByVal PaidType As PaidMode = PaidMode.VoGotovo)
            Me.Items = itms
            Me.FiskalnaSmetka(PaidType)
        End Sub
        <ComVisible(True)> _
        Public Sub StornaSmetka(Optional ByVal PaidType As PaidMode = PaidMode.VoGotovo)
            If Me.Items.Count = 0 Then
                Return
            End If
            Dim s As Synergy = New Synergy(String.Format("COM{0}", ComPortNumber), 9600)
            s.OpenPort()
            '" 01	1		0	"
            Dim result As PrinterResult = s.WriteCommand(48, "01" & Chr(9) & "0" & Chr(9) & "1")
            Log(result.ResultStatus & ", response bytes: " & result.Response.Count)
            Dim Iznos As Decimal = 0
            For Each tmp As Article In Items
                Dim d As Char = Chr(tmp.VAT)
                Dim VAT As Int16 = 1
                Select Case tmp.VAT
                    Case VATgroup.А
                        VAT = 1
                    Case VATgroup.Б
                        VAT = 2
                    Case VATgroup.В
                        VAT = 3
                    Case VATgroup.Г
                        VAT = 4
                End Select

                Dim data As String = tmp.Name.ToUpper & Chr(9) & VAT & Chr(9) & FormatNumber(tmp.Price, 2) & Chr(9) & FormatNumber(tmp.Amount, 3)
                result = s.WriteCommand(49, data)
                Log(result.ResultStatus & ", response bytes: " & result.Response.Count)
                'Iznos += CDec(tmp.Price * tmp.Amount)
            Next

            result = s.WriteCommand(53, PaidType & Chr(9))
            Log(result.ResultStatus & ", response bytes: " & result.Response.Count)
            result = s.WriteCommand(56, "")
            Log(result.ResultStatus & ", response bytes: " & result.Response.Count)
            s.ClosePort()
        End Sub
        <ComVisible(True)> _
        Public Sub StornaSmetka(ByVal itms As List(Of Article), Optional ByVal PaidType As PaidMode = PaidMode.VoGotovo)
            Me.Items = itms
            Me.StornaSmetka(PaidType)
        End Sub
        <ComVisible(True)> _
        Public Sub PodesuvajCas()
            Dim s As Synergy = New Synergy(String.Format("COM{0}", ComPortNumber), 9600)
            s.OpenPort()
            Dim result As PrinterResult = s.WriteCommand(61, Date.Now.ToString("dd-MM-yy HH:MM:ss"))
            s.ClosePort()
        End Sub
        <ComVisible(True)> _
        Public Sub Izvestaj_X()
            Dim s As Synergy = New Synergy(String.Format("COM{0}", ComPortNumber), 9600)
            s.OpenPort()
            Dim result As PrinterResult = s.WriteCommand(69, "X" & Chr(9))
            Log(result.ResultStatus & ", response bytes: " & result.Response.Count)
            s.ClosePort()
        End Sub
        <ComVisible(True)> _
        Public Sub ZatvoriDen()
            Dim s As Synergy = New Synergy(String.Format("COM{0}", ComPortNumber), 9600)
            s.OpenPort()
            Dim result As PrinterResult = s.WriteCommand(69, "Z" & Chr(9))
            Log(result.ResultStatus & ", response bytes: " & result.Response.Count)
            s.ClosePort()
        End Sub
        <ComVisible(True)> _
        Public Sub PeriodicenIzvestaj(ByVal OdDatum As DateTime, ByVal DoDatum As DateTime)
            Dim s As Synergy = New Synergy(String.Format("COM{0}", ComPortNumber), 9600)
            s.OpenPort()
            Dim result As PrinterResult = s.WriteCommand(94, "0" & Chr(9) & OdDatum.ToString("dd-MM-yy") & Chr(9) & DoDatum.ToString("dd-MM-yy") & Chr(9))
            Log(result.ResultStatus & ", response bytes: " & result.Response.Count)
            s.ClosePort()
        End Sub
        <ComVisible(True)> _
        Public Sub SluzbenVnes(ByVal Amount As Decimal, ByVal Type As TypeIn)
            Dim s As Synergy = New Synergy(String.Format("COM{0}", ComPortNumber), 9600)
            s.OpenPort()
            Dim result As PrinterResult = s.WriteCommand(70, Type & Chr(9) & FormatNumber(Amount, 2) & Chr(9))
            Log(result.ResultStatus & ", response bytes: " & result.Response.Count)
            s.ClosePort()
        End Sub
#End Region
    End Class

    <ComVisible(True)> _
   <Obsolete("Not used anymore. Instead use SY class.", False)> _
   Public NotInheritable Class SY50
#Region "Fields"
        Private _ComPortNumber As Integer
        Private _Items As New List(Of Article)
        Private AppPath As String '= AppConfig.AppLocal & "DAVID.EXE"
        Private INIPath As String
        Private TextFile As String
        Public Shared LastCommand As Integer = 0
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
            SoKredit = 2
            SoKarticka = 1
            VoGotovo = 0
        End Enum
#End Region
#Region "Ctor"
        Private Sub New()
            If IO.Directory.Exists(AppConfig.AppLocal) = False Then
                IO.Directory.CreateDirectory(AppConfig.AppLocal)
            End If
            Me.AppPath = AppConfig.AppLocal & "sy.exe"
            Me.INIPath = AppConfig.AppLocal & "fiskal.ini"
            Me.TextFile = AppConfig.AppLocal & "PF500.in"
            _ComPortNumber = 1
            Me._Items = New List(Of Article)

            'SY50.LastCommand = 0
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
            Me.Stavki = stavki
        End Sub
#End Region
#Region "Public Methods"
        <ComVisible(True)> _
        Public Sub FiskalnaSmetka(Optional ByVal PaidType As PaidMode = PaidMode.VoGotovo)
            If Me.Stavki.Count = 0 Then
                Return
            End If
            Me.CreateFiskalnaSY50(PaidType)
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
            Me.CreateStornaSY50(PaidType)
            'PF550.LastCommand += 1
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub StornaSmetka(ByVal Stavki As List(Of Article), Optional ByVal PaidType As PaidMode = PaidMode.VoGotovo)
            Me.Stavki = Stavki
            Me.StornaSmetka(PaidType)
        End Sub
        <ComVisible(True)> _
        Public Sub PodesuvajCas()
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic) 'File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If SY50.LastCommand = 1 Then
                    seq = Chr(35)
                    SY50.LastCommand = 2
                Else
                    seq = Chr(36)
                    SY50.LastCommand = 1
                End If
                sw.Write(seq & String.Format("={0}	", Date.Now.ToString("dd-MM-yy HH:MM:ss")))
                'sw.Flush()
                sw.Close()
            End Using
            Me.Run()
        End Sub
        '<ComVisible(True)> _
        'Public Sub PaperFeed(Optional ByVal Lines As Integer = 2)
        '    Using sw As StreamWriter = File.CreateText(TextFile)
        '        Dim seq As Char = Chr(35)
        '        If SY50.LastCommand = 1 Then
        '            seq = Chr(35)
        '            SY50.LastCommand = 2
        '        Else
        '            seq = Chr(36)
        '            SY50.LastCommand = 1
        '        End If
        '        sw.Write(seq + Chr(44) + Lines.ToString + Chr(13) + Chr(10))
        '        sw.Close()
        '    End Using
        '    Me.Run()
        'End Sub
        <ComVisible(True)> _
        Public Sub ZatvoriDen()
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic) 'File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If SY50.LastCommand = 1 Then
                    seq = Chr(35)
                    SY50.LastCommand = 2
                Else
                    seq = Chr(36)
                    SY50.LastCommand = 1
                End If
                sw.Write(seq + "EZ	")
                sw.Close()
            End Using
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub Izvestaj_X()
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic) 'File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If SY50.LastCommand = 1 Then
                    seq = Chr(35)
                    SY50.LastCommand = 2
                Else
                    seq = Chr(36)
                    SY50.LastCommand = 1
                End If
                sw.Write(seq & "EX	")
                sw.Close()
            End Using
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub SluzbenIzlez(ByVal Amount As Decimal)
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic) 'File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If SY50.LastCommand = 1 Then
                    seq = Chr(35)
                    SY50.LastCommand = 2
                Else
                    seq = Chr(36)
                    SY50.LastCommand = 1
                End If
                sw.Write(seq & String.Format("F1	{0}	", FormatNumber(Amount, 2)))
                'sw.Write(seq & Chr(70) & Amount & Chr(13) & Chr(10))
                sw.Close()
            End Using
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub SluzbenVlez(ByVal Amount As Decimal)
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic) 'File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If SY50.LastCommand = 1 Then
                    seq = Chr(35)
                    SY50.LastCommand = 2
                Else
                    seq = Chr(36)
                    SY50.LastCommand = 1
                End If
                sw.Write(seq & String.Format("F0	{0}	", FormatNumber(Amount, 2)))
                'sw.Write(seq & Chr(70) & Amount & Chr(13) & Chr(10))
                sw.Close()
            End Using
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub DetalenPeriodicenIzvestaj(ByVal OdDatum As DateTime, ByVal DoDatum As DateTime)
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic)
                Dim seq As Char = Chr(35)
                If SY50.LastCommand = 1 Then
                    seq = Chr(35)
                    SY50.LastCommand = 2
                Else
                    seq = Chr(36)
                    SY50.LastCommand = 1
                End If
                sw.Write(seq & String.Format("^1	{0}	{1}	", OdDatum.ToString("dd-MM-yy"), DoDatum.ToString("dd-MM-yy")))
                sw.Close()
            End Using
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub PeriodicenIzvestaj(ByVal OdDatum As DateTime, ByVal DoDatum As DateTime)
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic) 'File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If SY50.LastCommand = 1 Then
                    seq = Chr(35)
                    SY50.LastCommand = 2
                Else
                    seq = Chr(36)
                    SY50.LastCommand = 1
                End If
                sw.Write(seq & String.Format("^0	{0}	{1}	", OdDatum.ToString("dd-MM-yy"), DoDatum.ToString("dd-MM-yy")))
                'sw.Write(seq + Chr(95) + OdDatum.ToString("ddMMyy") & "," & DoDatum.ToString("ddMMyy") + Chr(13) + Chr(10))
                sw.Close()
            End Using
            Me.Run()
        End Sub
        <ComVisible(True)> _
        Public Sub Diagnostika()
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic)
                Dim seq As Char = Chr(35)
                If SY50.LastCommand = 1 Then
                    seq = Chr(35)
                    SY50.LastCommand = 2
                Else
                    seq = Chr(36)
                    SY50.LastCommand = 1
                End If
                sw.Write(seq & Chr(71) & Chr(13) & Chr(10))
                sw.Close()
            End Using
            Me.Run()
        End Sub
#End Region
#Region "Private Methods"
        Private Sub CreateStornaSY50(Optional ByVal PaidMode As PaidMode = PaidMode.VoGotovo)
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Dim sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic) 'File.CreateText(TextFile) 
            sw.Write(" 01	1		1	" & Chr(13) & Chr(10))
            Dim m As Integer = 0
            For Each tmp As Article In Me.Stavki
                Dim d As Char = Chr(tmp.VAT)
                Dim VAT As Int16 = 1
                Select Case tmp.VAT
                    Case VATgroup.А
                        VAT = 1
                    Case VATgroup.Б
                        VAT = 2
                    Case VATgroup.В
                        VAT = 3
                    Case VATgroup.Г
                        VAT = 4
                End Select
                If m Mod 2 = 0 Then
                    sw.Write(String.Format("#1{0}	{3}	{1}	{2}	0			", tmp.Name, FormatNumber(tmp.Price, 2), FormatNumber(tmp.Amount, 3), VAT) & Chr(13) & Chr(10))

                Else
                    sw.Write(String.Format(" 1{0}	{3}	{1}	{2}	0			", tmp.Name, FormatNumber(tmp.Price, 2), FormatNumber(tmp.Amount, 3), VAT) & Chr(13) & Chr(10))
                End If
                m += 1
            Next
            sw.Write(String.Format("&5{0}		", CInt(PaidMode)) & Chr(13) & Chr(10))
            sw.Write("%8")
            'sw.Write(Chr(37) & Chr(56))
            sw.Flush()
            sw.Close()
        End Sub
        Private Sub CreateFiskalnaSY50(Optional ByVal PaidMode As PaidMode = PaidMode.VoGotovo)
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Dim sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic) 'File.CreateText(TextFile) 
            sw.Write(" 01	1		0	" & Chr(13) & Chr(10))
            Dim m As Integer = 0
            For Each tmp As Article In Me.Stavki
                Dim d As Char = Chr(tmp.VAT)
                Dim VAT As Int16 = 1
                Select Case tmp.VAT
                    Case VATgroup.А
                        VAT = 1
                    Case VATgroup.Б
                        VAT = 2
                    Case VATgroup.В
                        VAT = 3
                    Case VATgroup.Г
                        VAT = 4
                End Select
                If m Mod 2 = 0 Then
                    sw.Write(String.Format("#1{0}	{3}	{1}	{2}	0			", tmp.Name, FormatNumber(tmp.Price, 2), FormatNumber(tmp.Amount, 3), VAT) & Chr(13) & Chr(10))

                Else
                    sw.Write(String.Format(" 1{0}	{3}	{1}	{2}	0			", tmp.Name, FormatNumber(tmp.Price, 2), FormatNumber(tmp.Amount, 3), VAT) & Chr(13) & Chr(10))
                End If
                m += 1
            Next
            sw.Write(String.Format("&5{0}		", CInt(PaidMode)) & Chr(13) & Chr(10))
            sw.Write("%8")
            'sw.Write(Chr(37) & Chr(56))
            sw.Flush()
            sw.Close()
        End Sub
        Private Sub CreateExecutable()
            Using createFile As New FileStream(Me.AppPath, FileMode.CreateNew, FileAccess.Write)
                createFile.Write(My.Resources.sy, 0, My.Resources.sy.Length)
            End Using
        End Sub
        Private Sub CreateIniFile()
            If IO.File.Exists(INIPath) = True Then
                IO.File.Delete(INIPath)
            End If
            Using sw As StreamWriter = File.CreateText(INIPath)
                sw.WriteLine(";Ini fajlot treba da bide vo ist dir. so fiscal32.exe")
                sw.WriteLine("")
                sw.WriteLine("[Setup]")
                sw.WriteLine(String.Format("Port=COM{0}", Me.ComPortNumber))
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
            Return (Strings.Left(Conversions.ToString(num), (Strings.Len(Conversions.ToString(num)) - DecimalDigits)) & "." & Strings.Right(Conversions.ToString(num), DecimalDigits))
        End Function
#End Region
    End Class
End Namespace