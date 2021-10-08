Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections.Generic
Imports System.IO.Ports
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.IO

Namespace Duna
    <ComVisible(True)> _
    Public NotInheritable Class Razvigorec

#Region "Delegates"
        Public Delegate Sub PrinterDetachedEventHandler()
        Public Delegate Sub PrinterPaperEndEventHandler()
        Public Delegate Sub OpenReceiptEventHandler()
        Public Delegate Sub PaymentModeInReceiptEventHandler()
        Public Delegate Sub DayIsOpenEventHandler()
        Public Delegate Sub ErrorCOMEventHandler(ByVal ErrorCode As String, ByVal ErrorDescription As String)
        Public Delegate Sub FatalErrorEventHandler()
        Public Delegate Sub FiscalMemoryDetachedEventHandler()
        Public Delegate Sub FiscalMemoryFullEventHandler()
        Public Delegate Sub LastCommandSentEventHandler(ByVal CommandText As String)
#End Region
#Region "Events"
        Public Event DayIsOpen As DayIsOpenEventHandler
        Public Event ErrorCOM As ErrorCOMEventHandler
        Public Event FatalError As FatalErrorEventHandler
        Public Event FiscalMemoryDetached As FiscalMemoryDetachedEventHandler
        Public Event FiscalMemoryFull As FiscalMemoryFullEventHandler
        Public Event LastCommandSent As LastCommandSentEventHandler

        Public Event OpenReceipt As OpenReceiptEventHandler
        Public Event PaymentModeInReceipt As PaymentModeInReceiptEventHandler
        Public Event PrinterDetached As PrinterDetachedEventHandler
        Public Event PrinterPaperEnd As PrinterPaperEndEventHandler
#End Region
#Region "Fields"
        <AccessedThroughProperty("COM")> _
        Private _COM As SerialPort
        Private _DailyTotals As DailyTotals
        Private _DayIsOpen As Boolean
        Private _DeviceDateTime As DeviceDateTime
        Private _FatalError As Boolean
        Private _FiscalMemoryDetached As Boolean
        Private _FiscalMemoryFull As Boolean
        Private _OpenReceipt As Boolean
        Private _PaymentModeInReceipt As Boolean
        Private _PrinterDetached As Boolean
        Private _PrinterPaperEnd As Boolean
        Private _ReceiptTotals As ReceiptTotals
        Private _TimeOut As Boolean
        Private Const ACK As Byte = 6
        Public dicProdazba As Dictionary(Of Prodazba, String)
        Public dicProdazbaTip As Dictionary(Of ProdazbaTip, String)
        Private Const ENQ As Byte = 5
        Private Shared ReadOnly ERR_CODE As String() = New String() {"OK", "Има отворена сметка", "Недостасуваат параметри во командата", "", "", "Нема хартија", "Overflow", "Нема доволно количина/износ", "Нема отворено сметка", "Не е програмиран или погрешен податок", "Погрешна команда", "Нема внесено цена", "Нема внесено количина", "Не е дозволено во моменталниот статус", "Overflow - Z извештај", "Overflow - PLU извештај", "Overflow - Оддели извештај", "Overflow - Група извештај", "Overflow - Оператор извештај", "Друга грешка"}
        Private Const ETX As Byte = 3
        Private LastCommand As Char
        Private Const NAK As Byte = &H15
        Private strError As String
        Private Const STX As Byte = &H10
        Private STX_SEQ As Byte
#End Region
#Region "Property"
        Private Property COM() As SerialPort
            Get
                Return Me._COM
            End Get
            <MethodImpl(MethodImplOptions.Synchronized)> _
            Set(ByVal WithEventsValue As SerialPort)
                Dim handler As SerialErrorReceivedEventHandler = New SerialErrorReceivedEventHandler(AddressOf Me.COM_ErrorReceived)
                If (Not Me._COM Is Nothing) Then
                    RemoveHandler Me._COM.ErrorReceived, handler
                End If
                Me._COM = WithEventsValue
                If (Not Me._COM Is Nothing) Then
                    AddHandler Me._COM.ErrorReceived, handler
                End If
            End Set
        End Property
        <ComVisible(True)> _
        Public ReadOnly Property ECR_DailyTotals() As DailyTotals
            Get
                Return Me._DailyTotals
            End Get
        End Property
        <ComVisible(True)> _
        Public ReadOnly Property ECR_DayIsOpen() As Boolean
            Get
                Return Me._DayIsOpen
            End Get
        End Property
        <ComVisible(True)> _
        Public ReadOnly Property ECR_DeviceDateTime() As DeviceDateTime
            Get
                Return Me._DeviceDateTime
            End Get
        End Property
        <ComVisible(True)> _
        Public ReadOnly Property ECR_FatalError() As Boolean
            Get
                Return Me._FatalError
            End Get
        End Property
        <ComVisible(True)> _
        Public ReadOnly Property ECR_FiscalMemoryDetached() As Boolean
            Get
                Return Me._FiscalMemoryDetached
            End Get
        End Property
        <ComVisible(True)> _
        Public ReadOnly Property ECR_FiscalMemoryFull() As Boolean
            Get
                Return Me._FiscalMemoryFull
            End Get
        End Property
        <ComVisible(True)> _
        Public ReadOnly Property ECR_OpenReceipt() As Boolean
            Get
                Return Me._OpenReceipt
            End Get
        End Property
        <ComVisible(True)> _
        Public ReadOnly Property ECR_PaymentModeInReceipt() As Boolean
            Get
                Return Me._PaymentModeInReceipt
            End Get
        End Property
        <ComVisible(True)> _
        Public ReadOnly Property ECR_PrinterDetached() As Boolean
            Get
                Return Me._PrinterDetached
            End Get
        End Property
        <ComVisible(True)> _
        Public ReadOnly Property ECR_PrinterPaperEnd() As Boolean
            Get
                Return Me._PrinterPaperEnd
            End Get
        End Property
        <ComVisible(True)> _
        Public ReadOnly Property ECR_ReceiptTotals() As ReceiptTotals
            Get
                Return Me._ReceiptTotals
            End Get
        End Property
        <ComVisible(True)> _
        Public ReadOnly Property ErrorDescription() As String
            Get
                Return Me.strError
            End Get
        End Property
#End Region
#Region "Ctor"
        Private Sub New()
            Me.STX_SEQ = 0
            Me._TimeOut = False
            Me.dicProdazba = New Dictionary(Of Prodazba, String)
            Me.dicProdazbaTip = New Dictionary(Of ProdazbaTip, String)
        End Sub
        Public Sub New(ByVal ComPortNumber As Integer)
            Me.STX_SEQ = 0
            Me._TimeOut = False
            Me.dicProdazba = New Dictionary(Of Prodazba, String)
            Me.dicProdazbaTip = New Dictionary(Of ProdazbaTip, String)
            Me.dicProdazba.Add(Prodazba.FiskalnaSmetka, "S")
            Me.dicProdazba.Add(Prodazba.OtkazanaSmetka, "V")
            Me.dicProdazba.Add(Prodazba.StornoSmetka, "R")
            Me.dicProdazba.Add(Prodazba.OtkazanaStornoSmetka, "X")
            Me.dicProdazbaTip.Add(ProdazbaTip.PLU, "P")
            Me.dicProdazbaTip.Add(ProdazbaTip.Oddel, "D")
            Me.COM = New SerialPort(("COM" & ComPortNumber.ToString), &H1C200, Parity.None, 8, StopBits.One)
            Me.COM.ReadTimeout = &H1388
            Me.COM.WriteTimeout = &H1388
        End Sub
#End Region
#Region "Private Methods"
        Private Function CacheInOut(ByVal CacheInOutType As CacheInOutType, ByVal PaymentKind As PaymentType, ByVal Amount As Decimal) As Boolean
            Me.LastCommand = "6"c
            Dim data As String = ((((Me.LastCommand.ToString & "/" & Conversions.ToString(CInt(CacheInOutType))) & "/" & Conversions.ToString(CInt(PaymentKind))) & "/" & Me.FormatNumber(Amount, 2)) & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function CancelReceipt() As Boolean
            Me.LastCommand = "9"c
            Dim data As String = (Me.LastCommand.ToString & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function Checksum(ByVal Data As Byte(), ByVal Size As Integer) As Byte
            Dim num2 As Byte = 0
            Dim num4 As Integer = (Size - 1)
            Dim i As Integer = 0
            Do While (i <= num4)
                num2 = CByte((CByte((num2 + Data(i))) Mod &H100))
                i += 1
            Loop
            Return CByte((num2 Mod 100))
        End Function
        Private Sub COM_ErrorReceived(ByVal sender As Object, ByVal e As SerialErrorReceivedEventArgs)
            Me.strError = e.EventType.ToString
            Dim errorCOMEvent As ErrorCOMEventHandler = Me.ErrorCOMEvent
            If (Not errorCOMEvent Is Nothing) Then
                errorCOMEvent.Invoke(e.EventType.ToString, e.EventType.ToString)
            End If
        End Sub
        Private Function COM_SendData(ByVal Paket As String) As Boolean
            Dim flag2 As Boolean
            Dim errorCOMEvent As ErrorCOMEventHandler
            Dim lastCommandSentEvent As LastCommandSentEventHandler = Me.LastCommandSentEvent
            If (Not lastCommandSentEvent Is Nothing) Then
                lastCommandSentEvent.Invoke(Paket)
            End If
            Dim buffer As Byte() = New Byte((Me.COM.ReadBufferSize + 1) - 1) {}
            Try
                If Not Me.COM.IsOpen Then
                    Me.COM.Open()
                End If
                Dim list As New List(Of Byte)
                If (Me.STX_SEQ < 15) Then
                    Me.STX_SEQ = CByte((Me.STX_SEQ + 1))
                Else
                    Me.STX_SEQ = 0
                End If
                list.Add(CByte((&H10 + Me.STX_SEQ)))
                Dim num As Byte
                For Each num In Encoding.GetEncoding(&H4E3).GetBytes(Paket)
                    list.Add(num)
                Next
                list.Add(3)
                Try
                    Me.COM.Write(list.ToArray, 0, list.Count)
                    Me.COM.Read(buffer, 0, Me.COM.ReadBufferSize)
                Catch exception1 As TimeoutException
                    ProjectData.SetProjectError(exception1)
                    Dim exception As TimeoutException = exception1
                    Me._TimeOut = True
                    errorCOMEvent = Me.ErrorCOMEvent
                    If (Not errorCOMEvent Is Nothing) Then
                        errorCOMEvent.Invoke(exception.Message, exception.Message)
                    End If
                    ProjectData.ClearProjectError()
                End Try
                Dim buff As Byte() = Me.TrimByte(buffer)
                Me.ParseReceivedData(buff)
                flag2 = True
                Me.strError = "OK"
                errorCOMEvent = Me.ErrorCOMEvent
                If (Not errorCOMEvent Is Nothing) Then
                    errorCOMEvent.Invoke("0", "OK")
                End If
            Catch exception3 As Exception
                ProjectData.SetProjectError(exception3)
                Dim exception2 As Exception = exception3
                Me.strError = exception2.Message
                errorCOMEvent = Me.ErrorCOMEvent
                If (Not errorCOMEvent Is Nothing) Then
                    errorCOMEvent.Invoke(exception2.Message, exception2.Message)
                End If
                flag2 = False
                ProjectData.ClearProjectError()
            Finally
                If Me.COM.IsOpen Then
                    Me.COM.Close()
                End If
            End Try
            Return flag2
        End Function
        Private Function DiscountSurcharge(ByVal Operation As Prodazba, ByVal Amount As Decimal, ByVal RabatMarza As RabatMarza, ByVal RabatMarzaScope As RabatMarzaScope, ByVal RabatMarzaTip As RabatMarzaTip) As Boolean
            Me.LastCommand = "4"c
            Dim data As String = ((((((Me.LastCommand.ToString & "/" & Me.dicProdazba.Item(Operation)) & "/" & Me.FormatNumber(Amount, 2)) & "/" & Conversions.ToString(CInt(RabatMarza))) & "/" & Conversions.ToString(CInt(RabatMarzaScope))) & "/" & Conversions.ToString(CInt(RabatMarzaTip))) & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function ExitFiscalPrinterMode() As Boolean
            Me.LastCommand = "n"c
            Dim data As String = (Me.LastCommand.ToString & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function FormatDate(ByVal Datum As DateTime) As String
            Return Strings.Format(Datum, "ddMMyy")
        End Function
        Private Function FreeText(ByVal Tekst As String) As Boolean
            Me.LastCommand = "7"c
            Dim data As String = ((Me.LastCommand.ToString & "/" & Strings.Left(Tekst, &H1F)) & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function IssueReportFromFM(ByVal ReportType As ReportType, ByVal StartingDate As String, ByVal EndingDate As String) As Boolean
            Me.LastCommand = "x"c
            Dim data As String = (((((Me.LastCommand.ToString & "/8") & "/" & Conversions.ToString(CInt(ReportType))) & "/" & StartingDate) & "/" & EndingDate) & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function IssueReportFromRAM(ByVal ReportType As ReportType, Optional ByVal Z As Integer = 0) As Boolean
            Me.LastCommand = "x"c
            Dim data As String = (((Me.LastCommand.ToString & "/" & Conversions.ToString(CInt(ReportType))) & "/" & Z.ToString) & "//" & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function ItemSale(ByVal Operation As Prodazba, ByVal ItemName As String, ByVal ItemExtendedName As String, ByVal Quantity As Decimal, ByVal Price As Decimal, ByVal VATgroup As VATgroup, ByVal MacedoniaProduct As MacedonianProduct, ByVal MineralOil As MineralOil) As Boolean
            Me.LastCommand = "3"c
            Dim data As String = (((((((((Me.LastCommand.ToString & "/" & Me.dicProdazba.Item(Operation)) & "/" & Strings.Left(ItemName, &H16)) & "/" & Strings.Left(ItemExtendedName, &H1F)) & "/" & Conversions.ToString(Quantity)) & "/" & Me.FormatNumber(Price, 2)) & "/" & Conversions.ToString(CInt(VATgroup))) & "/" & Conversions.ToString(CInt(MacedoniaProduct))) & "/" & Conversions.ToString(CInt(MineralOil))) & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function LogInOperator(ByVal OperatorNumber As String) As Boolean
            Me.LastCommand = "1"c
            Dim data As String = ((Me.LastCommand.ToString & "/" & OperatorNumber) & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function NonFiscal(ByVal Linii As List(Of String)) As Boolean
            Me.LastCommand = "8"c
            Dim data As String = Me.LastCommand.ToString
            If (Linii.Count > 40) Then
                Linii.RemoveRange(40, (Linii.Count - 40))
            End If
            Dim str2 As String
            For Each str2 In Linii
                data = (data & "/" & str2)
            Next
            data = (data & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function OpenDrawer() As Boolean
            Me.LastCommand = "q"c
            Dim data As String = (Me.LastCommand.ToString & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function PaperFeed() As Boolean
            Me.LastCommand = "w"c
            Dim data As String = (Me.LastCommand.ToString & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Sub ParseErrorFromECR(ByVal err As String)
            Dim index As Integer = Convert.ToInt32(err, &H10)
            Dim errorCOMEvent As ErrorCOMEventHandler = Me.ErrorCOMEvent
            If (Not errorCOMEvent Is Nothing) Then
                errorCOMEvent.Invoke(index.ToString, Razvigorec.ERR_CODE(index))
            End If
        End Sub
        Private Sub ParseReceivedData(ByVal buff As Byte())
            Try
                Dim destinationArray As Byte() = New Byte(((buff.Length - 3) + 1) - 1) {}
                If ((buff(0) = CByte((&H10 + Me.STX_SEQ))) AndAlso (buff((buff.Length - 1)) = 3)) Then
                    Array.Copy(buff, 1, destinationArray, 0, (buff.Length - 2))
                    Dim strArray As String() = Encoding.ASCII.GetString(destinationArray).Split(New Char() {"/"c})
                    If (Strings.Left(strArray(0), 1) = Conversions.ToString(Me.LastCommand)) Then
                        Me.ParseErrorFromECR(Strings.Right(strArray(0), 2))
                        Me.GetDeviceStatus_FatalError(Conversions.ToInteger(strArray(1)))
                        Me.GetDeviceStatus_PrinterPaperEnd(Conversions.ToInteger(strArray(1)))
                        Me.GetDeviceStatus_PrinterDetached(Conversions.ToInteger(strArray(1)))
                        Me.GetDeviceStatus_FiscalMemoryDetached(Conversions.ToInteger(strArray(1)))
                        Me.GetFiscalStatus_DayIsOpen(Conversions.ToInteger(strArray(2)))
                        Me.GetFiscalStatus_OpenReceipt(Conversions.ToInteger(strArray(2)))
                        Me.GetFiscalStatus_PaymentModeInReceipt(Conversions.ToInteger(strArray(2)))
                        Me.GetFiscalStatus_FiscalMemoryFull(Conversions.ToInteger(strArray(2)))
                        Select Case Me.LastCommand
                            Case "0"c
                                Me._DailyTotals.Turnover_Macedonian_goods_fiscal_receipts_VAT_group_A = Conversions.ToDecimal(strArray(3))
                                Me._DailyTotals.Turnover_Macedonian_goods_fiscal_receipts_VAT_group_B = Conversions.ToDecimal(strArray(4))
                                Me._DailyTotals.Turnover_Macedonian_goods_fiscal_receipts_VAT_group_C = Conversions.ToDecimal(strArray(5))
                                Me._DailyTotals.Turnover_Macedonian_goods_fiscal_receipts_VAT_group_D = Conversions.ToDecimal(strArray(6))
                                Me._DailyTotals.VAT_Macedonian_goods_fiscal_receipts_VAT_group_A = Conversions.ToDecimal(strArray(7))
                                Me._DailyTotals.VAT_Macedonian_goods_fiscal_receipts_VAT_group_B = Conversions.ToDecimal(strArray(8))
                                Me._DailyTotals.VAT_Macedonian_goods_fiscal_receipts_VAT_group_C = Conversions.ToDecimal(strArray(9))
                                Me._DailyTotals.VAT_Macedonian_goods_fiscal_receipts_VAT_group_D = Conversions.ToDecimal(strArray(10))
                                Me._DailyTotals.Turnover_total_fiscal_receipts_VAT_group_A = Conversions.ToDecimal(strArray(11))
                                Me._DailyTotals.Turnover_total_fiscal_receipts_VAT_group_B = Conversions.ToDecimal(strArray(12))
                                Me._DailyTotals.Turnover_total_fiscal_receipts_VAT_group_C = Conversions.ToDecimal(strArray(13))
                                Me._DailyTotals.Turnover_total_fiscal_receipts_VAT_group_D = Conversions.ToDecimal(strArray(14))
                                Me._DailyTotals.VAT_total_fiscal_receipts_VAT_group_A = Conversions.ToDecimal(strArray(15))
                                Me._DailyTotals.VAT_total_fiscal_receipts_VAT_group_B = Conversions.ToDecimal(strArray(&H10))
                                Me._DailyTotals.VAT_total_fiscal_receipts_VAT_group_C = Conversions.ToDecimal(strArray(&H11))
                                Me._DailyTotals.VAT_total_fiscal_receipts_VAT_group_D = Conversions.ToDecimal(strArray(&H12))
                                Me._DailyTotals.Turnover_Macedonian_goods_refund_receipts_VAT_group_A = Conversions.ToDecimal(strArray(&H13))
                                Me._DailyTotals.Turnover_Macedonian_goods_refund_receipts_VAT_group_B = Conversions.ToDecimal(strArray(20))
                                Me._DailyTotals.Turnover_Macedonian_goods_refund_receipts_VAT_group_C = Conversions.ToDecimal(strArray(&H15))
                                Me._DailyTotals.Turnover_Macedonian_goods_refund_receipts_VAT_group_D = Conversions.ToDecimal(strArray(&H16))
                                Me._DailyTotals.VAT_Macedonian_goods_refund_receipts_VAT_group_A = Conversions.ToDecimal(strArray(&H17))
                                Me._DailyTotals.VAT_Macedonian_goods_refund_receipts_VAT_group_B = Conversions.ToDecimal(strArray(&H18))
                                Me._DailyTotals.VAT_Macedonian_goods_refund_receipts_VAT_group_C = Conversions.ToDecimal(strArray(&H19))
                                Me._DailyTotals.VAT_Macedonian_goods_refund_receipts_VAT_group_D = Conversions.ToDecimal(strArray(&H1A))
                                Me._DailyTotals.Turnover_total_refund_receipts_VAT_group_A = Conversions.ToDecimal(strArray(&H1B))
                                Me._DailyTotals.Turnover_total_refund_receipts_VAT_group_B = Conversions.ToDecimal(strArray(&H1C))
                                Me._DailyTotals.Turnover_total_refund_receipts_VAT_group_C = Conversions.ToDecimal(strArray(&H1D))
                                Me._DailyTotals.Turnover_total_refund_receipts_VAT_group_D = Conversions.ToDecimal(strArray(30))
                                Me._DailyTotals.VAT_total_refund_receipts_VAT_group_A = Conversions.ToDecimal(strArray(&H1F))
                                Me._DailyTotals.VAT_total_refund_receipts_VAT_group_B = Conversions.ToDecimal(strArray(&H20))
                                Me._DailyTotals.VAT_total_refund_receipts_VAT_group_C = Conversions.ToDecimal(strArray(&H21))
                                Me._DailyTotals.VAT_total_refund_receipts_VAT_group_D = Conversions.ToDecimal(strArray(&H22))
                                Me._DailyTotals.Turnover_of_mineral_oils_with_marked_elements = Conversions.ToDecimal(strArray(&H23))
                                Me._DailyTotals.Quantity_of_mineral_oils_with_marked_elements = Conversions.ToDecimal(strArray(&H24))
                                Me._DailyTotals.Daily_counter_of_fiscal_receipts = Conversions.ToInteger(strArray(&H25))
                                Me._DailyTotals.Daily_counter_of_refund_receipts = Conversions.ToInteger(strArray(&H26))
                                Return
                            Case "r"c
                                Me._ReceiptTotals.UnpaidAmount = Conversions.ToDecimal(strArray(3))
                                Me._ReceiptTotals.DailyNumber = Conversions.ToInteger(strArray(4))
                                Return
                            Case "t"c
                                Me._DeviceDateTime.Date = strArray(3)
                                Me._DeviceDateTime.Time = strArray(4)
                                Return
                        End Select
                    End If
                End If
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Me.strError = exception.Message
                Dim errorCOMEvent As ErrorCOMEventHandler = Me.ErrorCOMEvent
                If (Not errorCOMEvent Is Nothing) Then
                    errorCOMEvent.Invoke(exception.Message, exception.Message)
                End If
                ProjectData.ClearProjectError()
            End Try
        End Sub
        Private Function Payment(ByVal PaymentKind As PaymentType, ByVal Amount As Decimal) As Boolean
            Me.LastCommand = "5"c
            Dim data As String = (((Me.LastCommand.ToString & "/" & Conversions.ToString(CInt(PaymentKind))) & "/" & Me.FormatNumber(Amount, 2)) & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Sub PecatiSmetka(ByVal Smetka As List(Of Article), ByVal NacinNaPlacanje As PaymentType, ByVal TipSmetka As Prodazba, Optional ByVal DadeniPari As Decimal = 0)
            Dim stavka As Article
            For Each stavka In Smetka
                If Not Me._TimeOut Then
                    Me.ItemSale(TipSmetka, stavka.Name, "", stavka.Amount, stavka.Price, stavka.VAT, 0, 0)
                    Thread.Sleep(100)
                End If
            Next
            If Not Me._TimeOut Then
                If (Decimal.Compare(DadeniPari, Decimal.Zero) <> 0) Then
                    Me.Payment(NacinNaPlacanje, DadeniPari)
                Else
                    Me.Payment(NacinNaPlacanje, Decimal.Zero)
                End If
                Thread.Sleep(500)
            End If
            If Me._TimeOut Then
                Me._TimeOut = False
                Me.CancelReceipt()
            End If
        End Sub
        Private Function PluDepSale(ByVal Operation As Prodazba, ByVal Type As ProdazbaTip, ByVal PLUorDEP As String, ByVal Quantity As Decimal, ByVal Price As Decimal) As Boolean
            Me.LastCommand = "2"c
            Dim data As String = ((((((Me.LastCommand.ToString & "/" & Me.dicProdazba.Item(Operation)) & "/" & Me.dicProdazbaTip.Item(Type)) & "/" & PLUorDEP) & "/" & Me.FormatNumber(Quantity, 3)) & "/" & Me.FormatNumber(Price, 2)) & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function ReadDailyTotals() As Boolean
            Me.LastCommand = "0"c
            Dim data As String = (Me.LastCommand.ToString & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function ReadDeviceDateTime() As Boolean
            Me.LastCommand = "t"c
            Dim data As String = (Me.LastCommand.ToString & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function ReadDeviceStatus() As Boolean
            Me.LastCommand = "?"c
            Dim data As String = (Me.LastCommand.ToString & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function ReadReceiptTotals() As Boolean
            Me.LastCommand = "r"c
            Dim data As String = (Me.LastCommand.ToString & "/")
            Return Me.COM_SendData((data & Me.Checksum(data)))
        End Function
        Private Function TrimByte(ByVal buff As Byte()) As Byte()
            Dim list As New List(Of Byte)
            Dim flag As Boolean = False
            Dim num As Byte
            For Each num In buff
                If (num = CByte((&H10 + Me.STX_SEQ))) Then
                    flag = True
                End If
                If flag Then
                    list.Add(num)
                End If
                If (num = 3) Then
                    Exit For
                End If
            Next
            Return list.ToArray
        End Function
#End Region
#Region "Public Mathods"
        <ComVisible(True)> _
        Public Function CacheIn(ByVal Iznos As Decimal, ByVal PaymentKind As PaymentType) As Boolean
            Return Me.CacheInOut(CacheInOutType.CacheIn, PaymentKind, Iznos)
        End Function
        <ComVisible(True)> _
       Public Function CacheOut(ByVal Iznos As Decimal, ByVal PaymentKind As PaymentType) As Boolean
            Return Me.CacheInOut(CacheInOutType.CacheOut, PaymentKind, Iznos)
        End Function
        Public Function Checksum(ByVal Data As String) As String
            Dim str2 As String = Me.Checksum(Encoding.GetEncoding(&H4E3).GetBytes(Data), Data.Length).ToString
            If (Conversions.ToInteger(str2) < 10) Then
                str2 = ("0" & str2)
            End If
            Return str2
        End Function
        <ComVisible(True)> _
       Public Function ExitFiscalPrinter() As Boolean
            Return Me.ExitFiscalPrinterMode
        End Function
        '<ComVisible(True)> _
        'Public Sub FiskalnaSmetka(ByVal Stavki As String(), ByVal NacinNaPlacanje As PaymentType, Optional ByVal Storna As Boolean = False)
        '    Dim flag As Boolean = False
        '    Dim list As New List(Of Article)
        '    Dim str As String
        '    For Each str In Stavki
        '        Dim strArray As String() = str.Split(New Char() {ChrW(9)})
        '        If ((strArray.Length = 4) Or (strArray.Length = 6)) Then
        '            Dim stavka As Article = New Article
        '            stavka.Name = strArray(0)
        '            Select Case strArray(1).ToUpper
        '                Case "A", "А", "1"
        '                    stavka.VAT = VATgroup.A
        '                    Exit Select
        '                Case "B", "Б", "2"
        '                    stavka.VAT = VATgroup.B
        '                    Exit Select
        '                Case "C", "V", "В", "3"
        '                    stavka.VAT = VATgroup.C
        '                    Exit Select
        '                Case "D", "G", "Г", "4"
        '                    stavka.VAT = VATgroup.D
        '                    Exit Select
        '            End Select
        '            stavka.Price = Conversions.ToDecimal(strArray(2))
        '            stavka.Amount = Conversions.ToDecimal(strArray(3))
        '            If (strArray.Length = 6) Then
        '                Dim str3 As String = strArray(4).ToUpper
        '                'If ((((str3 = "1") OrElse (str3 = "T")) OrElse ((str3 = "YES") OrElse (str3 = "DA"))) OrElse (((str3 = "ДА") OrElse (str3 = "Т")) OrElse (str3 = "-1"))) Then
        '                '    stavka.MakedonskiProizvod = MacedonianProduct.Yes
        '                'Else
        '                '    stavka.MakedonskiProizvod = MacedonianProduct.No
        '                'End If
        '                Dim str4 As String = strArray(5).ToUpper
        '                'If ((((str4 = "1") OrElse (str4 = "T")) OrElse ((str4 = "YES") OrElse (str4 = "DA"))) OrElse (((str4 = "ДА") OrElse (str4 = "Т")) OrElse (str4 = "-1"))) Then
        '                '    stavka.MineralnoMaslo = MineralOil.Yes
        '                'Else
        '                '    stavka.MineralnoMaslo = MineralOil.No
        '                'End If
        '            Else
        '                'stavka.MakedonskiProizvod = MacedonianProduct.No
        '                'stavka.MineralnoMaslo = MineralOil.No
        '            End If
        '            list.Add(stavka)
        '            Continue For
        '        End If
        '        flag = True
        '        Exit For
        '    Next
        '    If Not flag Then
        '        If Storna Then
        '            Me.PecatiSmetka(list.ToArray, NacinNaPlacanje, Prodazba.StornoSmetka, Decimal.Zero)
        '        Else
        '            Me.PecatiSmetka(list.ToArray, NacinNaPlacanje, Prodazba.FiskalnaSmetka, Decimal.Zero)
        '        End If
        '    End If
        'End Sub
        <ComVisible(True)> _
       Public Sub FiskalnaSmetka(ByVal Stavki As List(Of Article), ByVal NacinNaPlacanje As PaymentType, Optional ByVal Storna As Boolean = False, Optional ByVal DadeniPari As Decimal = 0)
            If Storna Then
                Me.PecatiSmetka(Stavki, NacinNaPlacanje, Prodazba.StornoSmetka, DadeniPari)
            Else
                Me.PecatiSmetka(Stavki, NacinNaPlacanje, Prodazba.FiskalnaSmetka, DadeniPari)
            End If
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
        <ComVisible(True)> _
        Public Function GetDailyTotals() As DailyTotals
            Dim totals2 As DailyTotals
            If Me.ReadDailyTotals Then
                Return Me._DailyTotals
            End If
            Return totals2
        End Function
        <ComVisible(True)> _
        Public Function GetDeviceDateTime() As DeviceDateTime
            Dim time2 As DeviceDateTime
            If Me.ReadDeviceDateTime Then
                Return Me._DeviceDateTime
            End If
            Return time2
        End Function
        Public Sub GetDeviceStatus_FatalError(ByVal number As Integer)
            Me._FatalError = ((number And 2) <> 0)
            If Me._FatalError Then
                Dim fatalErrorEvent As FatalErrorEventHandler = Me.FatalErrorEvent
                If (Not fatalErrorEvent Is Nothing) Then
                    fatalErrorEvent.Invoke()
                End If
            End If
        End Sub
        Public Sub GetDeviceStatus_FiscalMemoryDetached(ByVal number As Integer)
            Me._FiscalMemoryDetached = ((number And &H40) <> 0)
            If Me._FiscalMemoryDetached Then
                Dim fiscalMemoryDetachedEvent As FiscalMemoryDetachedEventHandler = Me.FiscalMemoryDetachedEvent
                If (Not fiscalMemoryDetachedEvent Is Nothing) Then
                    fiscalMemoryDetachedEvent.Invoke()
                End If
            End If
        End Sub
        Public Sub GetDeviceStatus_PrinterDetached(ByVal number As Integer)
            Me._PrinterDetached = ((number And &H10) <> 0)
            If Me._PrinterDetached Then
                Dim printerDetachedEvent As PrinterDetachedEventHandler = Me.PrinterDetachedEvent
                If (Not printerDetachedEvent Is Nothing) Then
                    printerDetachedEvent.Invoke()
                End If
            End If
        End Sub
        Public Sub GetDeviceStatus_PrinterPaperEnd(ByVal number As Integer)
            Me._PrinterPaperEnd = ((number And 4) <> 0)
            If Me._PrinterPaperEnd Then
                Dim printerPaperEndEvent As PrinterPaperEndEventHandler = Me.PrinterPaperEndEvent
                If (Not printerPaperEndEvent Is Nothing) Then
                    printerPaperEndEvent.Invoke()
                End If
            End If
        End Sub
        Public Sub GetFiscalStatus_DayIsOpen(ByVal number As Integer)
            Me._DayIsOpen = ((number And 2) <> 0)
            If Me._DayIsOpen Then
                Dim dayIsOpenEvent As DayIsOpenEventHandler = Me.DayIsOpenEvent
                If (Not dayIsOpenEvent Is Nothing) Then
                    dayIsOpenEvent.Invoke()
                End If
            End If
        End Sub
        Public Sub GetFiscalStatus_FiscalMemoryFull(ByVal number As Integer)
            Me._FiscalMemoryFull = ((number And &H80) <> 0)
            If Me._FiscalMemoryFull Then
                Dim fiscalMemoryFullEvent As FiscalMemoryFullEventHandler = Me.FiscalMemoryFullEvent
                If (Not fiscalMemoryFullEvent Is Nothing) Then
                    fiscalMemoryFullEvent.Invoke()
                End If
            End If
        End Sub
        Public Sub GetFiscalStatus_OpenReceipt(ByVal number As Integer)
            Me._OpenReceipt = ((number And 4) <> 0)
            If Me._OpenReceipt Then
                Dim openReceiptEvent As OpenReceiptEventHandler = Me.OpenReceiptEvent
                If (Not openReceiptEvent Is Nothing) Then
                    openReceiptEvent.Invoke()
                End If
            End If
        End Sub
        Public Sub GetFiscalStatus_PaymentModeInReceipt(ByVal number As Integer)
            Me._PaymentModeInReceipt = ((number And &H10) <> 0)
            If Me._PaymentModeInReceipt Then
                Dim paymentModeInReceiptEvent As PaymentModeInReceiptEventHandler = Me.PaymentModeInReceiptEvent
                If (Not paymentModeInReceiptEvent Is Nothing) Then
                    paymentModeInReceiptEvent.Invoke()
                End If
            End If
        End Sub
        <ComVisible(True)> _
        Public Function GetReceiptTotals() As ReceiptTotals
            Dim totals2 As ReceiptTotals
            If Me.ReadReceiptTotals Then
                Return Me._ReceiptTotals
            End If
            Return totals2
        End Function
        <ComVisible(True)> _
       Public Function Izvestaj_Grupi() As Boolean
            Return Me.IssueReportFromRAM(ReportType.Group, 0)
        End Function
        <ComVisible(True)> _
        Public Function Izvestaj_Oddel() As Boolean
            Return Me.IssueReportFromRAM(ReportType.DEP, 0)
        End Function
        <ComVisible(True)> _
        Public Function Izvestaj_Operator() As Boolean
            Return Me.IssueReportFromRAM(ReportType.Operator, 0)
        End Function
        <ComVisible(True)> _
        Public Function Izvestaj_Periodicen(ByVal OdDatum As DateTime, ByVal DoDatum As DateTime) As Boolean
            Return Me.IssueReportFromFM(ReportType.Z, Me.FormatDate(OdDatum), Me.FormatDate(DoDatum))
        End Function
        <ComVisible(True)> _
        Public Function Izvestaj_Periodicen(ByVal OdZreport As Integer, ByVal DoZreport As Integer) As Boolean
            Return Me.IssueReportFromFM(ReportType.Z, OdZreport.ToString, DoZreport.ToString)
        End Function
        <ComVisible(True)> _
        Public Function Izvestaj_PeriodicenDetalen(ByVal OdDatum As DateTime, ByVal DoDatum As DateTime) As Boolean
            Return Me.IssueReportFromFM(ReportType.Z, Me.FormatDate(OdDatum), Me.FormatDate(DoDatum))
        End Function
        <ComVisible(True)> _
        Public Function Izvestaj_PeriodicenDetalen(ByVal OdZreport As Integer, ByVal DoZreport As Integer) As Boolean
            Return Me.IssueReportFromFM(ReportType.Z, OdZreport.ToString, DoZreport.ToString)
        End Function
        <ComVisible(True)> _
        Public Function Izvestaj_PLU() As Boolean
            Return Me.IssueReportFromRAM(ReportType.PLU, 0)
        End Function
        <ComVisible(True)> _
        Public Function Izvestaj_PromenaNaDDV(ByVal Z As Integer) As Boolean
            Return Me.IssueReportFromRAM(ReportType.Journal, Z)
        End Function
        <ComVisible(True)> _
        Public Function Izvestaj_X(<DateTimeConstant(0)> Optional ByVal OdDatum As DateTime = #12:00:00 AM#, <DateTimeConstant(0)> Optional ByVal DoDatum As DateTime = #12:00:00 AM#) As Boolean
            If (DateTime.Compare(OdDatum, DateTime.MinValue) = 0) Then
                Return Me.IssueReportFromRAM(ReportType.X, 0)
            End If
            Return Me.IssueReportFromFM(ReportType.X, Me.FormatDate(OdDatum), Me.FormatDate(DoDatum))
        End Function
        <ComVisible(True)> _
        Public Function Izvestaj_Z(ByVal OdDatum As DateTime, ByVal DoDatum As DateTime) As Boolean
            Return Me.IssueReportFromFM(ReportType.Z, Me.FormatDate(OdDatum), Me.FormatDate(DoDatum))
        End Function
        <ComVisible(True)> _
        Public Function Izvestaj_Z(ByVal OdBroj As Integer, ByVal DoBroj As Integer) As Boolean
            Return Me.IssueReportFromFM(ReportType.Z, OdBroj.ToString, DoBroj.ToString)
        End Function
        <ComVisible(True)> _
       Public Function OpenDrawerNow() As Boolean
            Return Me.OpenDrawer
        End Function
        <ComVisible(True)> _
       Public Function PaperFeeedNow() As Boolean
            Return Me.PaperFeed
        End Function
        <ComVisible(True)> _
       Public Sub PrintFreeText(ByVal Linii As String())
            Dim str As String
            For Each str In Linii
                Me.FreeText(Strings.Left(str, &H1F))
            Next
        End Sub
        <ComVisible(True)> _
        Public Sub StornaSmetka(ByVal Stavki As List(Of Article), ByVal NacinNaPlacanje As PaymentType)
            Me.FiskalnaSmetka(Stavki, NacinNaPlacanje, True, Decimal.Zero)
        End Sub
        '<ComVisible(True)> _
        'Public Sub StornaSmetka(ByVal Stavki As String(), ByVal NacinNaPlacanje As PaymentType)
        '    Me.FiskalnaSmetka(Stavki, NacinNaPlacanje, True)
        'End Sub
        <ComVisible(True)> _
      Public Function ZatvoriDen() As Boolean
            Return Me.IssueReportFromRAM(ReportType.Z, 0)
        End Function
#End Region
#Region "Enums"
        <StructLayout(LayoutKind.Sequential)> _
    Public Structure ResponseFromECR
            Public Command As String
            Public ErrorCode As String
            Public DeviceStatus As String
            Public FiscalStatus As String
            Public DataBlok As String
        End Structure
        <StructLayout(LayoutKind.Sequential), ComVisible(True)> _
        Public Structure ReceiptTotals
            Public UnpaidAmount As Decimal
            Public DailyNumber As Integer
        End Structure
        <StructLayout(LayoutKind.Sequential), ComVisible(True)> _
        Public Structure DeviceDateTime
            Public [Date] As String
            Public Time As String
        End Structure
        <StructLayout(LayoutKind.Sequential), ComVisible(True)> _
        Public Structure DailyTotals
            Public Turnover_Macedonian_goods_fiscal_receipts_VAT_group_A As Decimal
            Public Turnover_Macedonian_goods_fiscal_receipts_VAT_group_B As Decimal
            Public Turnover_Macedonian_goods_fiscal_receipts_VAT_group_C As Decimal
            Public Turnover_Macedonian_goods_fiscal_receipts_VAT_group_D As Decimal
            Public VAT_Macedonian_goods_fiscal_receipts_VAT_group_A As Decimal
            Public VAT_Macedonian_goods_fiscal_receipts_VAT_group_B As Decimal
            Public VAT_Macedonian_goods_fiscal_receipts_VAT_group_C As Decimal
            Public VAT_Macedonian_goods_fiscal_receipts_VAT_group_D As Decimal
            Public Turnover_total_fiscal_receipts_VAT_group_A As Decimal
            Public Turnover_total_fiscal_receipts_VAT_group_B As Decimal
            Public Turnover_total_fiscal_receipts_VAT_group_C As Decimal
            Public Turnover_total_fiscal_receipts_VAT_group_D As Decimal
            Public VAT_total_fiscal_receipts_VAT_group_A As Decimal
            Public VAT_total_fiscal_receipts_VAT_group_B As Decimal
            Public VAT_total_fiscal_receipts_VAT_group_C As Decimal
            Public VAT_total_fiscal_receipts_VAT_group_D As Decimal
            Public Turnover_Macedonian_goods_refund_receipts_VAT_group_A As Decimal
            Public Turnover_Macedonian_goods_refund_receipts_VAT_group_B As Decimal
            Public Turnover_Macedonian_goods_refund_receipts_VAT_group_C As Decimal
            Public Turnover_Macedonian_goods_refund_receipts_VAT_group_D As Decimal
            Public VAT_Macedonian_goods_refund_receipts_VAT_group_A As Decimal
            Public VAT_Macedonian_goods_refund_receipts_VAT_group_B As Decimal
            Public VAT_Macedonian_goods_refund_receipts_VAT_group_C As Decimal
            Public VAT_Macedonian_goods_refund_receipts_VAT_group_D As Decimal
            Public Turnover_total_refund_receipts_VAT_group_A As Decimal
            Public Turnover_total_refund_receipts_VAT_group_B As Decimal
            Public Turnover_total_refund_receipts_VAT_group_C As Decimal
            Public Turnover_total_refund_receipts_VAT_group_D As Decimal
            Public VAT_total_refund_receipts_VAT_group_A As Decimal
            Public VAT_total_refund_receipts_VAT_group_B As Decimal
            Public VAT_total_refund_receipts_VAT_group_C As Decimal
            Public VAT_total_refund_receipts_VAT_group_D As Decimal
            Public Turnover_of_mineral_oils_with_marked_elements As Decimal
            Public Quantity_of_mineral_oils_with_marked_elements As Decimal
            Public Daily_counter_of_fiscal_receipts As Integer
            Public Daily_counter_of_refund_receipts As Integer
        End Structure
        Public Enum CacheInOutType
            ' Fields
            CacheIn = 0
            CacheOut = 1
        End Enum
        Public Enum MacedonianProduct
            ' Fields
            No = 0
            Yes = 1
        End Enum
        Public Enum MineralOil
            ' Fields
            No = 0
            Yes = 1
        End Enum
        Public Enum PaymentType
            ' Fields
            SoKarticka = 2
            VoGotovo = 1
        End Enum
        Public Enum Prodazba
            ' Fields
            FiskalnaSmetka = 0
            OtkazanaSmetka = 1
            OtkazanaStornoSmetka = 3
            StornoSmetka = 2
        End Enum
        Public Enum ProdazbaTip
            ' Fields
            Oddel = 1
            PLU = 0
        End Enum
        Public Enum RabatMarza
            ' Fields
            Marza = 1
            Rabat = 0
        End Enum
        Public Enum RabatMarzaScope
            ' Fields
            NaMegusuma = 1
            NaStavka = 0
        End Enum
        Public Enum RabatMarzaTip
            ' Fields
            Iznos = 1
            Procent = 0
        End Enum
        Public Enum ReportType
            ' Fields
            DEP = 4
            FM = 8
            Group = 5
            Journal = 7
            [Operator] = 6
            PLU = 3
            X = 2
            Z = 1
        End Enum
#End Region
    End Class
    Public NotInheritable Class Severec
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
            Me.AppPath = AppConfig.AppLocal & "SEVEREC.EXE"
            Me.INIPath = AppConfig.AppLocal & "FISKAL.INI"
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
        Public Sub ZatvoriDen()
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic) 'File.CreateText(TextFile)
                Dim seq As Char = Chr(35)
                If Severec.LastCommand = 1 Then
                    seq = Chr(35)
                    Severec.LastCommand = 2
                Else
                    seq = Chr(36)
                    Severec.LastCommand = 1
                End If
                sw.Write(" E1" & Chr(13) & Chr(10))
                sw.Write(" ?")
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
                If Severec.LastCommand = 1 Then
                    seq = Chr(35)
                    Severec.LastCommand = 2
                Else
                    seq = Chr(36)
                    Severec.LastCommand = 1
                End If
                sw.Write(" E3" & Chr(13) & Chr(10))
                sw.Write(" ?")
                sw.Close()
            End Using
            Me.Run()
        End Sub

        '<ComVisible(True)> _
        'Public Sub SluzbenIzlez(ByVal Amount As Decimal)
        '    Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
        '    If IO.File.Exists(TextFile) = True Then
        '        IO.File.Delete(TextFile)
        '    End If
        '    Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic) 'File.CreateText(TextFile)
        '        Dim seq As Char = Chr(35)
        '        If Severec.LastCommand = 1 Then
        '            seq = Chr(35)
        '            Severec.LastCommand = 2
        '        Else
        '            seq = Chr(36)
        '            Severec.LastCommand = 1
        '        End If
        '        sw.Write(seq & String.Format("F1	{0}	", FormatNumber(Amount, 2)))
        '        'sw.Write(seq & Chr(70) & Amount & Chr(13) & Chr(10))
        '        sw.Close()
        '    End Using
        '    Me.Run()
        'End Sub
        '<ComVisible(True)> _
        'Public Sub SluzbenVlez(ByVal Amount As Decimal)
        '    Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
        '    If IO.File.Exists(TextFile) = True Then
        '        IO.File.Delete(TextFile)
        '    End If
        '    Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic) 'File.CreateText(TextFile)
        '        Dim seq As Char = Chr(35)
        '        If Severec.LastCommand = 1 Then
        '            seq = Chr(35)
        '            Severec.LastCommand = 2
        '        Else
        '            seq = Chr(36)
        '            Severec.LastCommand = 1
        '        End If
        '        sw.Write(seq & String.Format("F0	{0}	", FormatNumber(Amount, 2)))
        '        'sw.Write(seq & Chr(70) & Amount & Chr(13) & Chr(10))
        '        sw.Close()
        '    End Using
        '    Me.Run()
        'End Sub
        <ComVisible(True)> _
        Public Sub DetalenPeriodicenIzvestaj(ByVal OdDatum As DateTime, ByVal DoDatum As DateTime)
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic)
                Dim seq As Char = Chr(35)
                If Severec.LastCommand = 1 Then
                    seq = Chr(35)
                    Severec.LastCommand = 2
                Else
                    seq = Chr(36)
                    Severec.LastCommand = 1
                End If
                sw.Write(seq & String.Format("O{0},{1}", OdDatum.ToString("ddMMyy"), DoDatum.ToString("ddMMyy")))
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
                If Severec.LastCommand = 1 Then
                    seq = Chr(35)
                    Severec.LastCommand = 2
                Else
                    seq = Chr(36)
                    Severec.LastCommand = 1
                End If
                sw.Write(seq & String.Format("O{0},{1}", OdDatum.ToString("ddMMyy"), DoDatum.ToString("ddMMyy")))
                sw.Close()
            End Using
            Me.Run()
        End Sub
        '<ComVisible(True)> _
        'Public Sub Diagnostika()
        '    Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
        '    If IO.File.Exists(TextFile) = True Then
        '        IO.File.Delete(TextFile)
        '    End If
        '    Using sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic)
        '        Dim seq As Char = Chr(35)
        '        If Severec.LastCommand = 1 Then
        '            seq = Chr(35)
        '            Severec.LastCommand = 2
        '        Else
        '            seq = Chr(36)
        '            Severec.LastCommand = 1
        '        End If
        '        sw.Write(seq & Chr(71) & Chr(13) & Chr(10))
        '        sw.Close()
        '    End Using
        '    Me.Run()
        'End Sub
#End Region
#Region "Private Methods"
        Private Sub CreateStornaSY50(Optional ByVal PaidMode As PaidMode = PaidMode.VoGotovo)
            Dim cyrillic As Encoding = Encoding.GetEncoding("windows-1251")
            If IO.File.Exists(TextFile) = True Then
                IO.File.Delete(TextFile)
            End If
            Dim sw As StreamWriter = New StreamWriter(TextFile, False, cyrillic) 'File.CreateText(TextFile) 
            sw.Write(" 01,0001,1,R" & Chr(13) & Chr(10))
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
                    sw.Write(String.Format("#1{0}	{1}{2}*{3}", tmp.Name, VAT, tmp.Price.ToString("00"), tmp.Amount) & Chr(13) & Chr(10))

                Else
                    sw.Write(String.Format(" 1{0}	{1}{2}*{3}", tmp.Name, VAT, tmp.Price.ToString("00"), tmp.Amount) & Chr(13) & Chr(10))
                End If
                m += 1
            Next
            sw.Write(String.Format("#5	", CInt(PaidMode)) & Chr(13) & Chr(10))
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
            sw.Write(" 01,0001,1" & Chr(13) & Chr(10))
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
                Dim seq As Char
                If Severec.LastCommand = 1 Then
                    seq = Chr(35)
                    Severec.LastCommand = 2
                Else
                    seq = Chr(36)
                    Severec.LastCommand = 1
                End If

                If m Mod 2 = 0 Then
                    sw.Write(String.Format("#1{0}	{1}{2}*{3}", tmp.Name, VAT, tmp.Price.ToString("00"), tmp.Amount) & Chr(13) & Chr(10))

                Else
                    sw.Write(String.Format(" 1{0}	{1}{2}*{3}", tmp.Name, VAT, tmp.Price.ToString("00"), tmp.Amount) & Chr(13) & Chr(10))
                End If
                m += 1
            Next
            sw.Write(String.Format("#5	", CInt(PaidMode)) & Chr(13) & Chr(10))
            sw.Write("%8")
            'sw.Write(Chr(37) & Chr(56))
            sw.Flush()
            sw.Close()
        End Sub
        Private Sub CreateExecutable()
            Using createFile As New FileStream(Me.AppPath, FileMode.CreateNew, FileAccess.Write)
                createFile.Write(My.Resources.SEVEREC, 0, My.Resources.SEVEREC.Length)
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
            Dim startInfo As ProcessStartInfo = New ProcessStartInfo()
            startInfo.CreateNoWindow = True
            startInfo.UseShellExecute = False
            startInfo.FileName = AppPath
            startInfo.WindowStyle = ProcessWindowStyle.Hidden
            startInfo.Arguments = TextFile
            startInfo.WindowStyle = ProcessWindowStyle.Hidden
            Using p As Process = Process.Start(startInfo)
                p.WaitForExit()
            End Using

            'Shell(String.Format("{0} {1}", AppPath, TextFile), AppWinStyle.Hide)
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