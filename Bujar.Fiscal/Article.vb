Imports System.Runtime.InteropServices

Public NotInheritable Class Article
#Region "Fields"
    Private _Name As String
    Private _VAT As VATgroup
    Private _Amount As Decimal
    Private _Price As Decimal
#End Region
#Region "Property"
    Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property
    Property VAT() As VATgroup
        Get
            Return _VAT
        End Get
        Set(ByVal value As VATgroup)
            _VAT = value
        End Set
    End Property
    Property Amount() As Decimal
        Get
            Return _Amount
        End Get
        Set(ByVal value As Decimal)
            _Amount = value
        End Set
    End Property
    Property Price() As Decimal
        Get
            Return _Price
        End Get
        Set(ByVal value As Decimal)
            _Price = value
        End Set
    End Property
#End Region
#Region "Ctor"
    Sub New()
        Me.Name = ""
        Me.VAT = VATgroup.А
        Me.Amount = 0
        Me.Price = 0
    End Sub
    Sub New(ByVal name As String, ByVal ddv As VATgroup, ByVal amount As Decimal, ByVal price As Decimal)
        Me.New()
        Me.Name = name
        Me.VAT = ddv
        Me.Amount = amount
        Me.Price = price
    End Sub
#End Region

End Class
Public Enum PaidMode
    SoKredit = 2
    SoKarticka = 1
    VoGotovo = 0
End Enum
Public Enum VATgroup
    ' Fields
    ''' <summary>
    ''' 18%
    ''' </summary>
    ''' <remarks></remarks>
    А = 192
    ''' <summary>
    ''' 5%
    ''' </summary>
    ''' <remarks></remarks>
    Б = 193
    ''' <summary>
    ''' 0%
    ''' </summary>
    ''' <remarks></remarks>
    В = 194
    ''' <summary>
    ''' 0% само кога не е ДДВ обврзник
    ''' </summary>
    ''' <remarks></remarks>
    Г = 195
End Enum
