Public Class Factuur

    Private IdValue As Integer
    Public Property Id() As Integer
        Get
            Return IdValue
        End Get
        Set(ByVal value As Integer)
            IdValue = value
        End Set
    End Property

    Private DatumValue As Date
    Public Property Datum() As Date
        Get
            Return DatumValue
        End Get
        Set(ByVal value As Date)
            DatumValue = value
        End Set
    End Property

    Private VervaldatumValue As Date
    Public Property Vervaldatum() As Date
        Get
            Return VervaldatumValue
        End Get
        Set(ByVal value As Date)
            VervaldatumValue = value
        End Set
    End Property

    Private KlantIDValue As Integer
    Public Property KlantId() As Integer
        Get
            Return KlantidValue
        End Get
        Set(ByVal value As Integer)
            KlantidValue = value
        End Set
    End Property

    Private TeBetalenValue As Double
    Public Property TeBetalen() As Double
        Get
            Return TeBetalenValue
        End Get
        Set(ByVal value As Double)
            TeBetalenValue = value
        End Set
    End Property

    Public Sub New()
        Id = 0
        Datum = Date.Today()
        Vervaldatum = #1/1/2013#
        KlantId = 0
        TeBetalen = 0
    End Sub

    Public Sub New(ByVal Id As Integer, ByVal Datum As Date, ByVal Vervaldatum As Date, ByVal KlantId As Integer, ByVal TeBetalen As Double)
        Me.Id = Id
        Me.Datum = Datum
        Me.Vervaldatum = Vervaldatum
        Me.KlantId = KlantId
        Me.TeBetalen = TeBetalen
    End Sub

End Class
