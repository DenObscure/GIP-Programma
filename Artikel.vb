Public Class Artikel

    Private ArtikelIDValue As UInteger
    Public Property ArtikelID() As UInteger
        Get
            Return ArtikelIDValue
        End Get
        Set(ByVal value As UInteger)
            ArtikelIDValue = value
        End Set
    End Property

    Private ArtikelOmschrijvingValue As String
    Public Property ArtikelOmschrijving() As String
        Get
            Return ArtikelOmschrijvingValue
        End Get
        Set(ByVal value As String)
            ArtikelOmschrijvingValue = value
        End Set
    End Property

    Private EenheidsprijsValue As Double
    Public Property Eenheidsprijs() As Double
        Get
            Return EenheidsprijsValue
        End Get
        Set(ByVal value As Double)
            EenheidsprijsValue = value
        End Set
    End Property

    Private BtwValue As Integer
    Public Property btw() As Integer
        Get
            Return BtwValue
        End Get
        Set(ByVal value As Integer)
            BtwValue = value
        End Set
    End Property

    Public Sub New()
        ArtikelID = 0
        ArtikelOmschrijving = String.Empty
        Eenheidsprijs = 0
        btw = 21
    End Sub

    Public Sub New(ByVal ArtikelID As UInteger, ByVal ArtikelOmschrijving As String, ByVal Eenheidsprijs As Double, ByVal btw As Integer)
        Me.ArtikelID = ArtikelID
        Me.ArtikelOmschrijving = ArtikelOmschrijving
        Me.Eenheidsprijs = Eenheidsprijs
        Me.btw = btw
    End Sub

End Class
