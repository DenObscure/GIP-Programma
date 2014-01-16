Public Class FactuurArtikels

    Private FactuurIDValue As Integer
    Public Property FactuurID() As Integer
        Get
            Return FactuurIDValue
        End Get
        Set(ByVal value As Integer)
            FactuurIDValue = value
        End Set
    End Property

    Private ArtikelIDValue As Integer
    Public Property ArtikelID() As Integer
        Get
            Return ArtikelIDValue
        End Get
        Set(ByVal value As Integer)
            ArtikelIDValue = value
        End Set
    End Property

    Private AantalValue As Integer
    Public Property Aantal() As Integer
        Get
            Return aantalValue
        End Get
        Set(ByVal value As Integer)
            aantalValue = value
        End Set
    End Property

    Public Sub New()
        FactuurID = 0
        ArtikelID = 0
        Aantal = 0
    End Sub

    Public Sub New(ByVal FactuurID As Integer, ByVal ArtikelID As Integer, ByVal Aantal As Integer)
        Me.FactuurID = FactuurID
        Me.ArtikelID = ArtikelID
        Me.Aantal = Aantal
    End Sub

End Class
