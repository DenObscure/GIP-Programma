Public Class Klant

    Private KlantIDValue As UInteger
    Public Property KlantID() As UInteger
        Get
            Return KlantIDValue
        End Get
        Set(ByVal value As UInteger)
            KlantIDValue = value
        End Set
    End Property


    Private BedrijfsnaamValue As String
    Public Property Bedrijfsnaam() As String
        Get
            Return BedrijfsnaamValue
        End Get
        Set(ByVal value As String)
            BedrijfsnaamValue = value
        End Set
    End Property


    Private StraatValue As String
    Public Property Straat() As String
        Get
            Return StraatValue
        End Get
        Set(ByVal value As String)
            StraatValue = value
        End Set
    End Property


    Private NummerValue As String
    Public Property Nummer() As String
        Get
            Return NummerValue
        End Get
        Set(ByVal value As String)
            NummerValue = value
        End Set
    End Property


    Private PostcodeValue As String
    Public Property Postcode() As String
        Get
            Return PostcodeValue
        End Get
        Set(ByVal value As String)
            PostcodeValue = value
        End Set
    End Property


    Private PlaatsValue As String
    Public Property Plaats() As String
        Get
            Return PlaatsValue
        End Get
        Set(ByVal value As String)
            PlaatsValue = value
        End Set
    End Property


    Private LandValue As String
    Public Property Land() As String
        Get
            Return LandValue
        End Get
        Set(ByVal value As String)
            LandValue = value
        End Set
    End Property


    Private BtwnrValue As String
    Public Property Btwnr() As String
        Get
            Return BtwnrValue
        End Get
        Set(ByVal value As String)
            BtwnrValue = value
        End Set
    End Property

    Public Sub New()
        KlantID = 0
        Bedrijfsnaam = String.Empty
        Straat = String.Empty
        Nummer = String.Empty
        Postcode = String.Empty
        Plaats = String.Empty
        Land = String.Empty
        Btwnr = String.Empty
    End Sub

    Public Sub New(ByVal KlantID As UInteger, ByVal Bedrijfsnaam As String, ByVal Straat As String, ByVal Nummer As String, ByVal Postcode As String, ByVal Plaats As String, ByVal Land As String, ByVal Btwnr As String)
        Me.KlantID = KlantID
        Me.Bedrijfsnaam = Bedrijfsnaam
        Me.Straat = Straat
        Me.Nummer = Nummer
        Me.Postcode = Postcode
        Me.Plaats = Plaats
        Me.Land = Land
        Me.Btwnr = Btwnr
    End Sub

End Class
