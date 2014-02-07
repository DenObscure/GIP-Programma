Imports System.Data.Common
Imports System.Data.OleDb

Public Class DBmanager
    Shared strTbl As String
    Private Shared conDBString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + My.Application.Info.DirectoryPath + "\Resources\database\Q8.accdb;Persist Security Info=False;"

    Public Shared Function Getconnection() As OleDbConnection
        Dim conDB = New OleDbConnection
        conDB.ConnectionString = conDBString
        Return conDB
    End Function

    Public Shared Function getKlanten() As Klant()
        Dim aantal = 0
        Dim st() As Klant = Nothing

        Using conDB = Getconnection()
            Using comKlanten = conDB.CreateCommand
                comKlanten.CommandType = CommandType.Text
                comKlanten.CommandText = "SELECT KlantID, Bedrijfsnaam, Straat, Nummer, Postcode, Plaats, Land, Btwnr FROM tblKlant;"
                conDB.Open()
                Using rdrKlanten = comKlanten.ExecuteReader()
                    Dim KlantIDPos = rdrKlanten.GetOrdinal("KlantID")
                    Dim BedrijfsnaamPos = rdrKlanten.GetOrdinal("Bedrijfsnaam")
                    Dim StraatPos = rdrKlanten.GetOrdinal("Straat")
                    Dim NummerPos = rdrKlanten.GetOrdinal("Nummer")
                    Dim PostcodePos = rdrKlanten.GetOrdinal("Postcode")
                    Dim PlaatsPos = rdrKlanten.GetOrdinal("Plaats")
                    Dim LandPos = rdrKlanten.GetOrdinal("Land")
                    Dim BtwnrPos = rdrKlanten.GetOrdinal("Btwnr")
                    Do While rdrKlanten.Read()
                        ReDim Preserve st(aantal)
                        st(aantal) = New Klant(rdrKlanten.GetInt32(KlantIDPos), rdrKlanten.GetString(BedrijfsnaamPos), rdrKlanten.GetString(StraatPos), rdrKlanten.GetString(NummerPos), rdrKlanten.GetString(PostcodePos), rdrKlanten.GetString(PlaatsPos), rdrKlanten.GetString(LandPos), rdrKlanten.GetString(BtwnrPos))
                        aantal += 1
                    Loop
                End Using
            End Using
        End Using
        Return st
    End Function

    Public Shared Function getArtikels() As Artikel()
        Dim aantal = 0
        Dim ar() As Artikel = Nothing

        Using conDB = Getconnection()
            Using comArtikels = conDB.CreateCommand
                comArtikels.CommandType = CommandType.Text
                comArtikels.CommandText = "SELECT ArtikelID, ArtikelOmschrijving, Eenheidsprijs, BTW FROM tblArtikel;"
                conDB.Open()
                Using rdrArtikels = comArtikels.ExecuteReader()
                    Dim ArtikelIDPos = rdrArtikels.GetOrdinal("ArtikelID")
                    Dim ArtikelOmschrijvingPos = rdrArtikels.GetOrdinal("ArtikelOmschrijving")
                    Dim EenheidsprijsPos = rdrArtikels.GetOrdinal("Eenheidsprijs")
                    Dim BTWPos = rdrArtikels.GetOrdinal("BTW")
                    Do While rdrArtikels.Read()
                        ReDim Preserve ar(aantal)
                        ar(aantal) = New Artikel(rdrArtikels.GetInt32(ArtikelIDPos), rdrArtikels.GetString(ArtikelOmschrijvingPos), rdrArtikels.GetDouble(EenheidsprijsPos), rdrArtikels.GetInt32(BTWPos))
                        aantal += 1
                    Loop
                End Using
            End Using
        End Using
        Return ar
    End Function

    Public Shared Function ShowArtikel() As System.Data.DataSet
        Dim sql As String = "SELECT * FROM tblArtikel;"

        Dim dataadapter As New OleDbDataAdapter(sql, conDBString)
        Dim ds As New DataSet()
        Getconnection.Open()
        strTbl = "tblArtikel"
        dataadapter.Fill(ds, "Authors_table")
        Getconnection.Close()
        frmMain.dgvPrijzen.DataSource = ds
        frmMain.dgvPrijzen.DataMember = "Authors_table"
        frmMain.dgvShowArtikels.DataSource = ds
        frmMain.dgvShowArtikels.DataMember = "Authors_table"
        Return ds
    End Function

    Public Shared Function SelectedKlant(ByVal SelectedK As String) As Klant
        Using conDB = Getconnection()
            Using comKlanten = conDB.CreateCommand
                comKlanten.CommandType = CommandType.Text
                comKlanten.CommandText = "SELECT KlantID, Bedrijfsnaam, Straat, Nummer, Postcode, Plaats, Land, Btwnr FROM tblKlant WHERE Bedrijfsnaam = '" & SelectedK & "'"
                conDB.Open()
                Using rdrKlanten = comKlanten.ExecuteReader()
                    Dim KlantIDPos = rdrKlanten.GetOrdinal("KlantID")
                    Dim BedrijfsnaamPos = rdrKlanten.GetOrdinal("Bedrijfsnaam")
                    Dim StraatPos = rdrKlanten.GetOrdinal("Straat")
                    Dim NummerPos = rdrKlanten.GetOrdinal("Nummer")
                    Dim PostcodePos = rdrKlanten.GetOrdinal("Postcode")
                    Dim PlaatsPos = rdrKlanten.GetOrdinal("Plaats")
                    Dim LandPos = rdrKlanten.GetOrdinal("Land")
                    Dim BtwnrPos = rdrKlanten.GetOrdinal("Btwnr")
                    rdrKlanten.Read()
                    Dim k As New Klant(rdrKlanten.GetInt32(KlantIDPos), rdrKlanten.GetString(BedrijfsnaamPos), rdrKlanten.GetString(StraatPos), rdrKlanten.GetString(NummerPos), rdrKlanten.GetString(PostcodePos), rdrKlanten.GetString(PlaatsPos), rdrKlanten.GetString(LandPos), rdrKlanten.GetString(BtwnrPos))
                    Return k
                End Using
            End Using
        End Using
    End Function

    Public Shared Function ShowFacturen() As System.Data.DataSet
        Dim sql As String = "SELECT * FROM tblFactuur"
        '"SELECT Id, datum, vervaldatum, KlantID, TeBetalen, tblGegevens.ArtikelID, Aantal, ArtikelOmschrijving FROM tblFactuur, tblGegevens, tblArtikel where tblFactuur.id=tblGegevens.FactuurID and tblGegevens.artikelID=tblArtikel.artikelid;"
        Dim dataadapter As New OleDbDataAdapter(sql, conDBString)
        Dim ds As New DataSet()
        Getconnection.Open()
        strTbl = "tblFactuur"
        dataadapter.Fill(ds, "Authors_table")
        Getconnection.Close()
        frmMain.dgvShowFacturen.DataSource = ds
        frmMain.dgvShowFacturen.DataMember = "Authors_table"
        Return ds
    End Function

    Public Shared Function SelectedItem(ByVal SelectedA As String) As Artikel
        Using conDB = Getconnection()
            Using comArtikel = conDB.CreateCommand
                comArtikel.CommandType = CommandType.Text
                comArtikel.CommandText = "SELECT ArtikelID, ArtikelOmschrijving, Eenheidsprijs, BTW FROM tblArtikel WHERE Artikelomschrijving = '" & SelectedA & "'"
                conDB.Open()
                Using rdrArtikel = comArtikel.ExecuteReader()
                    Dim ArtikelIDPos = rdrArtikel.GetOrdinal("ArtikelID")
                    Dim ArtikelOmschrijvingPos = rdrArtikel.GetOrdinal("ArtikelOmschrijving")
                    Dim EenheidsprijsPos = rdrArtikel.GetOrdinal("Eenheidsprijs")
                    Dim BTWPos = rdrArtikel.GetOrdinal("BTW")
                    rdrArtikel.Read()
                    Dim a As New Artikel(rdrArtikel.GetInt32(ArtikelIDPos), rdrArtikel.GetString(ArtikelOmschrijvingPos), rdrArtikel.GetDouble(EenheidsprijsPos), rdrArtikel.GetInt32(BTWPos))
                    Return a
                End Using
            End Using
        End Using
    End Function

    Public Shared Function FactuurToevoegen(ByVal fac As Factuur) As Integer
        Using conDB = DBmanager.Getconnection()
            Using comAddFact = conDB.CreateCommand()
                comAddFact.CommandType = CommandType.Text
                comAddFact.CommandText = "INSERT INTO tblFactuur VALUES ('" & fac.Id & "','" & fac.Datum & "','" & fac.Vervaldatum & "','" & fac.KlantId & "','" & fac.TeBetalen & "')"
                conDB.Open()
                Return comAddFact.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Shared Function FactuurGegevensToevoegen(ByVal f As FactuurArtikels) As Integer
        Using conDB = DBmanager.Getconnection()
            Using comAddFact = conDB.CreateCommand()
                comAddFact.CommandType = CommandType.Text
                If IsNothing(f) Then
                    Return 0
                Else
                    comAddFact.CommandText = "insert into tblGegevens(FactuurID, ArtikelID, Aantal) values ('" & f.FactuurID & "','" & f.ArtikelID & "','" & f.Aantal & "')"
                End If
                conDB.Open()
                Return comAddFact.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Shared Function KlantGegevensToevoegen(ByVal k As Klant) As Integer
        Using conDB = DBmanager.Getconnection()
            Using comAddFact = conDB.CreateCommand()
                comAddFact.CommandType = CommandType.Text
                If IsNothing(k) Then
                    Return 0
                Else
                    comAddFact.CommandText = "INSERT INTO tblKlant(KlantID, Bedrijfsnaam, Straat, Nummer, Postcode, Plaats, Land, BTWnr) values ('" & k.KlantID & "','" & k.Bedrijfsnaam & "','" & k.Straat & "','" & k.Nummer & "','" & k.Postcode & "','" & k.Plaats & "','" & k.Land & "','" & k.Btwnr & "')"
                End If
                conDB.Open()
                Return comAddFact.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Shared Function ArtikelToevoegen(ByVal a As Artikel) As Integer
        Using conDB = DBmanager.Getconnection()
            Using comAddFact = conDB.CreateCommand()
                comAddFact.CommandType = CommandType.Text
                If IsNothing(a) Then
                    Return 0
                Else
                    comAddFact.CommandText = "insert into tblArtikel(ArtikelID, ArtikelOmschrijving, Eenheidsprijs, BTW) values ('" & a.ArtikelID & "','" & a.ArtikelOmschrijving & "','" & a.Eenheidsprijs & "','" & a.btw & "')"
                End If
                conDB.Open()
                Return comAddFact.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Shared Function SelectedFactuur(ByVal SelectedF As Integer)
        Dim sql As String = "SELECT  Artikelomschrijving, aantal FROM tblGegevens, tblArtikel WHERE tblArtikel.ArtikelID=tblGegevens.artikelID AND tblGegevens.FactuurID=" & SelectedF
        
        Dim dataadapter As New OleDbDataAdapter(sql, conDBString)
        Dim ds As New DataSet()
        Getconnection.Open()
        strTbl = "tblFactuur"
        dataadapter.Fill(ds, "Authors_table")
        Getconnection.Close()
        frmMain.dgvDetails.DataSource = ds
        frmMain.dgvDetails.DataMember = "Authors_table"
        Return ds
    End Function

    Public Shared Function ShowKlanten() As System.Data.DataSet
        Dim sql As String = "SELECT * FROM tblKlant"
        Dim dataadapter As New OleDbDataAdapter(sql, conDBString)
        Dim ds As New DataSet()
        Getconnection.Open()
        strTbl = "tblKlant"
        dataadapter.Fill(ds, "Authors_table")
        Getconnection.Close()
        frmMain.dgvShowKlanten.DataSource = ds
        frmMain.dgvShowKlanten.DataMember = "Authors_table"
        frmMain.dgvEditKlanten.DataSource = ds
        frmMain.dgvEditKlanten.DataMember = "Authors_table"
        Return ds
    End Function

    Public Shared Function UpdateDB(ByVal klantID As Integer, ByVal Bedrijfsnaam As String, ByVal Straat As String, ByVal Nummer As String, ByVal Postcode As String, ByVal Plaats As String, ByVal Land As String, ByVal BTWnr As String) As Integer
        Using conDB = Getconnection()
            Dim SqlString As String = "UPDATE [tblKlant] SET [Bedrijfsnaam] = ?, [Straat] = ?, [Nummer] = ?, [Postcode] = ?, [Plaats] = ?, [Land] = ?, [BTWnr] = ? WHERE KlantID = ?"
            Using cmd As New OleDbCommand(SqlString, conDB)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.AddWithValue("Bedrijfsnaam", Bedrijfsnaam)
                cmd.Parameters.AddWithValue("Straat", Straat)
                cmd.Parameters.AddWithValue("Nummer", Nummer)
                cmd.Parameters.AddWithValue("Postcode", Postcode)
                cmd.Parameters.AddWithValue("Plaats", Plaats)
                cmd.Parameters.AddWithValue("Land", Land)
                cmd.Parameters.AddWithValue("BTWnr", BTWnr)
                cmd.Parameters.AddWithValue("KlantID", klantID)
                conDB.Open()
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Shared Function UpdateDbArt(a As Artikel) As Integer
        Using conDB = Getconnection()
            Dim SqlString As String = "UPDATE [tblArtikel] SET [ArtikelOmschrijving] = ?, [Eenheidsprijs] = ?, [BTW] = ? WHERE KlantID = ?"
            Using cmd As New OleDbCommand(SqlString, conDB)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.AddWithValue("ArtikelOmschrijving", a.ArtikelOmschrijving)
                cmd.Parameters.AddWithValue("Eenheidsprijs", a.Eenheidsprijs)
                cmd.Parameters.AddWithValue("BTW", a.btw)
                cmd.Parameters.AddWithValue("ArtikelID", a.ArtikelID)
                conDB.Open()
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Shared Function DeleteRowKlant(ByVal SelectedRow As Integer) As Integer
        DeleteFacGegFromKlant(SelectedRow)
        DeleteFacFromKlant(SelectedRow)
        Using conDB = Getconnection()
            Dim SqlString As String = "DELETE FROM tblKlant WHERE KlantID = ?"
            Using cmd As New OleDbCommand(SqlString, conDB)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.AddWithValue("selectedRow", SelectedRow)
                conDB.Open()
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Shared Function DeleteFacGegFromKlant(ByVal SelectedRow As Integer) As Integer
        Using conDB = Getconnection()
            Dim SqlString As String = "delete g.*  from tblgegevens g  inner join tblfactuur f ON f.Id= g.factuurID where f.KlantID = ?;"
            Using cmd As New OleDbCommand(SqlString, conDB)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.AddWithValue("selectedRow", SelectedRow)
                conDB.Open()
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Shared Function DeleteFacFromKlant(ByVal SelectedRow As Integer) As Integer
        Using conDB = Getconnection()
            Dim SqlString As String = "DELETE FROM tblFactuur WHERE KlantID = ?"
            Using cmd As New OleDbCommand(SqlString, conDB)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.AddWithValue("selectedRow", SelectedRow)
                conDB.Open()
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Shared Function DeleteRowFacGeg(ByVal SelectedRow As Integer) As Integer
        Using conDB = Getconnection()
            Dim SqlString As String = "DELETE FROM tblGegevens WHERE FactuurID=?"
            Using cmd As New OleDbCommand(SqlString, conDB)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.AddWithValue("selectedRow", SelectedRow)
                conDB.Open()
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Shared Function DeleteRowFac(ByVal SelectedRow As Integer) As Integer
        Using conDB = Getconnection()
            Dim SqlString As String = "DELETE FROM tblFactuur WHERE Id=?"
            Using cmd As New OleDbCommand(SqlString, conDB)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.AddWithValue("selectedRow", SelectedRow)
                conDB.Open()
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Shared Function DeleteRowArtikel(ByVal SelectedRow As Integer) As Integer
        Using conDB = Getconnection()
            Dim SqlString As String = "DELETE FROM tblArtikel WHERE ArtikelID = ?"
            Using cmd As New OleDbCommand(SqlString, conDB)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.AddWithValue("selectedRow", SelectedRow)
                conDB.Open()
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function

End Class