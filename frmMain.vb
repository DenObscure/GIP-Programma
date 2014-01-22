Imports System.Data.OleDb

Public Class frmMain
    Dim klanten() As Klant = DBmanager.getKlanten()
    Dim artikels() As Artikel = DBmanager.getArtikels()
    Dim FactuurArtikels() As FactuurArtikels
    Dim reset As Boolean = False
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'form setts
        txtAantal1.Enabled = False
        txtAantal2.Enabled = False
        txtAantal3.Enabled = False
        txtAantal4.Enabled = False
        txtAantal5.Enabled = False
        txtAantal6.Enabled = False
        txtAantal7.Enabled = False
        txtAantal8.Enabled = False
        txtAantal9.Enabled = False
        txtAantal10.Enabled = False
        'COUNT & DATE
        DBmanager.ShowFacturen()                        'onload voor facnr count
        lblFacNr.Text = dgvShowFacturen.RowCount()      ' ""
        lblFacDat.Text = Today()
        'LOAD IMAGES
        Dim image1 As Image = Image.FromFile("Resources\img\Invoice3.png")
        pbInvoice.Image = image1

        Dim image2 As Image = Image.FromFile("Resources\img\q82.gif")
        pbQ8.Image = image2
        pbQ82.Image = image2

        Dim image3 As Image = Image.FromFile("Resources\img\invoice.png")
        pbUitkomst.Image = image3

        Dim image4 As Image = Image.FromFile("Resources\img\pijl.jpg")
        pbPijl.Image = image4


        'DB CONNECTION
        Try
            DBmanager.Getconnection.Open()              'Makes connection using DBMANAGER function "Getconnection()"  TO CHANGE DB CHANGE PATH IN DBMANAGER
            lblStatus.ForeColor = Color.Green           'Status If Connection is OK
            lblStatus.Text = "OK!"
            pbStatus.Value = 100
        Catch ex As Exception
            lblStatus.Text = ex.Message
            lblStatus.ForeColor = Color.Red             'Status when error with connection
            lblStatus.Text = "Nope! Error:" & ex.Message
        End Try

        'Fill dbKlant
        For Each k In klanten
            cbKlant.Items.Add(k.Bedrijfsnaam.ToString)
        Next

        'Fill cbomschrijvingen
        For Each a In artikels
            cbOms1.Items.Add(a.ArtikelOmschrijving.ToString)
            cbOms2.Items.Add(a.ArtikelOmschrijving.ToString)
            cbOms3.Items.Add(a.ArtikelOmschrijving.ToString)
            cbOms4.Items.Add(a.ArtikelOmschrijving.ToString)
            cbOms5.Items.Add(a.ArtikelOmschrijving.ToString)
            cbOms6.Items.Add(a.ArtikelOmschrijving.ToString)
            cbOms7.Items.Add(a.ArtikelOmschrijving.ToString)
            cbOms8.Items.Add(a.ArtikelOmschrijving.ToString)
            cbOms9.Items.Add(a.ArtikelOmschrijving.ToString)
            cbOms10.Items.Add(a.ArtikelOmschrijving.ToString)
        Next
    End Sub

    Private Sub NieuweFactuurToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NieuweFactuurToolStripMenuItem.Click
        TabControl1.SelectedTab = TabNewFac
    End Sub

    Private Sub PrijzenAanpassenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrijzenAanpassenToolStripMenuItem.Click
        TabControl1.SelectedTab = TabEditPrijzen
    End Sub

    Private Sub btnNewFact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewFact.Click
        TabControl1.SelectedTab = TabNewFac
        ToolTip1.Show("Kies een klant", cbKlant)
    End Sub

    Private Sub btnPrijzenAanp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrijzenAanp.Click
        DBmanager.ShowArtikel()
        TabControl1.SelectedTab = TabEditPrijzen
    End Sub

    Private Sub btnAddKlant_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddKlant.Click
        DBmanager.ShowKlanten()
        txtNewKlantID.Text = dgvShowKlanten.RowCount() + 1
        TabControl1.SelectedTab = TabNewKlant
    End Sub

    Private Sub btnKlantAanp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKlantAanp.Click
        DBmanager.ShowKlanten()
        txtNewKlantID.Text = dgvShowKlanten.RowCount() + 1
        TabControl1.SelectedTab = TabEditKlanten
    End Sub

    Private Sub frmMain_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        TabControl1.Width = Me.Width
        TabControl1.Height = Me.Height
        gbMenu.Height = StatusStrip1.Location.Y - 22
    End Sub

    Private Sub cbKlant_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbKlant.SelectedIndexChanged
        Dim SelectedKlant As String = cbKlant.SelectedItem
        Dim GegevensKlant As Klant = DBmanager.SelectedKlant(SelectedKlant)
        lblBtwNr.Text = GegevensKlant.Btwnr
        lblGemeente.Text = GegevensKlant.Plaats
        lblLand.Text = GegevensKlant.Land
        lblPostcode.Text = GegevensKlant.Postcode
        lblStraat.Text = GegevensKlant.Straat
        lblStraatNr.Text = GegevensKlant.Nummer
        ToolTip1.Hide(cbKlant)
    End Sub

    Private Sub btnShowFacs_Click(sender As System.Object, e As System.EventArgs) Handles btnShowFacs.Click
        DBmanager.ShowFacturen()
        TabControl1.SelectedTab = TabShowFacs
        lblFacNr.Text = dgvShowFacturen.RowCount()

        Dim selectedF As Integer ' als db leeg is
        If Not dgvShowFacturen.RowCount = 1 Then
            selectedF = dgvShowFacturen.CurrentRow.Cells("id").Value
            DBmanager.SelectedFactuur(selectedF)
        End If
    End Sub

    Private Sub cbOms1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOms1.SelectedIndexChanged
        If reset = False Then
            If Not cbOms1.SelectedIndex = -1 Then
                Dim SelectedItem As String = cbOms1.SelectedItem
                Dim GegevensItem As Artikel = DBmanager.SelectedItem(SelectedItem)
                lblPrijs1.Text = "€ " & GegevensItem.Eenheidsprijs
                lblArtId1.Text = cbOms1.SelectedIndex + 1
                lblBtw1.Text = GegevensItem.btw & " %"
                txtAantal1.Enabled = True
            Else
                txtAantal1.Enabled = False
            End If
            lblTotEx1.Text = Integer.Parse(txtAantal1.Text) * Double.Parse(lblPrijs1.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub cbOms2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOms2.SelectedIndexChanged
        If reset = False Then
            If cbOms2.SelectedItem.ToString = cbOms1.SelectedItem.ToString Then
                MsgBox("Je kan niet 2 dezelfde Items selecteren")
                cbOms2.SelectedIndex = -1
            End If
            If Not cbOms2.SelectedIndex = -1 Then
                Dim SelectedItem As String = cbOms2.SelectedItem
                Dim GegevensItem As Artikel = DBmanager.SelectedItem(SelectedItem)
                lblPrijs2.Text = "€ " & GegevensItem.Eenheidsprijs
                lblArtId2.Text = cbOms2.SelectedIndex + 1
                lblBtw2.Text = GegevensItem.btw & " %"
                txtAantal2.Enabled = True
            Else
                txtAantal2.Enabled = False
            End If
            lblTotEx2.Text = Integer.Parse(txtAantal2.Text) * Double.Parse(lblPrijs2.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub cbOms3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOms3.SelectedIndexChanged
        If reset = False Then
            If cbOms3.SelectedItem.ToString = cbOms2.SelectedItem.ToString Then
                MsgBox("Je kan niet 2 dezelfde Items selecteren")
                cbOms3.SelectedIndex = -1
            End If

            If Not cbOms3.SelectedIndex = -1 Then
                Dim SelectedItem As String = cbOms3.SelectedItem
                Dim GegevensItem As Artikel = DBmanager.SelectedItem(SelectedItem)
                lblPrijs3.Text = "€ " & GegevensItem.Eenheidsprijs
                lblArtId3.Text = cbOms3.SelectedIndex + 1
                lblBtw3.Text = GegevensItem.btw & " %"
                txtAantal3.Enabled = True
            Else
                txtAantal3.Enabled = False
            End If
            lblTotEx3.Text = Integer.Parse(txtAantal3.Text) * Double.Parse(lblPrijs3.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub cbOms4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOms4.SelectedIndexChanged
        If reset = False Then
            If cbOms4.SelectedItem.ToString = cbOms3.SelectedItem.ToString Then
                MsgBox("Je kan niet 2 dezelfde Items selecteren")
                cbOms4.SelectedIndex = -1
            End If

            If Not cbOms4.SelectedIndex = -1 Then
                Dim SelectedItem As String = cbOms4.SelectedItem
                Dim GegevensItem As Artikel = DBmanager.SelectedItem(SelectedItem)
                lblPrijs4.Text = "€ " & GegevensItem.Eenheidsprijs
                lblArtId4.Text = cbOms4.SelectedIndex + 1
                lblBtw4.Text = GegevensItem.btw & " %"
                txtAantal4.Enabled = True
            Else
                txtAantal4.Enabled = False
            End If
            lblTotEx4.Text = Integer.Parse(txtAantal4.Text) * Double.Parse(lblPrijs4.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub cbOms5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOms5.SelectedIndexChanged
        If reset = False Then
            If cbOms5.SelectedItem.ToString = cbOms5.SelectedItem.ToString Then
                MsgBox("Je kan niet 2 dezelfde Items selecteren")
                cbOms5.SelectedIndex = -1
            End If

            If Not cbOms5.SelectedIndex = -1 Then
                Dim SelectedItem As String = cbOms5.SelectedItem
                Dim GegevensItem As Artikel = DBmanager.SelectedItem(SelectedItem)
                lblPrijs5.Text = "€ " & GegevensItem.Eenheidsprijs
                lblArtId5.Text = cbOms5.SelectedIndex + 1
                lblBtw5.Text = GegevensItem.btw & " %"
                txtAantal5.Enabled = True
            Else
                txtAantal5.Enabled = False
            End If
            lblTotEx5.Text = Integer.Parse(txtAantal5.Text) * Double.Parse(lblPrijs5.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub cbOms6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOms6.SelectedIndexChanged
        If reset = False Then
            If cbOms6.SelectedItem.ToString = cbOms5.SelectedItem.ToString Then
                MsgBox("Je kan niet 2 dezelfde Items selecteren")
                cbOms6.SelectedIndex = -1
            End If

            If Not cbOms6.SelectedIndex = -1 Then
                Dim SelectedItem As String = cbOms6.SelectedItem
                Dim GegevensItem As Artikel = DBmanager.SelectedItem(SelectedItem)
                lblPrijs6.Text = "€ " & GegevensItem.Eenheidsprijs
                lblArtId6.Text = cbOms6.SelectedIndex + 1
                lblBtw6.Text = GegevensItem.btw & " %"
                txtAantal6.Enabled = True
            Else
                txtAantal6.Enabled = False
            End If
            lblTotEx6.Text = Integer.Parse(txtAantal6.Text) * Double.Parse(lblPrijs6.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub cbOms7_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOms7.SelectedIndexChanged
        If reset = False Then
            If cbOms7.SelectedItem.ToString = cbOms8.SelectedItem.ToString Then
                MsgBox("Je kan niet 2 dezelfde Items selecteren")
                cbOms7.SelectedIndex = -1
            End If

            If Not cbOms7.SelectedIndex = -1 Then
                Dim SelectedItem As String = cbOms7.SelectedItem
                Dim GegevensItem As Artikel = DBmanager.SelectedItem(SelectedItem)
                lblPrijs7.Text = "€ " & GegevensItem.Eenheidsprijs
                lblArtId7.Text = cbOms7.SelectedIndex + 1
                lblBtw7.Text = GegevensItem.btw & " %"
                txtAantal7.Enabled = True
            Else
                txtAantal7.Enabled = False
            End If
            lblTotEx7.Text = Integer.Parse(txtAantal7.Text) * Double.Parse(lblPrijs7.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub
    Private Sub cbOms8_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOms8.SelectedIndexChanged
        If reset = False Then
            If cbOms8.SelectedItem.ToString = cbOms7.SelectedItem.ToString Then
                MsgBox("Je kan niet 2 dezelfde Items selecteren")
                cbOms8.SelectedIndex = -1
            End If

            If Not cbOms8.SelectedIndex = -1 Then
                Dim SelectedItem As String = cbOms8.SelectedItem
                Dim GegevensItem As Artikel = DBmanager.SelectedItem(SelectedItem)
                lblPrijs8.Text = "€ " & GegevensItem.Eenheidsprijs
                lblArtId8.Text = cbOms8.SelectedIndex + 1
                lblBtw8.Text = GegevensItem.btw & " %"
                txtAantal8.Enabled = True
            Else
                txtAantal8.Enabled = False
            End If
            lblTotEx8.Text = Integer.Parse(txtAantal8.Text) * Double.Parse(lblPrijs8.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub cbOms9_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOms9.SelectedIndexChanged
        If reset = False Then
            If cbOms9.SelectedItem.ToString = cbOms8.SelectedItem.ToString Then
                MsgBox("Je kan niet 2 dezelfde Items selecteren")
                cbOms9.SelectedIndex = -1
            End If

            If Not cbOms9.SelectedIndex = -1 Then
                Dim SelectedItem As String = cbOms9.SelectedItem
                Dim GegevensItem As Artikel = DBmanager.SelectedItem(SelectedItem)
                lblPrijs9.Text = "€ " & GegevensItem.Eenheidsprijs
                lblArtId9.Text = cbOms9.SelectedIndex + 1
                lblBtw9.Text = GegevensItem.btw & " %"
                txtAantal9.Enabled = True
            Else
                txtAantal9.Enabled = False
            End If
            lblTotEx9.Text = Integer.Parse(txtAantal9.Text) * Double.Parse(lblPrijs9.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub cbOms10_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOms10.SelectedIndexChanged
        If reset = False Then
            If cbOms10.SelectedItem.ToString = cbOms9.SelectedItem.ToString Then
                MsgBox("Je kan niet 2 dezelfde Items selecteren")
                cbOms10.SelectedIndex = -1
            End If

            If Not cbOms10.SelectedIndex = -1 Then
                Dim SelectedItem As String = cbOms10.SelectedItem
                Dim GegevensItem As Artikel = DBmanager.SelectedItem(SelectedItem)
                lblPrijs10.Text = "€ " & GegevensItem.Eenheidsprijs
                lblArtId10.Text = cbOms10.SelectedIndex + 1
                lblBtw10.Text = GegevensItem.btw & " %"
                txtAantal10.Enabled = True
            Else
                txtAantal10.Enabled = False
            End If
            lblTotEx10.Text = Integer.Parse(txtAantal10.Text) * Double.Parse(lblPrijs10.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub txtAantal1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAantal1.TextChanged
        If Not txtAantal1.Text = String.Empty And Not lblPrijs1.Text = String.Empty Then
            lblTotEx1.Text = Integer.Parse(txtAantal1.Text) * Double.Parse(lblPrijs1.Text.Replace("€", ""))

            UpdateUitkomst()
        End If
    End Sub

    Private Sub txtAantal2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAantal2.TextChanged
        If Not txtAantal2.Text = String.Empty And Not lblPrijs2.Text = String.Empty Then
            lblTotEx2.Text = Integer.Parse(txtAantal2.Text) * Double.Parse(lblPrijs2.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub txtAantal3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAantal3.TextChanged
        If Not txtAantal3.Text = String.Empty And Not lblPrijs3.Text = String.Empty Then
            lblTotEx3.Text = Integer.Parse(txtAantal3.Text) * Double.Parse(lblPrijs3.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub txtAantal4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAantal4.TextChanged
        If Not txtAantal4.Text = String.Empty And Not lblPrijs4.Text = String.Empty Then
            lblTotEx4.Text = Integer.Parse(txtAantal4.Text) * Double.Parse(lblPrijs4.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub txtAantal5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAantal5.TextChanged
        If Not txtAantal5.Text = String.Empty And Not lblPrijs5.Text = String.Empty Then
            lblTotEx5.Text = Integer.Parse(txtAantal5.Text) * Double.Parse(lblPrijs5.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub txtAantal6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAantal6.TextChanged
        If Not txtAantal6.Text = String.Empty And Not lblPrijs6.Text = String.Empty Then
            lblTotEx6.Text = Integer.Parse(txtAantal6.Text) * Double.Parse(lblPrijs6.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub txtAantal7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAantal7.TextChanged
        If Not txtAantal7.Text = String.Empty And Not lblPrijs7.Text = String.Empty Then
            lblTotEx7.Text = Integer.Parse(txtAantal7.Text) * Double.Parse(lblPrijs7.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub txtAantal8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAantal8.TextChanged
        If Not txtAantal8.Text = String.Empty And Not lblPrijs8.Text = String.Empty Then
            lblTotEx8.Text = Integer.Parse(txtAantal8.Text) * Double.Parse(lblPrijs8.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub txtAantal9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAantal9.TextChanged
        If Not txtAantal9.Text = String.Empty And Not lblPrijs9.Text = String.Empty Then
            lblTotEx9.Text = Integer.Parse(txtAantal9.Text) * Double.Parse(lblPrijs9.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub txtAantal10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAantal10.TextChanged
        If Not txtAantal10.Text = String.Empty And Not lblPrijs10.Text = String.Empty Then
            lblTotEx10.Text = Integer.Parse(txtAantal1.Text) * Double.Parse(lblPrijs10.Text.Replace("€", ""))
            UpdateUitkomst()
        End If
    End Sub

    Private Sub btnOpslaan_Click(sender As System.Object, e As System.EventArgs) Handles btnOpslaan.Click
        ErrorProvider1.Clear()
        For Each tb In New TextBox() {txtAantal1, txtAantal2, txtAantal3, txtAantal4, txtAantal5, txtAantal6, txtAantal7, txtAantal8, txtAantal9, txtAantal10}
            If tb.Enabled = True Then
                If tb.Text = String.Empty Then
                    ErrorProvider1.SetError(tb, "Vul in")      'if empty -> error
                End If
            End If
        Next

        If cbKlant.SelectedIndex = -1 Then
            ErrorProvider1.SetError(cbKlant, "Kies een klant")      'if empty -> error
        End If

        If cbOms1.SelectedIndex = -1 Then
            ErrorProvider1.SetError(cbOms1, "Kies Artikel")
        End If

        If cbOms2.SelectedIndex = -1 And txtAantal2.Enabled = True Then
            ErrorProvider1.SetError(cbOms2, "Kies Artikel")      'if empty -> error
        End If
        If cbOms3.SelectedIndex = -1 And txtAantal3.Enabled = True Then
            ErrorProvider1.SetError(cbOms3, "Kies Artikel")      'if empty -> error
        End If
        If Not cbOms4.SelectedIndex = -1 And txtAantal4.Enabled = True Then
            ErrorProvider1.SetError(cbOms4, "Kies Artikel")      'if empty -> error
        End If
        If cbOms5.SelectedIndex = -1 And txtAantal5.Enabled = True Then
            ErrorProvider1.SetError(cbOms5, "Kies Artikel")      'if empty -> error
        End If
        If cbOms6.SelectedIndex = -1 And txtAantal6.Enabled = True Then
            ErrorProvider1.SetError(cbOms6, "Kies Artikel")      'if empty -> error
        End If
        If cbOms7.SelectedIndex = -1 And txtAantal7.Enabled = True Then
            ErrorProvider1.SetError(cbOms7, "Kies Artikel")      'if empty -> error
        End If
        If cbOms8.SelectedIndex = -1 And txtAantal8.Enabled = True Then
            ErrorProvider1.SetError(cbOms8, "Kies Artikel")      'if empty -> error
        End If
        If cbOms9.SelectedIndex = -1 And txtAantal9.Enabled = True Then
            ErrorProvider1.SetError(cbOms9, "Kies Artikel")      'if empty -> error
        End If
        If cbOms10.SelectedIndex = -1 And txtAantal10.Enabled = True Then
            ErrorProvider1.SetError(cbOms10, "Kies Artikel")      'if empty -> error
        End If

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If ErrorProvider1.GetError(cbKlant) = String.Empty And ErrorProvider1.GetError(Me.txtAantal1) = String.Empty And ErrorProvider1.GetError(Me.txtAantal2) = String.Empty And ErrorProvider1.GetError(Me.txtAantal3) = String.Empty And ErrorProvider1.GetError(Me.txtAantal4) = String.Empty And ErrorProvider1.GetError(Me.txtAantal5) = String.Empty And ErrorProvider1.GetError(Me.txtAantal6) = String.Empty And ErrorProvider1.GetError(Me.txtAantal7) = String.Empty And ErrorProvider1.GetError(Me.txtAantal8) = String.Empty And ErrorProvider1.GetError(Me.txtAantal9) = String.Empty And ErrorProvider1.GetError(Me.txtAantal10) = String.Empty And ErrorProvider1.GetError(cbOms1) = String.Empty And ErrorProvider1.GetError(cbOms2) = String.Empty And ErrorProvider1.GetError(cbOms3) = String.Empty And ErrorProvider1.GetError(cbOms4) = String.Empty And ErrorProvider1.GetError(cbOms5) = String.Empty And ErrorProvider1.GetError(cbOms6) = String.Empty And ErrorProvider1.GetError(cbOms7) = String.Empty And ErrorProvider1.GetError(cbOms8) = String.Empty And ErrorProvider1.GetError(cbOms9) = String.Empty And ErrorProvider1.GetError(cbOms10) = String.Empty Then

            Dim aantal = 0
            Dim laatsteAantal = lblFacNr.Text 'vorige aantal facs voor controle

            If Not lblTotEx1.Text = "" Then
                ReDim Preserve FactuurArtikels(aantal)
                FactuurArtikels(aantal) = New FactuurArtikels(lblFacNr.Text, cbOms1.SelectedIndex + 1, txtAantal1.Text)
                aantal += 1
            End If
            If Not lblTotEx2.Text = "" Then
                ReDim Preserve FactuurArtikels(aantal)
                FactuurArtikels(aantal) = New FactuurArtikels(lblFacNr.Text, cbOms2.SelectedIndex + 1, txtAantal2.Text)
                aantal += 1
            End If
            If Not lblTotEx3.Text = "" Then
                ReDim Preserve FactuurArtikels(aantal)
                FactuurArtikels(aantal) = New FactuurArtikels(lblFacNr.Text, cbOms3.SelectedIndex + 1, txtAantal3.Text)
                aantal += 1
            End If
            If Not lblTotEx4.Text = "" Then
                ReDim Preserve FactuurArtikels(aantal)
                FactuurArtikels(aantal) = New FactuurArtikels(lblFacNr.Text, cbOms4.SelectedIndex + 1, txtAantal4.Text)
                aantal += 1
            End If
            If Not lblTotEx5.Text = "" Then
                ReDim Preserve FactuurArtikels(aantal)
                FactuurArtikels(aantal) = New FactuurArtikels(lblFacNr.Text, cbOms5.SelectedIndex + 1, txtAantal5.Text)
                aantal += 1
            End If
            If Not lblTotEx6.Text = "" Then
                ReDim Preserve FactuurArtikels(aantal)
                FactuurArtikels(aantal) = New FactuurArtikels(lblFacNr.Text, cbOms6.SelectedIndex + 1, txtAantal6.Text)
                aantal += 1
            End If
            If Not lblTotEx7.Text = "" Then
                ReDim Preserve FactuurArtikels(aantal)
                FactuurArtikels(aantal) = New FactuurArtikels(lblFacNr.Text, cbOms7.SelectedIndex + 1, txtAantal7.Text)
                aantal += 1
            End If
            If Not lblTotEx8.Text = "" Then
                ReDim Preserve FactuurArtikels(aantal)
                FactuurArtikels(aantal) = New FactuurArtikels(lblFacNr.Text, cbOms8.SelectedIndex + 1, txtAantal8.Text)
                aantal += 1
            End If
            If Not lblTotEx9.Text = "" Then
                ReDim Preserve FactuurArtikels(aantal)
                FactuurArtikels(aantal) = New FactuurArtikels(lblFacNr.Text, cbOms9.SelectedIndex + 1, txtAantal9.Text)
                aantal += 1
            End If
            If Not lblTotEx10.Text = "" Then
                ReDim Preserve FactuurArtikels(aantal)
                FactuurArtikels(aantal) = New FactuurArtikels(lblFacNr.Text, cbOms10.SelectedIndex + 1, txtAantal10.Text)
                aantal += 1
            End If

            Dim fac As New Factuur(lblFacNr.Text, lblFacDat.Text, dtp1.Text, cbKlant.SelectedIndex + 1, lblTeBetalen.Text)
            DBmanager.FactuurToevoegen(fac)
            For Each fat In FactuurArtikels
                DBmanager.FactuurGegevensToevoegen(fat)
            Next
            DBmanager.ShowFacturen()                        'onload voor facnr count
            lblFacNr.Text = dgvShowFacturen.RowCount()      ' ""

            If lblFacNr.Text > laatsteAantal Then           'als er meerdere facturen in zitten dan ervoor : gelukt
                MsgBox("Factuur is toegevoegd aan de databank!")
                ''''''' NAAR EXCEL ''''''

                ' Open specific Excel document  
                Dim oExcel As Object = CreateObject("Excel.Application")
                Dim oBook As Object = oExcel.Workbooks.Open(My.Application.Info.DirectoryPath + "/Resources/InvoiceDeftig.xlsx")
                Dim oSheet As Object = oBook.Worksheets(1)

                oExcel.DisplayAlerts = False

                ' Read particular Cell  
                'Dim cellValue As String = oSheet.Range("B3").Value
                ' Write particular Cell  
                oSheet.Range("D8").Value = lblFacDat.Text
                oSheet.Range("D9").Value = lblFacNr.Text
                oSheet.Range("D10").Value = dtp1.Text
                oSheet.Range("D11").Value = lblBtwNr.Text

                oSheet.Range("G7").Value = cbKlant.SelectedItem
                oSheet.Range("G8").Value = lblStraat.Text + " " + lblStraatNr.Text
                oSheet.Range("G10").Value = lblPostcode.Text + " " + lblGemeente.Text
                oSheet.Range("G11").Value = lblLand.Text

                oSheet.Range("B15").Value = txtAantal1.Text
                oSheet.Range("B16").Value = txtAantal2.Text
                oSheet.Range("B17").Value = txtAantal3.Text
                oSheet.Range("B18").Value = txtAantal4.Text
                oSheet.Range("B19").Value = txtAantal5.Text
                oSheet.Range("B20").Value = txtAantal6.Text
                oSheet.Range("B21").Value = txtAantal7.Text
                oSheet.Range("B22").Value = txtAantal8.Text
                oSheet.Range("B23").Value = txtAantal9.Text
                oSheet.Range("B24").Value = txtAantal10.Text

                oSheet.Range("C15").Value = lblArtId1.Text
                oSheet.Range("C16").Value = lblArtId2.Text
                oSheet.Range("C17").Value = lblArtId3.Text
                oSheet.Range("C18").Value = lblArtId4.Text
                oSheet.Range("C19").Value = lblArtId5.Text
                oSheet.Range("C20").Value = lblArtId6.Text
                oSheet.Range("C21").Value = lblArtId7.Text
                oSheet.Range("C22").Value = lblArtId8.Text
                oSheet.Range("C23").Value = lblArtId9.Text
                oSheet.Range("C24").Value = lblArtId10.Text

                oSheet.Range("D15").Value = cbOms1.SelectedItem
                oSheet.Range("D16").Value = cbOms2.SelectedItem
                oSheet.Range("D17").Value = cbOms3.SelectedItem
                oSheet.Range("D18").Value = cbOms4.SelectedItem
                oSheet.Range("D19").Value = cbOms5.SelectedItem
                oSheet.Range("D20").Value = cbOms6.SelectedItem
                oSheet.Range("D21").Value = cbOms7.SelectedItem
                oSheet.Range("D22").Value = cbOms8.SelectedItem
                oSheet.Range("D23").Value = cbOms9.SelectedItem
                oSheet.Range("D24").Value = cbOms10.SelectedItem



                oBook.SaveAs(My.Application.Info.DirectoryPath + "/Resources/facturen/Invoice" + lblFacNr.Text + ".xls", True)
                oExcel.Quit()
            Else
                MsgBox("Mislukt!")
            End If
            reset = True ' boolean zodat index changed niet getriggered word...
            ResetNewFac()
            reset = False ' terug op beginwaarde
        End If
    End Sub

    Private Sub UpdateUitkomst()
        Dim som As Double = 0
        If Not lblTotEx1.Text = "" Then
            som += Double.Parse(lblTotEx1.Text)
        End If
        If Not lblTotEx2.Text = "" Then
            som += Double.Parse(lblTotEx2.Text)
        End If
        If Not lblTotEx3.Text = "" Then
            som += Double.Parse(lblTotEx3.Text)
        End If
        If Not lblTotEx4.Text = "" Then
            som += Double.Parse(lblTotEx4.Text)
        End If
        If Not lblTotEx5.Text = "" Then
            som += Double.Parse(lblTotEx5.Text)
        End If
        If Not lblTotEx6.Text = "" Then
            som += Double.Parse(lblTotEx6.Text)
        End If
        If Not lblTotEx7.Text = "" Then
            som += Double.Parse(lblTotEx7.Text)
        End If
        If Not lblTotEx8.Text = "" Then
            som += Double.Parse(lblTotEx8.Text)
        End If
        If Not lblTotEx9.Text = "" Then
            som += Double.Parse(lblTotEx9.Text)
        End If
        If Not lblTotEx10.Text = "" Then
            som += Double.Parse(lblTotEx10.Text)
        End If

        lblUitkomstMvh.Text = som
    End Sub

    Private Sub ResetNewFac()
        cbOms10.SelectedIndex = -1
        cbOms9.SelectedIndex = -1
        cbOms8.SelectedIndex = -1
        cbOms7.SelectedIndex = -1
        cbOms6.SelectedIndex = -1
        cbOms5.SelectedIndex = -1
        cbOms4.SelectedIndex = -1
        cbOms3.SelectedIndex = -1
        cbOms2.SelectedIndex = -1
        cbOms1.SelectedIndex = -1

        txtAantal1.Text = String.Empty
        txtAantal2.Text = String.Empty
        txtAantal3.Text = String.Empty
        txtAantal4.Text = String.Empty
        txtAantal5.Text = String.Empty
        txtAantal6.Text = String.Empty
        txtAantal7.Text = String.Empty
        txtAantal8.Text = String.Empty
        txtAantal9.Text = String.Empty
        txtAantal10.Text = String.Empty

        lblArtId1.Text = String.Empty
        lblArtId2.Text = String.Empty
        lblArtId3.Text = String.Empty
        lblArtId4.Text = String.Empty
        lblArtId5.Text = String.Empty
        lblArtId6.Text = String.Empty
        lblArtId7.Text = String.Empty
        lblArtId8.Text = String.Empty
        lblArtId9.Text = String.Empty
        lblArtId10.Text = String.Empty
        lblTotEx1.Text = String.Empty
        lblTotEx2.Text = String.Empty
        lblTotEx3.Text = String.Empty
        lblTotEx4.Text = String.Empty
        lblTotEx5.Text = String.Empty
        lblTotEx6.Text = String.Empty
        lblTotEx8.Text = String.Empty
        lblTotEx9.Text = String.Empty
        lblTotEx10.Text = String.Empty
        lblPrijs1.Text = String.Empty
        lblPrijs2.Text = String.Empty
        lblPrijs3.Text = String.Empty
        lblPrijs4.Text = String.Empty
        lblPrijs5.Text = String.Empty
        lblPrijs6.Text = String.Empty
        lblPrijs7.Text = String.Empty
        lblPrijs8.Text = String.Empty
        lblPrijs9.Text = String.Empty
        lblPrijs10.Text = String.Empty
        lblBtw1.Text = String.Empty
        lblBtw2.Text = String.Empty
        lblBtw3.Text = String.Empty
        lblBtw4.Text = String.Empty
        lblBtw5.Text = String.Empty
        lblBtw6.Text = String.Empty
        lblBtw7.Text = String.Empty
        lblBtw8.Text = String.Empty
        lblBtw9.Text = String.Empty
        lblBtw10.Text = String.Empty
        lblUitkomstBedBtw.Text = String.Empty
        lbluitkombtw.Text = String.Empty
        lblUitkomstMvh.Text = String.Empty
        lblTeBetalen.Text = String.Empty

        txtAantal1.Enabled = False
        txtAantal2.Enabled = False
        txtAantal3.Enabled = False
        txtAantal4.Enabled = False
        txtAantal5.Enabled = False
        txtAantal6.Enabled = False
        txtAantal7.Enabled = False
        txtAantal8.Enabled = False
        txtAantal9.Enabled = False
        txtAantal10.Enabled = False
    End Sub

    Private Sub lblUitkomstMvh_TextChanged(sender As System.Object, e As System.EventArgs) Handles lblUitkomstMvh.TextChanged
        If Not lblUitkomstMvh.Text = String.Empty Then
            Dim btw As Double = Double.Parse(lblUitkomstMvh.Text) * 21 / 100
            lblUitkomstBedBtw.Text = btw
            lblTeBetalen.Text = Format(btw + Double.Parse(lblUitkomstMvh.Text), "####0.00") ' "€ " &
        End If
    End Sub

    Private Sub dgvShowFacturen_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvShowFacturen.CellClick
        If IsNumeric(dgvShowFacturen.CurrentRow.Cells("id").Value) Then
            Dim selectedF As Integer = dgvShowFacturen.CurrentRow.Cells("id").Value
            'MsgBox(selectedF.ToString)
            DBmanager.SelectedFactuur(selectedF)
        End If
    End Sub

    Private Sub btnToevoegen_Click(sender As System.Object, e As System.EventArgs) Handles btnToevoegen.Click
        Dim k As Klant = New Klant(txtNewKlantID.Text, txtNewBedrijfsnaam.Text, txtNewStraat.Text, txtNewNummer.Text, txtNewPostcode.Text, txtNewPlaats.Text, txtNewLand.Text, txtNewBTWnr.Text)
        DBmanager.KlantGegevensToevoegen(k)
        DBmanager.ShowKlanten()
        txtNewKlantID.Text = dgvShowKlanten.RowCount() + 1
    End Sub

    Private Sub dgvEditKlanten_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEditKlanten.CellEndEdit
        Dim klantID As String = dgvEditKlanten.Item(0, selectedRow).Value.ToString
        Dim Bedrijfsnaam As String = dgvEditKlanten.Item(1, selectedRow).Value.ToString
        Dim Straat As String = dgvEditKlanten.Item(2, selectedRow).Value.ToString
        Dim Nummer As String = dgvEditKlanten.Item(3, selectedRow).Value.ToString
        Dim Postcode As String = dgvEditKlanten.Item(4, selectedRow).Value.ToString
        Dim Plaats As String = dgvEditKlanten.Item(5, selectedRow).Value.ToString
        Dim Land As String = dgvEditKlanten.Item(6, selectedRow).Value.ToString
        Dim Btwnr As String = dgvEditKlanten.Item(7, selectedRow).Value.ToString
        If dgvEditKlanten.CurrentCell.Value.ToString.Equals(String.Empty) Then
        Else
            DBmanager.UpdateDB(klantID, Bedrijfsnaam, Straat, Nummer, Postcode, Plaats, Land, Btwnr)
        End If

    End Sub
    Dim selectedRow As Integer
    Private Sub dgvEditKlanten_CellEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEditKlanten.CellEnter
        selectedRow = dgvEditKlanten.CurrentRow.Index
    End Sub

    Private Sub dgvEditKlanten_UserDeletingRow(sender As System.Object, e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvEditKlanten.UserDeletingRow
        selectedRow += 1
        DBmanager.DeleteRowKlant(selectedRow)
    End Sub

    Private Sub btnDeleteKlantIn_Click(sender As System.Object, e As System.EventArgs) Handles btnDeleteKlantIn.Click
        selectedRow += 1
        DBmanager.DeleteRowKlant(selectedRow)
        DBmanager.ShowKlanten()
    End Sub

    Private Sub btnAddklantIn_Click(sender As System.Object, e As System.EventArgs) Handles btnAddklantIn.Click
        DBmanager.ShowKlanten()
        txtNewKlantID.Text = dgvShowKlanten.RowCount()
        TabControl1.SelectedTab = TabNewKlant
    End Sub

    Private Sub BtnResetNewFac_Click(sender As System.Object, e As System.EventArgs) Handles btnResetNewFac.Click
        ResetNewFac()
        ErrorProvider1.Clear()
    End Sub

    Private Sub dgvShowFacturen_CellEnter(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvShowFacturen.CellEnter
        selectedRow = dgvShowFacturen.CurrentRow.Index
    End Sub

    Private Sub dgvShowFacturen_UserDeletingRow(sender As System.Object, e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvShowFacturen.UserDeletingRow
        selectedRow += 1
        DBmanager.DeleteRowFacGeg(selectedRow)
        DBmanager.DeleteRowFac(selectedRow)
    End Sub

    Private Sub dgvPrijzen_CellValueChanged(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPrijzen.CellValueChanged
        Dim ArtikelUpd As Artikel

        Dim ArtikelID As String = dgvPrijzen.Item(0, selectedRow).Value.ToString
        Dim Artikelomschrijving As String = dgvPrijzen.Item(1, selectedRow).Value.ToString
        Dim Eenheidsprijs As String = dgvPrijzen.Item(2, selectedRow).Value.ToString
        Dim BTW As String = dgvPrijzen.Item(3, selectedRow).Value.ToString

        If dgvPrijzen.CurrentCell.Value.ToString.Equals(String.Empty) Then
        Else
            ArtikelUpd = New Artikel(ArtikelID, Artikelomschrijving, Eenheidsprijs, BTW)
            DBmanager.UpdateDbArt(ArtikelUpd)
        End If
    End Sub


    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        selectedRow = dgvShowFacturen.CurrentRow.Index + 1
        Dim strFile As String = My.Application.Info.DirectoryPath + "\Resources\facturen\Invoice" + selectedRow.ToString + ".xls"
        Dim objProcess As New System.Diagnostics.ProcessStartInfo

        With objProcess
            .FileName = strFile
            .WindowStyle = ProcessWindowStyle.Hidden
            .Verb = "print"

            .CreateNoWindow = True
            .UseShellExecute = True
        End With
        Try
            System.Diagnostics.Process.Start(objProcess)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub CheckChar(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAantal1.KeyPress, txtAantal2.KeyPress, txtAantal3.KeyPress, txtAantal4.KeyPress, txtAantal5.KeyPress, txtAantal6.KeyPress, txtAantal7.KeyPress, txtAantal8.KeyPress, txtAantal9.KeyPress, txtAantal10.KeyPress
        If e.KeyChar <> ControlChars.Back Then
            If CType(sender, System.Windows.Forms.TextBox).Text.Contains(",") Or CType(sender, System.Windows.Forms.TextBox).Text.Contains(".") Then
                e.Handled = Not (Char.IsDigit(e.KeyChar))
            ElseIf CType(sender, System.Windows.Forms.TextBox).Text.Length = 0 Then
                e.Handled = Not (Char.IsDigit(e.KeyChar))
            ElseIf CType(sender, System.Windows.Forms.TextBox).Text.Length > 0 Then
                e.Handled = Not (Char.IsDigit(e.KeyChar) Or e.KeyChar = "," Or e.KeyChar = ".")
            End If
        End If
    End Sub

    Private Sub lblTotEx1_TextChanged(sender As System.Object, e As System.EventArgs) Handles lblTotEx1.TextChanged
        lblTotEx1.Text = FormatCurrency(lblTotEx1.Text, 3)
    End Sub

    Private Sub lblTotEx2_TextChanged(sender As System.Object, e As System.EventArgs) Handles lblTotEx2.TextChanged
        lblTotEx2.Text = FormatCurrency(lblTotEx2.Text, 3)
    End Sub

    Private Sub lblTotEx3_TextChanged(sender As System.Object, e As System.EventArgs) Handles lblTotEx3.TextChanged
        lblTotEx3.Text = FormatCurrency(lblTotEx3.Text, 3)
    End Sub

    Private Sub lblTotEx4_TextChanged(sender As System.Object, e As System.EventArgs) Handles lblTotEx4.TextChanged
        lblTotEx4.Text = FormatCurrency(lblTotEx4.Text, 3)
    End Sub

    Private Sub lblTotEx5_TextChanged(sender As System.Object, e As System.EventArgs) Handles lblTotEx5.TextChanged
        lblTotEx5.Text = FormatCurrency(lblTotEx5.Text, 3)
    End Sub

    Private Sub lblTotEx6_TextChanged(sender As System.Object, e As System.EventArgs) Handles lblTotEx6.TextChanged
        lblTotEx6.Text = FormatCurrency(lblTotEx6.Text, 3)
    End Sub

    Private Sub lblTotEx7_TextChanged(sender As System.Object, e As System.EventArgs) Handles lblTotEx7.TextChanged
        lblTotEx7.Text = FormatCurrency(lblTotEx7.Text, 3)
    End Sub

    Private Sub lblTotEx8_TextChanged(sender As System.Object, e As System.EventArgs) Handles lblTotEx8.TextChanged
        lblTotEx8.Text = FormatCurrency(lblTotEx8.Text, 3)
    End Sub

    Private Sub lblTotEx9_TextChanged(sender As System.Object, e As System.EventArgs) Handles lblTotEx9.TextChanged
        lblTotEx9.Text = FormatCurrency(lblTotEx9.Text, 3)
    End Sub

    Private Sub lblTotEx10_TextChanged(sender As System.Object, e As System.EventArgs) Handles lblTotEx10.TextChanged
        lblTotEx10.Text = FormatCurrency(lblTotEx10.Text, 3)
    End Sub
End Class