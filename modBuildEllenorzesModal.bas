Attribute VB_Name = "modBuildEllenorzesModal"
Option Explicit

' ============================================================
' modBuildEllenorzesModal
' Futtatás: Alt+F8 -> BuildEllenorzesModal_Run
' Feltételezi hogy frmEllenorzesModal már létezik.
' Kontrollokat törli és újra létrehozza teljes pozíció/design beállítással.
' EFJ szekció ELTÁVOLÍTVA – helyette btnEfjReszletek gomb.
' ============================================================

' --- Paletta ---
Private Const CLR_ACCENT    As Long = &H5C3317        ' sötét barna (brand)
Private Const CLR_ACCENT2   As Long = &H7A4A28        ' másodlagos barna
Private Const CLR_HDR_BG    As Long = &HF5F0EB        ' meleg homok – szekció háttér
Private Const CLR_HDR_FG    As Long = &H5C3317        ' szekció fejléc szöveg
Private Const CLR_SEP       As Long = &HAA8866        ' elválasztó vonal szín
Private Const CLR_RO_BG     As Long = &HF2F2F2        ' read-only mezo háttér
Private Const CLR_RO_FG     As Long = &H444444        ' read-only szöveg
Private Const CLR_INPUT_BG  As Long = &HFFFFFF        ' szerkesztheto mezo háttér
Private Const CLR_INPUT_FG  As Long = &H222222        ' szerkesztheto mezo szöveg
Private Const CLR_STATUS_BG As Long = &HCCE5FF        ' státusz mezohöz kék tónus
Private Const CLR_WHITE     As Long = &HFFFFFF
Private Const CLR_FORM_BG   As Long = &HFAFAFA

' --- Form méretek ---
Private Const FORM_W   As Integer = 700

' --- Fejléc zóna ---
Private Const KY_HDR   As Integer = 6
Private Const KY_SEP1  As Integer = 28    ' vízszintes elválasztó fejléc alatt

' --- Tartalom zóna ---
Private Const KY_BODY  As Integer = 36    ' tartalom kezdete
Private Const KX_L     As Integer = 10    ' bal oszlop X origó
Private Const LBL_W    As Integer = 122   ' bal oszlop label szélesség
Private Const TB_W     As Integer = 168   ' bal oszlop textbox szélesség
Private Const KX_L_TB  As Integer = KX_L + LBL_W + 2   ' = 134

Private Const KX_VDIV  As Integer = 316   ' függoleges elválasztó X
Private Const KX_R     As Integer = 324   ' jobb oszlop label X origó
Private Const RLBL_W   As Integer = 128   ' jobb oszlop label szélesség
Private Const RTB_W    As Integer = 212   ' jobb oszlop textbox szélesség
Private Const KX_R_TB  As Integer = KX_R + RLBL_W + 2  ' = 454

Private Const ROW_H    As Integer = 18
Private Const ROW_S    As Integer = 26    ' sor lépés (normál)
Private Const SEC_GAP  As Integer = 10   ' szekció közi gap
Private Const SEC_TH   As Integer = 16   ' szekció fejléc magasság
Private Const LBL_H    As Integer = 14   ' normál label magasság

' ============================================================
Public Sub BuildEllenorzesModal_Run()
    BuildControls
    MsgBox "frmEllenorzesModal kontrollok elkészültek.", vbInformation, "Build"
End Sub

' ============================================================
Private Sub BuildControls()
    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents("frmEllenorzesModal")
    'vbc.Properties("Width") = FORM_W + 12
    'vbc.Properties("BackColor") = CLR_FORM_BG
    
    Dim f As Object
    Set f = vbc.Designer

    ' -- Meglévo kontrollok törlése --
    Dim i As Integer
    For i = f.Controls.count - 1 To 0 Step -1
        f.Controls.Remove f.Controls(i).name
    Next i

    ' =========================================================
    ' FEJLÉC INFO SOR
    ' =========================================================
    With f.Controls.Add("Forms.Label.1", "lblStatusFix")
        .Left = KX_L: .Top = KY_HDR: .Width = 52: .Height = 14
        .Caption = "Státusz:"
        .Font.Bold = True: .Font.size = 9
        .ForeColor = CLR_ACCENT
        .BackColor = CLR_FORM_BG: .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", "txtStatus")
        .Left = KX_L + 54: .Top = KY_HDR - 1: .Width = 198: .Height = ROW_H
        .locked = True
        .BackColor = CLR_STATUS_BG: .ForeColor = CLR_RO_FG
        .Font.Bold = True: .Font.size = 9
        .BorderStyle = 0
    End With
    With f.Controls.Add("Forms.Label.1", "lblEfjAzonFix")
        .Left = KX_VDIV + 10: .Top = KY_HDR: .Width = 90: .Height = 14
        .Caption = "EFJ azonosító:"
        .Font.size = 9: .ForeColor = CLR_ACCENT
        .BackColor = CLR_FORM_BG: .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", "txtEfjAzonInfo")
        .Left = KX_VDIV + 102: .Top = KY_HDR - 1: .Width = 148: .Height = ROW_H
        .locked = True
        .BackColor = CLR_RO_BG: .ForeColor = CLR_RO_FG
        .Font.size = 9: .BorderStyle = 0
    End With

    ' -- Fejléc alatti vízszintes elválasztó --
    With f.Controls.Add("Forms.Label.1", "lblTopSep")
        .Left = 0: .Top = KY_SEP1: .Width = FORM_W: .Height = 2
        .Caption = "": .BackColor = CLR_ACCENT: .BackStyle = 1
    End With

    ' =========================================================
    ' FÜGGOLEGES ELVÁLASZTÓ (ideiglenesen rövid, Initialize végén méretezve)
    ' =========================================================
    With f.Controls.Add("Forms.Label.1", "lblVDivider")
        .Left = KX_VDIV: .Top = KY_BODY: .Width = 2: .Height = 250
        .Caption = "": .BackColor = CLR_SEP: .BackStyle = 1
    End With

    ' =========================================================
    ' BAL OSZLOP
    ' =========================================================
    Dim r As Integer: r = KY_BODY

    ' --- Megrendelo szekció fejléc ---
    With f.Controls.Add("Forms.Label.1", "lblMegrendeloHdr")
        .Left = KX_L: .Top = r: .Width = KX_L_TB + TB_W - KX_L: .Height = SEC_TH
        .Caption = "Megrendelo adatai"
        .Font.Bold = True: .Font.size = 9: .ForeColor = CLR_HDR_FG
        .BackColor = CLR_HDR_BG: .BackStyle = 1
    End With
    r = r + SEC_TH + 4

    AddLblTxt f, "lblVevokodFix", "txtVevokod", "Vevõkód", r, True
    r = r + ROW_S
    AddLblTxt f, "lblCegnevFix", "txtCegnev", "Cégnév", r, True
    r = r + ROW_S + SEC_GAP

    ' --- Kapcsolattartó szekció fejléc ---
    With f.Controls.Add("Forms.Label.1", "lblKapcsolattartoHdr")
        .Left = KX_L: .Top = r: .Width = KX_L_TB + TB_W - KX_L: .Height = SEC_TH
        .Caption = "Kapcsolattartó adatai"
        .Font.Bold = True: .Font.size = 9: .ForeColor = CLR_HDR_FG
        .BackColor = CLR_HDR_BG: .BackStyle = 1
    End With
    r = r + SEC_TH + 4

    AddLblTxt f, "lblKontakt", "txtKontakt", "Név", r, False
    r = r + ROW_S
    AddLblTxt f, "lblTelefon", "txtTelefon", "Telefonszám", r, False
    r = r + ROW_S
    AddLblTxt f, "lblEmail", "txtEmail", "E-mail", r, False
    r = r + ROW_S + SEC_GAP

    ' --- RAW / debug sorok (kis betuméret) ---
    With f.Controls.Add("Forms.Label.1", "lblRawAtveteliFix")
        .Left = KX_L: .Top = r: .Width = LBL_W: .Height = 12
        .Caption = "Op. Row Id:": .Font.size = 7: .ForeColor = RGB(160, 160, 160)
        .BackColor = CLR_FORM_BG: .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", "txtRawAtveteli")
        .Left = KX_L_TB: .Top = r - 1: .Width = TB_W: .Height = 14
        .locked = True: .Enabled = False
        .BackColor = CLR_RO_BG: .ForeColor = RGB(160, 160, 160)
        .Font.size = 7: .BorderStyle = 0
    End With
    r = r + 18

    With f.Controls.Add("Forms.Label.1", "lblRawIdokapuFix")
        .Left = KX_L: .Top = r: .Width = LBL_W: .Height = 12
        .Caption = "Proc Row Id:": .Font.size = 7: .ForeColor = RGB(160, 160, 160)
        .BackColor = CLR_FORM_BG: .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", "txtRawIdokapu")
        .Left = KX_L_TB: .Top = r - 1: .Width = TB_W: .Height = 14
        .locked = True: .Enabled = False
        .BackColor = CLR_RO_BG: .ForeColor = RGB(160, 160, 160)
        .Font.size = 7: .BorderStyle = 0
    End With
    r = r + 16

    ' =========================================================
    ' JOBB OSZLOP
    ' =========================================================
    Dim rR As Integer: rR = KY_BODY

    ' --- Felvételi adatok szekció fejléc ---
    With f.Controls.Add("Forms.Label.1", "lblFelvételHdr")
        .Left = KX_R: .Top = rR: .Width = KX_R_TB + RTB_W - KX_R: .Height = SEC_TH
        .Caption = "Küldemény felvételi adatai"
        .Font.Bold = True: .Font.size = 9: .ForeColor = CLR_HDR_FG
        .BackColor = CLR_HDR_BG: .BackStyle = 1
    End With
    rR = rR + SEC_TH + 4

    AddRLblTxt f, "lblFelrakoNev", "txtFelrakoNev", "Cég", rR, False
    rR = rR + ROW_S

    ' Irsz + Város egy sorban
    With f.Controls.Add("Forms.Label.1", "lblFelrakoIrsz")
        .Left = KX_R: .Top = rR: .Width = 62: .Height = LBL_H
        .Caption = "Irányítószám": .Font.size = 9: .ForeColor = CLR_RO_FG
        .BackColor = CLR_FORM_BG: .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", "txtFelrakoIrsz")
        .Left = KX_R_TB: .Top = rR - 2: .Width = 56: .Height = ROW_H
        .Font.size = 9: .BackColor = CLR_INPUT_BG: .ForeColor = CLR_INPUT_FG
    End With
    With f.Controls.Add("Forms.Label.1", "lblFelrakoVaros")
        .Left = KX_R_TB + 62: .Top = rR: .Width = 44: .Height = LBL_H
        .Caption = "Település": .Font.size = 9: .ForeColor = CLR_RO_FG
        .BackColor = CLR_FORM_BG: .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", "txtFelrakoVaros")
        .Left = KX_R_TB + 108: .Top = rR - 2: .Width = RTB_W - 108: .Height = ROW_H
        .Font.size = 9: .BackColor = CLR_INPUT_BG: .ForeColor = CLR_INPUT_FG
    End With
    rR = rR + ROW_S

    AddRLblTxt f, "lblKozteruletNev", "txtKozteruletNev", "Közterület neve", rR, False
    rR = rR + ROW_S

    ' KozteruletTipus + Hazszam egy sorban
    With f.Controls.Add("Forms.Label.1", "lblKozteruletTipus")
        .Left = KX_R: .Top = rR: .Width = RLBL_W: .Height = LBL_H
        .Caption = "Közterület típusa": .Font.size = 9: .ForeColor = CLR_RO_FG
        .BackColor = CLR_FORM_BG: .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", "txtKozteruletTipus")
        .Left = KX_R_TB: .Top = rR - 2: .Width = 88: .Height = ROW_H
        .Font.size = 9: .BackColor = CLR_INPUT_BG: .ForeColor = CLR_INPUT_FG
    End With
    With f.Controls.Add("Forms.Label.1", "lblHazszam")
        .Left = KX_R_TB + 94: .Top = rR: .Width = 44: .Height = LBL_H
        .Caption = "Házszám": .Font.size = 9: .ForeColor = CLR_RO_FG
        .BackColor = CLR_FORM_BG: .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", "txtHazszam")
        .Left = KX_R_TB + 140: .Top = rR - 2: .Width = RTB_W - 140: .Height = ROW_H
        .Font.size = 9: .BackColor = CLR_INPUT_BG: .ForeColor = CLR_INPUT_FG
    End With
    rR = rR + ROW_S

    AddRLblTxt f, "lblMegjegyzes", "txtMegjegyzes", "Megjegyzés", rR, False
    f.Controls("txtMegjegyzes").Height = 36
    f.Controls("txtMegjegyzes").MultiLine = True
    f.Controls("txtMegjegyzes").WordWrap = True
    rR = rR + 44

    AddRLblTxt f, "lblAtveteli", "txtAtveteli", "Átvétel dátuma", rR, False
    f.Controls("txtAtveteli").Width = 88
    rR = rR + ROW_S

    AddRLblTxt f, "lblIdokapu", "txtIdokapu", "Átvétel idõpontja", rR, False
    f.Controls("txtIdokapu").Width = 88
    rR = rR + ROW_S

    ' Függoleges elválasztó magasságát igazítjuk a magasabb oszlophoz
    Dim divH As Integer: divH = (IIf(r > rR, r, rR)) - KY_BODY
    f.Controls("lblVDivider").Height = divH + 4

    ' =========================================================
    ' KÜLDEMÉNY SZEKCIÓ
    ' =========================================================
    Dim sepY As Integer: sepY = (IIf(r > rR, r, rR)) + 8

    With f.Controls.Add("Forms.Label.1", "lblSepLine")
        .Left = 0: .Top = sepY: .Width = FORM_W: .Height = 2
        .Caption = "": .BackColor = CLR_ACCENT: .BackStyle = 1
    End With

    Dim kY As Integer: kY = sepY + 6

    With f.Controls.Add("Forms.Label.1", "lblKuldemenyHdr")
        .Left = KX_L: .Top = kY: .Width = FORM_W - 2 * KX_L: .Height = SEC_TH
        .Caption = "Küldemény adatok"
        .Font.Bold = True: .Font.size = 9: .ForeColor = CLR_HDR_FG
        .BackColor = CLR_HDR_BG: .BackStyle = 1
    End With
    kY = kY + SEC_TH + 2

    ' Fejléc labelek (felso sor)
    Dim kx As Integer: kx = KX_L
    AddKuldLbl f, "lblCsomagAzonFix", kx, kY, 188, "Csomagazonosító":     kx = kx + 192
    AddKuldLbl f, "lblHosszFix", kx, kY, 46, "H (cm)":                    kx = kx + 50
    AddKuldLbl f, "lblSzelesFix", kx, kY, 46, "Sz (cm)":                  kx = kx + 50
    AddKuldLbl f, "lblMagasFix", kx, kY, 46, "M (cm)":                    kx = kx + 50
    AddKuldLbl f, "lblEkaerFix", kx, kY, 54, "EKAER":                     kx = kx + 58
    AddKuldLbl f, "lblCimiratFix", kx, kY, 54, "Címirat":                 kx = kx + 58
    AddKuldLbl f, "lblAdrJeloloFix", kx, kY, 62, "ADR jelölõ":            kx = kx + 66
    AddKuldLbl f, "lblAdrPontFix", kx, kY, 52, "ADR pont":                kx = kx + 56
    AddKuldLbl f, "lblLerakasIdokapu", kx, kY, 88, "Lerakási idõkapu"
    kY = kY + 14

    ' Kontrollok (alsó sor)
    kx = KX_L
    With f.Controls.Add("Forms.TextBox.1", "txtCsomagAzon")
        .Left = kx: .Top = kY: .Width = 188: .Height = ROW_H
        .locked = True: .ForeColor = CLR_RO_FG
        .Font.Bold = True: .Font.size = 9: .BorderStyle = 0
        .BackColor = &HFFE8CC  ' narancs tónus – azonosító kiemelés
    End With: kx = kx + 192
    With f.Controls.Add("Forms.TextBox.1", "txtHossz")
        .Left = kx: .Top = kY: .Width = 46: .Height = ROW_H: .Font.size = 9
        .BackColor = CLR_INPUT_BG: .ForeColor = CLR_INPUT_FG
    End With: kx = kx + 50
    With f.Controls.Add("Forms.TextBox.1", "txtSzeles")
        .Left = kx: .Top = kY: .Width = 46: .Height = ROW_H: .Font.size = 9
        .BackColor = CLR_INPUT_BG: .ForeColor = CLR_INPUT_FG
    End With: kx = kx + 50
    With f.Controls.Add("Forms.TextBox.1", "txtMagas")
        .Left = kx: .Top = kY: .Width = 46: .Height = ROW_H: .Font.size = 9
        .BackColor = CLR_INPUT_BG: .ForeColor = CLR_INPUT_FG
    End With: kx = kx + 50
    With f.Controls.Add("Forms.TextBox.1", "txtEkaer")
        .Left = kx: .Top = kY: .Width = 50: .Height = ROW_H: .Font.size = 9
        .locked = True: .Enabled = False
        .BackColor = CLR_RO_BG: .ForeColor = CLR_RO_FG: .BorderStyle = 0
    End With: kx = kx + 58
    With f.Controls.Add("Forms.ComboBox.1", "cboCimirat")
        .Left = kx: .Top = kY: .Width = 50: .Height = ROW_H: .Font.size = 9
        .Style = 2
    End With: kx = kx + 58
    With f.Controls.Add("Forms.TextBox.1", "txtAdrJelolo")
        .Left = kx: .Top = kY: .Width = 60: .Height = ROW_H: .Font.size = 9
        .BackColor = CLR_INPUT_BG: .ForeColor = CLR_INPUT_FG
    End With: kx = kx + 66
    With f.Controls.Add("Forms.TextBox.1", "txtAdrPont")
        .Left = kx: .Top = kY: .Width = 50: .Height = ROW_H: .Font.size = 9
        .BackColor = CLR_INPUT_BG: .ForeColor = CLR_INPUT_FG
    End With: kx = kx + 56
    With f.Controls.Add("Forms.TextBox.1", "txtLerakasIdokapu")
        .Left = kx: .Top = kY: .Width = 86: .Height = ROW_H: .Font.size = 9
        .BackColor = CLR_INPUT_BG: .ForeColor = CLR_INPUT_FG
    End With
    kY = kY + ROW_H + 6

    ' =========================================================
    ' FOOTER
    ' =========================================================
    Dim btnY As Integer: btnY = kY + 6

    With f.Controls.Add("Forms.CheckBox.1", "chkEllenorizve")
        .Left = KX_L: .Top = btnY + 2: .Width = 150: .Height = 16
        .Caption = "Szállítónak átadható"
        .Font.Bold = True: .Font.size = 9
        .ForeColor = RGB(0, 128, 60): .BackColor = CLR_FORM_BG: .BackStyle = 1
    End With
    With f.Controls.Add("Forms.CommandButton.1", "btnEfjReszletek")
        .Left = 150: .Top = btnY: .Width = 100: .Height = 22
        .Caption = "EFJ részletek…": .Font.size = 9
        .BackColor = RGB(80, 100, 160): .ForeColor = CLR_WHITE
    End With
    With f.Controls.Add("Forms.CommandButton.1", "btnEfjEletut")
        .Left = 270: .Top = btnY: .Width = 100: .Height = 22
        .Caption = "Életút": .Font.size = 9
        .BackColor = RGB(80, 100, 160): .ForeColor = CLR_WHITE
    End With
    With f.Controls.Add("Forms.CommandButton.1", "btnStornoOrder")
        .Left = 398: .Top = btnY: .Width = 88: .Height = 22
        .Caption = "Stornó": .Font.size = 9
        .BackColor = RGB(185, 30, 30): .ForeColor = CLR_WHITE
    End With
    With f.Controls.Add("Forms.CommandButton.1", "btnMentes")
        .Left = 494: .Top = btnY: .Width = 88: .Height = 22
        .Caption = "Mentés": .Font.size = 9
        .BackColor = RGB(0, 128, 60): .ForeColor = CLR_WHITE
    End With
    With f.Controls.Add("Forms.CommandButton.1", "btnMegse")
        .Left = 590: .Top = btnY: .Width = 88: .Height = 22
        .Caption = "Mégsem": .Font.size = 9
        .BackColor = RGB(100, 100, 100): .ForeColor = CLR_WHITE
    End With

    ' -- Form teljes magassága --
    vbc.Properties("Height") = btnY + 62
End Sub

' ============================================================
' Helper: bal oszlop sor (label + textbox)
' ============================================================
Private Sub AddLblTxt(f As Object, lblName As String, tbName As String, _
                      cap As String, rowTop As Integer, isReadOnly As Boolean)
    With f.Controls.Add("Forms.Label.1", lblName)
        .Left = KX_L: .Top = rowTop: .Width = LBL_W: .Height = LBL_H
        .Caption = cap: .Font.size = 9: .ForeColor = CLR_RO_FG
        .BackColor = CLR_FORM_BG: .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", tbName)
        .Left = KX_L_TB: .Top = rowTop - 2: .Width = TB_W: .Height = ROW_H
        .Font.size = 9: .BorderStyle = 0
        If isReadOnly Then
            .locked = True: .BackColor = CLR_RO_BG: .ForeColor = CLR_RO_FG
        Else
            .BackColor = CLR_INPUT_BG: .ForeColor = CLR_INPUT_FG
        End If
    End With
End Sub

' ============================================================
' Helper: jobb oszlop sor
' ============================================================
Private Sub AddRLblTxt(f As Object, lblName As String, tbName As String, _
                       cap As String, rowTop As Integer, isReadOnly As Boolean)
    With f.Controls.Add("Forms.Label.1", lblName)
        .Left = KX_R: .Top = rowTop: .Width = RLBL_W: .Height = LBL_H
        .Caption = cap: .Font.size = 9: .ForeColor = CLR_RO_FG
        .BackColor = CLR_FORM_BG: .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", tbName)
        .Left = KX_R_TB: .Top = rowTop - 2: .Width = RTB_W: .Height = ROW_H
        .Font.size = 9: .BorderStyle = 0
        If isReadOnly Then
            .locked = True: .BackColor = CLR_RO_BG: .ForeColor = CLR_RO_FG
        Else
            .BackColor = CLR_INPUT_BG: .ForeColor = CLR_INPUT_FG
        End If
    End With
End Sub

' ============================================================
' Helper: küldemény szekció fejléc label
' ============================================================
Private Sub AddKuldLbl(f As Object, lblName As String, x As Integer, y As Integer, _
                       w As Integer, cap As String)
    With f.Controls.Add("Forms.Label.1", lblName)
        .Left = x: .Top = y: .Width = w: .Height = 12
        .Caption = cap: .Font.size = 8: .ForeColor = CLR_HDR_FG
        .BackColor = CLR_FORM_BG: .BackStyle = 1
    End With
End Sub


