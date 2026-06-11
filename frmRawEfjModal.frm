VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRawEfjModal 
   Caption         =   "EFJ adatok"
   ClientHeight    =   9975.001
   ClientLeft      =   -225
   ClientTop       =   -720
   ClientWidth     =   8400.001
   OleObjectBlob   =   "frmRawEfjModal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRawEfjModal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ============================================================
' frmRawEfjModal — EFJ nyers adat megtekinto (read-only modal)
'
' Elofeltétel: modBuildRawEfjModal.BuildRawEfjModal_Run() lefutott
' Hívás:       frmRawEfjModal.LoadRecord "ABC1234567890"
'              frmRawEfjModal.Show vbModal
' ============================================================
Option Explicit

' --- Layout konstansok (egyezzen a build scripttel) ---
Private Const MP_INNER_X        As Integer = 8
Private Const MP_INNER_Y        As Integer = 20
Private Const LBL_FIX_W         As Integer = 110
Private Const LBL_VAL_W         As Integer = 230
Private Const ROW_H             As Integer = 18
Private Const ROW_GAP           As Integer = 4
Private Const FONT_SIZE_DEFAULT As Integer = 9

Private Sub lblFooterBg_Click()

End Sub

' ============================================================
' USERFORM INITIALIZE
' ============================================================
Private Sub UserForm_Initialize()
    ' --- Fülek biztosítása (ha a build script csak 2-t hozott létre) ---
    Do While mpData.Pages.count < 4
        mpData.Pages.Add
    Loop
    mpData.Pages(0).Caption = "Feladó"
    mpData.Pages(1).Caption = "Küldemény"
    mpData.Pages(2).Caption = "Címzett"
    mpData.Pages(3).Caption = "Audit"

    ' --- Oldalankénti kontrollépítés ---
    SetupPageFelado
    SetupPageKuldemeny
    SetupPageCimzett
    SetupPageAudit
End Sub

' ============================================================
' SEGÉDELJÁRÁS — egy sor hozzáadása egy Page-hez
' fixName  ? Label  (fix felirat, nem interaktív)
' valName  ? TextBox readonly (kijelölheto, másolható, nem szerkesztheto)
' ============================================================
Private Sub AddRow(pg As Object, rowIdx As Integer, _
                   fixName As String, fixCaption As String, _
                   valName As String)
    Dim yPos As Integer
    yPos = MP_INNER_Y + rowIdx * (ROW_H + ROW_GAP)

    With pg.Controls.Add("Forms.Label.1", fixName)
        .Left = MP_INNER_X
        .Top = yPos
        .Width = LBL_FIX_W
        .Height = ROW_H
        .Caption = fixCaption
        .Font.size = FONT_SIZE_DEFAULT - 1
        .ForeColor = RGB(90, 90, 90)
    End With

    With pg.Controls.Add("Forms.TextBox.1", valName)
        .Left = MP_INNER_X + LBL_FIX_W + 4
        .Top = yPos
        .Width = LBL_VAL_W
        .Height = ROW_H
        .Font.size = FONT_SIZE_DEFAULT - 1
        .ForeColor = RGB(20, 20, 20)
        .BackColor = &H8000000F       ' rendszer ButtonFace — biztonságos, mindig elérheto
        .BorderStyle = 0              ' fmBorderStyleNone — label-szeru megjelenés
        .locked = True                ' nem szerkesztheto
        .TabStop = False              ' Tab navigációból kihagyva
        .Text = ""
    End With
End Sub

' ============================================================
' FÜL 0 — Feladó
' ============================================================
Private Sub SetupPageFelado()
    Dim pg As Object: Set pg = mpData.Pages(0)
    AddRow pg, 0, "lblFelVevokodFix", "Vevokód:", "txtFelVevokodVal"
    AddRow pg, 1, "lblFelMegallapodasFix", "Megállapodás:", "txtFelMegallapodasVal"
    AddRow pg, 2, "lblFelNevFix", "Név:", "txtFelNevVal"
    AddRow pg, 3, "lblFelIrszFix", "IRSZ:", "txtFelIrszVal"
    AddRow pg, 4, "lblFelHelyFix", "Helység:", "txtFelHelyVal"
    AddRow pg, 5, "lblFelKozelebbiCimFix", "Közelebbi cím:", "txtFelKozelebbiCimVal"
    AddRow pg, 6, "lblFelKozteruletNevFix", "Közt. név:", "txtFelKozteruletNevVal"
    AddRow pg, 7, "lblFelKozteruletJelFix", "Közt. jelleg:", "txtFelKozteruletJelVal"
    AddRow pg, 8, "lblFelHazszamFix", "Házszám:", "txtFelHazszamVal"
    AddRow pg, 9, "lblFelMegjegyzesFix", "Megjegyzés:", "txtFelMegjegyzesVal"
    AddRow pg, 10, "lblFelEmailFix", "E-mail:", "txtFelEmailVal"
    AddRow pg, 11, "lblFelTelefonFix", "Telefon:", "txtFelTelefonVal"
End Sub

' ============================================================
' FÜL 1 — Küldemény
' ============================================================
Private Sub SetupPageKuldemeny()
    Dim pg As Object: Set pg = mpData.Pages(1)
    AddRow pg, 0, "lblSorszamFix", "Sorszám:", "txtSorszamVal"
    AddRow pg, 1, "lblAlapszolgFix", "Alapszolg.:", "txtAlapszolgVal"
    AddRow pg, 2, "lblSzolgaltatasFix", "Szolgáltatás:", "txtSzolgaltatasVal"
    AddRow pg, 3, "lblTomegFix", "Tömeg (kg):", "txtTomegVal"
    AddRow pg, 4, "lblKezbesitesModjaFix", "Kézbesítés módja:", "txtKezbesitesModjaVal"
    AddRow pg, 5, "lblUtanvetFix", "Utánvét (Ft):", "txtUtanvetVal"
    AddRow pg, 6, "lblAruertekFix", "Áruérték:", "txtAruertekVal"
    AddRow pg, 7, "lblLerakoRefFix", "Lerakó ref.:", "txtLerakoRefVal"
    AddRow pg, 8, "lblVezerAzonFix", "Vezér azon.:", "txtVezerAzonVal"
    AddRow pg, 9, "lblUgyfadat1Fix", "Ügyfadat 1:", "txtUgyfadat1Val"
    AddRow pg, 10, "lblUgyfadat2Fix", "Ügyfadat 2:", "txtUgyfadat2Val"
End Sub

' ============================================================
' FÜL 2 — Címzett
' ============================================================
Private Sub SetupPageCimzett()
    Dim pg As Object: Set pg = mpData.Pages(2)
    AddRow pg, 0, "lblCimNevFix", "Név:", "txtCimNevVal"
    AddRow pg, 1, "lblCimIrszFix", "IRSZ:", "txtCimIrszVal"
    AddRow pg, 2, "lblCimVarosFix", "Város:", "txtCimVarosVal"
    AddRow pg, 3, "lblCimKozelebbiCimFix", "Közelebbi cím:", "txtCimKozelebbiCimVal"
    AddRow pg, 4, "lblCimKozteruletNevFix", "Közt. név:", "txtCimKozteruletNevVal"
    AddRow pg, 5, "lblCimKozteruletJelFix", "Közt. jelleg:", "txtCimKozteruletJelVal"
    AddRow pg, 6, "lblCimHazszamFix", "Házszám:", "txtCimHazszamVal"
    AddRow pg, 7, "lblCimMegjegyzesFix", "Megjegyzés:", "txtCimMegjegyzesVal"
    AddRow pg, 8, "lblCimKontaktFix", "Kontakt személy:", "txtCimKontaktVal"
    AddRow pg, 9, "lblCimEmailFix", "E-mail:", "txtCimEmailVal"
    AddRow pg, 10, "lblCimTelefonFix", "Telefon:", "txtCimTelefonVal"
End Sub

' ============================================================
' FÜL 3 — Audit
' ============================================================
Private Sub SetupPageAudit()
    Dim pg As Object: Set pg = mpData.Pages(3)
    AddRow pg, 0, "lblIdFix", "ID:", "txtIdVal"
    AddRow pg, 1, "lblEfjIdFix", "EFJ ID:", "txtEfjIdVal"
    AddRow pg, 2, "lblEfjAzonFix", "EFJ j.azon.:", "txtEfjAzonVal"
    AddRow pg, 3, "lblBetoltesDatFix", "Betöltés dátuma:", "txtBetoltesDatVal"
    AddRow pg, 4, "lblEfjZarasFix", "EFJ zárás:", "txtEfjZarasVal"
    AddRow pg, 5, "lblEfjSzoftverFix", "EFJ szoftver:", "txtEfjSzoftverVal"
    AddRow pg, 6, "lblVarhatoDatFix", "Várható feladás:", "txtVarhatoDatVal"
    AddRow pg, 7, "lblCreatedAtFix", "Létrehozva:", "txtCreatedAtVal"
    AddRow pg, 8, "lblCreatedByFix", "Létrehozta:", "txtCreatedByVal"
    AddRow pg, 9, "lblModifiedAtFix", "Módosítva:", "txtModifiedAtVal"
    AddRow pg, 10, "lblModifiedByFix", "Módosította:", "txtModifiedByVal"
    AddRow pg, 11, "lblObsoleteFix", "Obsolete:", "txtObsoleteVal"
End Sub

' ============================================================
' ADATBETÖLTÉS
' Hívás: frmRawEfjModal.LoadRecord "ABC1234567890"
' ============================================================
Public Sub LoadRecord(kuldemenyAzon As String)
    lblKuldemenyVal.Caption = kuldemenyAzon

    Dim params(0) As Variant
    params(0) = Array("@kuldemeny_azonosito_13", Left(kuldemenyAzon, 13))

    Dim rs As Object
    Set rs = modConnection.ExecuteSP("usp_GetRawEfj", params)

    If rs Is Nothing Then
        MsgBox "Kapcsolódási hiba — EFJ adat nem töltheto be.", vbExclamation, "Hiba"
        Exit Sub
    End If

    If rs.EOF Then
        lblKuldemenyVal.Caption = kuldemenyAzon & "  [nincs EFJ adat]"
        rs.Close
        Exit Sub
    End If

    ' --- Feladó fül ---
    mpData.Pages(0).Controls("txtFelVevokodVal").Text = Nz(rs("felado_vevokod"))
    mpData.Pages(0).Controls("txtFelMegallapodasVal").Text = Nz(rs("felado_megallapodas"))
    mpData.Pages(0).Controls("txtFelNevVal").Text = Nz(rs("felado_nev"))
    mpData.Pages(0).Controls("txtFelIrszVal").Text = Nz(rs("felado_irsz"))
    mpData.Pages(0).Controls("txtFelHelyVal").Text = Nz(rs("felado_hely"))
    mpData.Pages(0).Controls("txtFelKozelebbiCimVal").Text = Nz(rs("felado_kozelebbi_cim"))
    mpData.Pages(0).Controls("txtFelKozteruletNevVal").Text = Nz(rs("felado_kozterulet_nev"))
    mpData.Pages(0).Controls("txtFelKozteruletJelVal").Text = Nz(rs("felado_kozterulet_jelleg"))
    mpData.Pages(0).Controls("txtFelHazszamVal").Text = Nz(rs("felado_hazszam"))
    mpData.Pages(0).Controls("txtFelMegjegyzesVal").Text = Nz(rs("felado_megjegyzes"))
    mpData.Pages(0).Controls("txtFelEmailVal").Text = Nz(rs("felado_email"))
    mpData.Pages(0).Controls("txtFelTelefonVal").Text = Nz(rs("felado_telefon"))

    ' --- Küldemény fül ---
    mpData.Pages(1).Controls("txtSorszamVal").Text = Nz(rs("sorszam"))
    mpData.Pages(1).Controls("txtAlapszolgVal").Text = Nz(rs("alapszolg"))
    mpData.Pages(1).Controls("txtSzolgaltatasVal").Text = Nz(rs("szolgaltatas"))
    mpData.Pages(1).Controls("txtTomegVal").Text = Nz(rs("kuldemeny_brutto_tomege_kg"))
    mpData.Pages(1).Controls("txtKezbesitesModjaVal").Text = Nz(rs("kezbesites_modja"))
    mpData.Pages(1).Controls("txtUtanvetVal").Text = Nz(rs("utanvet_osszege_ft"))
    mpData.Pages(1).Controls("txtAruertekVal").Text = Nz(rs("kuldemeny_aruertek"))
    mpData.Pages(1).Controls("txtLerakoRefVal").Text = Nz(rs("lerako_referencia_szam"))
    mpData.Pages(1).Controls("txtVezerAzonVal").Text = Nz(rs("vezer_azon"))
    mpData.Pages(1).Controls("txtUgyfadat1Val").Text = Nz(rs("ugyfadat1"))
    mpData.Pages(1).Controls("txtUgyfadat2Val").Text = Nz(rs("ugyfadat2"))

    ' --- Címzett fül ---
    mpData.Pages(2).Controls("txtCimNevVal").Text = Nz(rs("cimzett_nev"))
    mpData.Pages(2).Controls("txtCimIrszVal").Text = Nz(rs("cimzett_irsz"))
    mpData.Pages(2).Controls("txtCimVarosVal").Text = Nz(rs("cimzett_varos"))
    mpData.Pages(2).Controls("txtCimKozelebbiCimVal").Text = Nz(rs("cimzett_kozerlebbi_cim"))
    mpData.Pages(2).Controls("txtCimKozteruletNevVal").Text = Nz(rs("cimzett_kozterulet_nev"))
    mpData.Pages(2).Controls("txtCimKozteruletJelVal").Text = Nz(rs("cimzett_kozterulet_jelleg"))
    mpData.Pages(2).Controls("txtCimHazszamVal").Text = Nz(rs("cimzett_hazszam"))
    mpData.Pages(2).Controls("txtCimMegjegyzesVal").Text = Nz(rs("cimzett_megjegyzes"))
    mpData.Pages(2).Controls("txtCimKontaktVal").Text = Nz(rs("cimzett_kontakt_szemely"))
    mpData.Pages(2).Controls("txtCimEmailVal").Text = Nz(rs("cimzett_email"))
    mpData.Pages(2).Controls("txtCimTelefonVal").Text = Nz(rs("cimzett_telefonszam"))

    ' --- Audit fül ---
    mpData.Pages(3).Controls("txtIdVal").Text = Nz(rs("id"))
    mpData.Pages(3).Controls("txtEfjIdVal").Text = Nz(rs("efj_id"))
    mpData.Pages(3).Controls("txtEfjAzonVal").Text = Nz(rs("efeladojegyzek_azon"))
    mpData.Pages(3).Controls("txtBetoltesDatVal").Text = Nz(rs("betoltes_datum"))
    mpData.Pages(3).Controls("txtEfjZarasVal").Text = Nz(rs("efj_zaras"))
    mpData.Pages(3).Controls("txtEfjSzoftverVal").Text = Nz(rs("efj_szoftver"))
    mpData.Pages(3).Controls("txtVarhatoDatVal").Text = Nz(rs("varhato_feladas_datum"))
    mpData.Pages(3).Controls("txtCreatedAtVal").Text = Nz(rs("created_at"))
    mpData.Pages(3).Controls("txtCreatedByVal").Text = Nz(rs("created_by"))
    mpData.Pages(3).Controls("txtModifiedAtVal").Text = Nz(rs("modified_at"))
    mpData.Pages(3).Controls("txtModifiedByVal").Text = Nz(rs("modified_by"))
    Dim obsText As String
    If rs("obsolete") = True Then
        obsText = "Igen"
    Else
        obsText = "Nem"
    End If
    mpData.Pages(3).Controls("txtObsoleteVal").Text = obsText

    rs.Close
End Sub

' ============================================================
' ESEMÉNYKEZELOK
' ============================================================
Private Sub btnBezaras_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' ESC / ablak X gomb ? ugyanaz mint Bezárás
    If CloseMode = vbFormControlMenu Then Unload Me
End Sub

' ============================================================
' SEGÉDFÜGGVÉNY
' ============================================================
Private Function Nz(val As Variant) As String
    If IsNull(val) Or IsEmpty(val) Then
        Nz = ""
    Else
        Nz = CStr(val)
    End If
End Function


