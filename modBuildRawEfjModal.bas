Attribute VB_Name = "modBuildRawEfjModal"
' ============================================================
' modBuildRawEfjModal
' Feltételezi hogy frmRawEfjModal már létezik.
' Kontrollokat törli és újraépíti.
'
' Form: frmRawEfjModal
' Struktúra: Header + MultiPage (4 fül) + Footer
' Adatforrás: [RPA_Processes_DEV].[dbo].[csli_001_raw_efj]
' ============================================================
Option Explicit

' --- Témapaletta (zöld – CSS forrás: --g* változók) ---
Private Const CLR_G900      As Long = &H1A3D1A        ' --g900  sötét zöld (console bg)
Private Const CLR_G700      As Long = &H2D6A2D        ' --g700  elsodleges zöld (brand)
Private Const CLR_G600      As Long = &H3A7F3A        ' --g600  hover zöld
Private Const CLR_G500      As Long = &H4A9A4A        ' --g500  közepes zöld
Private Const CLR_G200      As Long = &HC5DFC5        ' --g200  halvány zöld border
Private Const CLR_G100      As Long = &HE8F4E8        ' --g100  nagyon halvány zöld
Private Const CLR_G050      As Long = &HF3FAF3        ' --g050  szinte fehér zöld
Private Const CLR_INK       As Long = &H1C2B1C        ' --ink   fo szöveg
Private Const CLR_INK2      As Long = &H4A614A        ' --ink2  másodlagos szöveg
Private Const CLR_INK3      As Long = &H8AA08A        ' --ink3  halványított szöveg
Private Const CLR_LINE      As Long = &HD0DDD0        ' --line  elválasztó / keret
Private Const CLR_FORM_BG   As Long = &HEEF3EE        ' --bg    form háttér
Private Const CLR_SURF      As Long = &HFFFFFF        ' --surf  felszín / fehér
Private Const CLR_RED       As Long = &H1111CC        ' --red   hiba piros (BGR!)
Private Const CLR_STATUS_BG As Long = &HCCE5FF        ' státusz mezo – kék tónus (változatlan)

' --- Layout konstansok ---
Const FORM_W            As Integer = 420
Const KX_LEFT           As Integer = 10
Const KY_HEADER         As Integer = 8
Const KY_SEP_TOP        As Integer = 30
Const KY_BODY           As Integer = 38      ' MultiPage teteje
Const KY_SEP_BOT        As Integer = 448
Const KY_FOOTER         As Integer = 458
Const LBL_FIX_W         As Integer = 110    ' fix felirat oszlop
Const LBL_VAL_W         As Integer = 230    ' érték oszlop
Const ROW_H             As Integer = 18     ' sormagasság
Const ROW_GAP           As Integer = 4      ' sorok közötti rés
Const MP_INNER_X        As Integer = 8      ' MultiPage belso bal margó
Const MP_INNER_Y        As Integer = 20     ' MultiPage belso felso margó (fül alatt)
Const FONT_SIZE_DEFAULT As Integer = 10

' ============================================================
' BELÉPÉSI PONT
' ============================================================
Public Sub BuildRawEfjModal_Run()
    BuildControls
    MsgBox "frmRawEfjModal kontrollok elkészültek.", vbInformation, "Build kész"
End Sub

' ============================================================
' FO BUILDER
' ============================================================
Private Sub BuildControls()
    Dim f As Object
    Set f = ThisWorkbook.VBProject.VBComponents("frmRawEfjModal").Designer
    f.Caption = ""
    
    ' Form méret és háttérszín
    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents("frmRawEfjModal")
    vbc.Properties("Width") = FORM_W + 12
    vbc.Properties("Height") = KY_FOOTER + 70
    vbc.Properties("BackColor") = CLR_FORM_BG
    vbc.Properties("Caption") = "EFJ adatok"

    ' 1. Meglévo kontrollok törlése
    Dim i As Integer
    For i = f.Controls.count - 1 To 0 Step -1
        f.Controls.Remove f.Controls(i).name
    Next i

    ' 2. HEADER
    BuildHeader f

    ' 3. SZEPARÁTOR — felso
    With f.Controls.Add("Forms.Label.1", "lblSepTop")
        .Left = 0
        .Top = KY_SEP_TOP
        .Width = FORM_W
        .Height = 2
        .BackColor = CLR_G700
        .Caption = ""
    End With

    ' 4. MULTIPAGE
    BuildMultiPage f

    ' 5. SZEPARÁTOR — alsó
    With f.Controls.Add("Forms.Label.1", "lblSepBot")
        .Left = 0
        .Top = KY_SEP_BOT
        .Width = FORM_W
        .Height = 2
        .BackColor = CLR_G700
        .Caption = ""
    End With

    ' 6. FOOTER
    BuildFooter f
End Sub

' ============================================================
' HEADER
' ============================================================
Private Sub BuildHeader(f As Object)
    ' Header háttérsáv
    With f.Controls.Add("Forms.Label.1", "lblHeaderBg")
        .Left = 0
        .Top = 0
        .Width = FORM_W
        .Height = KY_SEP_TOP
        .BackColor = CLR_G050
        .Caption = ""
    End With

    With f.Controls.Add("Forms.Label.1", "lblKuldemenyFix")
        .Left = KX_LEFT
        .Top = KY_HEADER
        .Width = 80
        .Height = ROW_H
        .Caption = "Küldemény:"
        .Font.Bold = True
        .Font.size = FONT_SIZE_DEFAULT
        .ForeColor = CLR_G700
        .BackColor = CLR_G050
        .BackStyle = 1
    End With

    With f.Controls.Add("Forms.Label.1", "lblKuldemenyVal")
        .Left = KX_LEFT + 84
        .Top = KY_HEADER
        .Width = 300
        .Height = ROW_H
        .Caption = ""              ' LoadRecord tölti
        .Font.Bold = True
        .Font.size = FONT_SIZE_DEFAULT
        .ForeColor = CLR_G700
        .BackColor = CLR_G050
        .BackStyle = 1
    End With
End Sub

' ============================================================
' MULTIPAGE + FÜLEK
' ============================================================
Private Sub BuildMultiPage(f As Object)
    With f.Controls.Add("Forms.MultiPage.1", "mpData")
        .Left = KX_LEFT
        .Top = KY_BODY
        .Width = FORM_W - 20
        .Height = KY_SEP_BOT - KY_BODY - 4
        .Pages(0).Caption = "Feladó"
        .Pages(1).Caption = "Küldemény"
        ' Pages(2) és (3) az Initialize-ban adandók hozzá — lásd vázlat lent
    End With
End Sub

' ============================================================
' FOOTER
' ============================================================
Private Sub BuildFooter(f As Object)
    ' Footer háttérsáv
    With f.Controls.Add("Forms.Label.1", "lblFooterBg")
        .Left = 0
        .Top = KY_SEP_BOT + 2
        .Width = FORM_W
        .Height = KY_FOOTER + 30
        .BackColor = CLR_G050
        .Caption = ""
    End With

    With f.Controls.Add("Forms.CommandButton.1", "btnBezaras")
        .Left = FORM_W - 90
        .Top = KY_FOOTER
        .Width = 80
        .Height = 22
        .Caption = "Bezárás"
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_G700
        .ForeColor = CLR_SURF
    End With
End Sub

' ============================================================
' USERFORM_INITIALIZE VÁZLAT
' Másold be a frmRawEfjModal kódmodulba
' ============================================================
'
' Private Sub UserForm_Initialize()
'     Me.BackColor = CLR_FORM_BG
'
'     Do While mpData.Pages.Count < 4
'         mpData.Pages.Add
'     Loop
'     mpData.Pages(0).Caption = "Feladó"
'     mpData.Pages(1).Caption = "Küldemény"
'     mpData.Pages(2).Caption = "Címzett"
'     mpData.Pages(3).Caption = "Audit"
'
'     SetupPageFelado
'     SetupPageKuldemeny
'     SetupPageCimzett
'     SetupPageAudit
' End Sub
'
' -------------------------------------------------------
' Segédeljárás: egy sor hozzáadása egy Page-hez
' -------------------------------------------------------
' Private Sub AddRow(pg As Object, rowIdx As Integer, _
'                    fixName As String, fixCaption As String, _
'                    valName As String)
'     Dim yPos As Integer
'     yPos = MP_INNER_Y + rowIdx * (ROW_H + ROW_GAP)
'
'     With pg.Controls.Add("Forms.Label.1", fixName)
'         .Left      = MP_INNER_X
'         .Top       = yPos
'         .Width     = LBL_FIX_W
'         .Height    = ROW_H
'         .Caption   = fixCaption
'         .Font.Size = FONT_SIZE_DEFAULT - 1
'         .ForeColor = CLR_INK2
'         .BackColor = CLR_FORM_BG
'         .BackStyle = 1
'     End With
'
'     With pg.Controls.Add("Forms.Label.1", valName)
'         .Left      = MP_INNER_X + LBL_FIX_W + 4
'         .Top       = yPos
'         .Width     = LBL_VAL_W
'         .Height    = ROW_H
'         .Caption   = ""
'         .Font.Size = FONT_SIZE_DEFAULT - 1
'         .ForeColor = CLR_INK
'         .BackColor = CLR_G050
'         .BackStyle = 1
'     End With
' End Sub
'
' -------------------------------------------------------
' FÜL 0: Feladó
' -------------------------------------------------------
' Private Sub SetupPageFelado()
'     Dim pg As Object: Set pg = mpData.Pages(0)
'     AddRow pg, 0,  "lblFelVevokodFix",       "Vevokód:",         "lblFelVevokodVal"
'     AddRow pg, 1,  "lblFelMegallapodasFix",   "Megállapodás:",    "lblFelMegallapodasVal"
'     AddRow pg, 2,  "lblFelNevFix",            "Név:",             "lblFelNevVal"
'     AddRow pg, 3,  "lblFelIrszFix",           "IRSZ:",            "lblFelIrszVal"
'     AddRow pg, 4,  "lblFelHelyFix",           "Helység:",         "lblFelHelyVal"
'     AddRow pg, 5,  "lblFelKozelebbiCimFix",   "Közelebbi cím:",   "lblFelKozelebbiCimVal"
'     AddRow pg, 6,  "lblFelKozteruletNevFix",  "Közt. név:",       "lblFelKozteruletNevVal"
'     AddRow pg, 7,  "lblFelKozteruletJelFix",  "Közt. jelleg:",    "lblFelKozteruletJelVal"
'     AddRow pg, 8,  "lblFelHazszamFix",        "Házszám:",         "lblFelHazszamVal"
'     AddRow pg, 9,  "lblFelMegjegyzesFix",     "Megjegyzés:",      "lblFelMegjegyzesVal"
'     AddRow pg, 10, "lblFelEmailFix",           "E-mail:",          "lblFelEmailVal"
'     AddRow pg, 11, "lblFelTelefonFix",         "Telefon:",         "lblFelTelefonVal"
' End Sub
'
' -------------------------------------------------------
' FÜL 1: Küldemény
' -------------------------------------------------------
' Private Sub SetupPageKuldemeny()
'     Dim pg As Object: Set pg = mpData.Pages(1)
'     AddRow pg, 0,  "lblSorszamFix",         "Sorszám:",          "lblSorszamVal"
'     AddRow pg, 1,  "lblAlapszolgFix",        "Alapszolg.:",       "lblAlapszolgVal"
'     AddRow pg, 2,  "lblSzolgaltatasFix",     "Szolgáltatás:",     "lblSzolgaltatasVal"
'     AddRow pg, 3,  "lblTomegFix",            "Tömeg (kg):",       "lblTomegVal"
'     AddRow pg, 4,  "lblKezbesitesModjaFix",  "Kézbesítés módja:", "lblKezbesitesModjaVal"
'     AddRow pg, 5,  "lblUtanvetFix",          "Utánvét (Ft):",     "lblUtanvetVal"
'     AddRow pg, 6,  "lblAruertekFix",         "Áruérték:",         "lblAruertekVal"
'     AddRow pg, 7,  "lblLerakoRefFix",        "Lerakó ref.:",      "lblLerakoRefVal"
'     AddRow pg, 8,  "lblVezerAzonFix",        "Vezér azon.:",      "lblVezerAzonVal"
'     AddRow pg, 9,  "lblUgyfadat1Fix",        "Ügyfadat 1:",       "lblUgyfadat1Val"
'     AddRow pg, 10, "lblUgyfadat2Fix",        "Ügyfadat 2:",       "lblUgyfadat2Val"
' End Sub
'
' -------------------------------------------------------
' FÜL 2: Címzett
' -------------------------------------------------------
' Private Sub SetupPageCimzett()
'     Dim pg As Object: Set pg = mpData.Pages(2)
'     AddRow pg, 0,  "lblCimNevFix",          "Név:",              "lblCimNevVal"
'     AddRow pg, 1,  "lblCimIrszFix",         "IRSZ:",             "lblCimIrszVal"
'     AddRow pg, 2,  "lblCimVarosFix",        "Város:",            "lblCimVarosVal"
'     AddRow pg, 3,  "lblCimKozelebbiCimFix", "Közelebbi cím:",    "lblCimKozelebbiCimVal"
'     AddRow pg, 4,  "lblCimKozteruletNevFix","Közt. név:",        "lblCimKozteruletNevVal"
'     AddRow pg, 5,  "lblCimKozteruletJelFix","Közt. jelleg:",     "lblCimKozteruletJelVal"
'     AddRow pg, 6,  "lblCimHazszamFix",      "Házszám:",          "lblCimHazszamVal"
'     AddRow pg, 7,  "lblCimMegjegyzesFix",   "Megjegyzés:",       "lblCimMegjegyzesVal"
'     AddRow pg, 8,  "lblCimKontaktFix",      "Kontakt személy:",  "lblCimKontaktVal"
'     AddRow pg, 9,  "lblCimEmailFix",        "E-mail:",           "lblCimEmailVal"
'     AddRow pg, 10, "lblCimTelefonFix",      "Telefon:",          "lblCimTelefonVal"
' End Sub
'
' -------------------------------------------------------
' FÜL 3: Audit
' -------------------------------------------------------
' Private Sub SetupPageAudit()
'     Dim pg As Object: Set pg = mpData.Pages(3)
'     AddRow pg, 0,  "lblIdFix",              "ID:",               "lblIdVal"
'     AddRow pg, 1,  "lblEfjIdFix",           "EFJ ID:",           "lblEfjIdVal"
'     AddRow pg, 2,  "lblEfjAzonFix",         "EFJ j.azon.:",      "lblEfjAzonVal"
'     AddRow pg, 3,  "lblBetoltesDatFix",     "Betöltés dátuma:",  "lblBetoltesDatVal"
'     AddRow pg, 4,  "lblEfjZarasFix",        "EFJ zárás:",        "lblEfjZarasVal"
'     AddRow pg, 5,  "lblEfjSzoftverFix",     "EFJ szoftver:",     "lblEfjSzoftverVal"
'     AddRow pg, 6,  "lblVarhatoDatFix",      "Várható feladás:",  "lblVarhatoDatVal"
'     AddRow pg, 7,  "lblCreatedAtFix",       "Létrehozva:",       "lblCreatedAtVal"
'     AddRow pg, 8,  "lblCreatedByFix",       "Létrehozta:",       "lblCreatedByVal"
'     AddRow pg, 9,  "lblModifiedAtFix",      "Módosítva:",        "lblModifiedAtVal"
'     AddRow pg, 10, "lblModifiedByFix",      "Módosította:",      "lblModifiedByVal"
'     AddRow pg, 11, "lblObsoleteFix",        "Obsolete:",         "lblObsoleteVal"
' End Sub

' ============================================================
' ADATBETÖLTÉSI VÁZLAT
' Hívás: frmRawEfjModal.LoadRecord "ABC123..."
' ============================================================
'
' Public Sub LoadRecord(kuldemenyAzon As String)
'     lblKuldemenyVal.Caption = kuldemenyAzon
'
'     Dim params(0) As Variant
'     params(0) = Array("@kuldemeny_azonosito", kuldemenyAzon)
'     Dim rs As ADODB.Recordset
'     Set rs = modConnection.ExecuteSP("usp_GetRawEfj", params)
'
'     If Not rs.EOF Then
'         ' Feladó fül
'         lblFelVevokodVal.Caption       = NZ(rs("felado_vevokod"))
'         lblFelMegallapodasVal.Caption  = NZ(rs("felado_megallapodas"))
'         lblFelNevVal.Caption           = NZ(rs("felado_nev"))
'         lblFelIrszVal.Caption          = NZ(rs("felado_irsz"))
'         lblFelHelyVal.Caption          = NZ(rs("felado_hely"))
'         lblFelKozelebbiCimVal.Caption  = NZ(rs("felado_kozelebbi_cim"))
'         lblFelKozteruletNevVal.Caption = NZ(rs("felado_kozterulet_nev"))
'         lblFelKozteruletJelVal.Caption = NZ(rs("felado_kozterulet_jelleg"))
'         lblFelHazszamVal.Caption       = NZ(rs("felado_hazszam"))
'         lblFelMegjegyzesVal.Caption    = NZ(rs("felado_megjegyzes"))
'         lblFelEmailVal.Caption         = NZ(rs("felado_email"))
'         lblFelTelefonVal.Caption       = NZ(rs("felado_telefon"))
'         ' Küldemény fül
'         lblSorszamVal.Caption          = NZ(rs("sorszam"))
'         lblAlapszolgVal.Caption        = NZ(rs("alapszolg"))
'         lblSzolgaltatasVal.Caption     = NZ(rs("szolgaltatas"))
'         lblTomegVal.Caption            = NZ(rs("kuldemeny_brutto_tomege_kg"))
'         lblKezbesitesModjaVal.Caption  = NZ(rs("kezbesites_modja"))
'         lblUtanvetVal.Caption          = NZ(rs("utanvet_osszege_ft"))
'         lblAruertekVal.Caption         = NZ(rs("kuldemeny_aruertek"))
'         lblLerakoRefVal.Caption        = NZ(rs("lerako_referencia_szam"))
'         lblVezerAzonVal.Caption        = NZ(rs("vezer_azon"))
'         lblUgyfadat1Val.Caption        = NZ(rs("ugyfadat1"))
'         lblUgyfadat2Val.Caption        = NZ(rs("ugyfadat2"))
'         ' Címzett fül
'         lblCimNevVal.Caption           = NZ(rs("cimzett_nev"))
'         lblCimIrszVal.Caption          = NZ(rs("cimzett_irsz"))
'         lblCimVarosVal.Caption         = NZ(rs("cimzett_varos"))
'         lblCimKozelebbiCimVal.Caption  = NZ(rs("cimzett_kozerlebbi_cim"))
'         lblCimKozteruletNevVal.Caption = NZ(rs("cimzett_kozterulet_nev"))
'         lblCimKozteruletJelVal.Caption = NZ(rs("cimzett_kozterulet_jelleg"))
'         lblCimHazszamVal.Caption       = NZ(rs("cimzett_hazszam"))
'         lblCimMegjegyzesVal.Caption    = NZ(rs("cimzett_megjegyzes"))
'         lblCimKontaktVal.Caption       = NZ(rs("cimzett_kontakt_szemely"))
'         lblCimEmailVal.Caption         = NZ(rs("cimzett_email"))
'         lblCimTelefonVal.Caption       = NZ(rs("cimzett_telefonszam"))
'         ' Audit fül
'         lblIdVal.Caption               = NZ(rs("id"))
'         lblEfjIdVal.Caption            = NZ(rs("efj_id"))
'         lblEfjAzonVal.Caption          = NZ(rs("efeladojegyzek_azon"))
'         lblBetoltesDatVal.Caption      = NZ(rs("betoltes_datum"))
'         lblEfjZarasVal.Caption         = NZ(rs("efj_zaras"))
'         lblEfjSzoftverVal.Caption      = NZ(rs("efj_szoftver"))
'         lblVarhatoDatVal.Caption       = NZ(rs("varhato_feladas_datum"))
'         lblCreatedAtVal.Caption        = NZ(rs("created_at"))
'         lblCreatedByVal.Caption        = NZ(rs("created_by"))
'         lblModifiedAtVal.Caption       = NZ(rs("modified_at"))
'         lblModifiedByVal.Caption       = NZ(rs("modified_by"))
'         lblObsoleteVal.Caption         = IIf(rs("obsolete") = True, "Igen", "Nem")
'         lblObsoleteVal.BackColor       = IIf(rs("obsolete") = True, CLR_STATUS_BG, CLR_G050)
'     End If
'     rs.Close
' End Sub
'
' Private Function NZ(val As Variant) As String
'     NZ = IIf(IsNull(val), "", CStr(val))
' End Function
'
' Private Sub btnBezaras_Click()
'     Unload Me
' End Sub


