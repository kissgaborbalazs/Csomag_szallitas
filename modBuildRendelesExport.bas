Attribute VB_Name = "modBuildRendelesExport"
' ============================================================
' modBuildRendelesExport
' Feltételezi hogy frmRendelesExprt már létezik.
' Kontrollokat törli és újraépíti.
'
' Form: frmRendelesExprt
' SP:   usp_GetRendelesExportList
'
' Struktúra:
'   Header (szurok + gombok)
'   Szeparátor (felso)
'   ListView (MSComctlLib) — vízszintes görgetéssel
'   Szeparátor (alsó)
'   Footer (gombok)
'
' ListView oszlopok (38 látható + proc_data_id rejtett Tag-ben):
'   0  CheckBox
'   1  Küldemény azonosító      kuldemeny_azonosito
'   2  Vevokód                  megrendelo_vevokod
'   3  Áruérték                 kuldemeny_aruertek
'   4  Küldemény megj.          kuldemeny_megjegyzes
'   5  Felrakó neve             felrako_nev
'   6  Felrakó irsz             felrako_irsz
'   7  Felrakó város            felrako_varos
'   8  Felrakó közterület       felrako_kozterulet_elnevezese
'   9  Felrakó típus            felrako_kozterulet_tipus
'  10  Felrakó hsz              felrako_hazszam
'  11  Felrakó megj.            felrako_megjegyzes
'  12  Felrakó kontakt          felrako_kontakt_szemely
'  13  Felrakó tel.             felrako_telefonszam
'  14  Felrakó email            felrako_email
'  15  Felrakó idokapu          felrako_idokapu
'  16  Címzett neve             cimzett_nev
'  17  Címzett irsz             cimzett_irsz
'  18  Címzett város            cimzett_varos
'  19  Címzett közterület       cimzett_kozterulet_nev
'  20  Címzett jelleg           cimzett_kozterulet_jelleg
'  21  Címzett hsz              cimzett_hazszam
'  22  Címzett megj.            cimzett_megjegyzes
'  23  Címzett kontakt          cimzett_kontakt_szemely
'  24  Címzett tel.             cimzett_telefonszam
'  25  Címzett email            cimzett_email
'  26  Lerakás idokapu          lerakas_idokapu
'  27  Raklap hossza (cm)       raklap_hossza_cm
'  28  Raklap széless. (cm)     raklap_szelessege_cm
'  29  Raklap magass. (cm)      raklap_magassaga_cm
'  30  Bruttó tömeg (kg)        kuldemeny_brutto_tomege_kg
'  31  Utánvét (Ft)             utanvet_osszege_ft
'  32  EKAER                    ekaer
'  33  ADR jelölo               adr_jelolo
'  34  ADR pont                 adr_pont
'  35  Szolgáltatás             szolgaltatas
'  36  Lerakó ref.              lerako_referencia
'  37  Felvételi ref.           felveteli_referencia
'   —  (rejtett) proc_data_id  ? ListItem.Tag
' ============================================================
Option Explicit

' --- Témapaletta (zöld – BGR kódolás) ---
Private Const CLR_G900      As Long = &H1A3D1A
Private Const CLR_G700      As Long = &H2D6A2D
Private Const CLR_G600      As Long = &H3A7F3A
Private Const CLR_G500      As Long = &H4A9A4A
Private Const CLR_G200      As Long = &HC5DFC5
Private Const CLR_G100      As Long = &HE8F4E8
Private Const CLR_G050      As Long = &HF3FAF3
Private Const CLR_INK       As Long = &H1C2B1C
Private Const CLR_INK2      As Long = &H4A614A
Private Const CLR_INK3      As Long = &H8AA08A
Private Const CLR_LINE      As Long = &HD0DDD0
Private Const CLR_FORM_BG   As Long = &HEEF3EE
Private Const CLR_SURF      As Long = &HFFFFFF
Private Const CLR_RED       As Long = &H1111CC
Private Const CLR_ORANGE    As Long = &H6600CC  ' BGR: #CC0066 ? narancsos akció

' --- Layout konstansok ---
Private Const FORM_W            As Integer = 1400
Private Const KX_LEFT           As Integer = 8
Private Const KY_HEADER         As Integer = 8
Private Const KY_INPUT          As Integer = 22
Private Const KY_SEP_TOP        As Integer = 48
Private Const KY_LIST           As Integer = 54
Private Const KY_SEP_BOT        As Integer = 458
Private Const KY_FOOTER         As Integer = 466
Private Const LIST_H            As Integer = 400
Private Const INPUT_H           As Integer = 18
Private Const LBL_H             As Integer = 14
Private Const BTN_H             As Integer = 22
Private Const FONT_SIZE_DEFAULT As Integer = 9

' ============================================================
' BELÉPÉSI PONT
' ============================================================
Public Sub BuildRendelesExport_Run()
    BuildControls
    MsgBox "frmRendelesExprt kontrollok elkészültek.", vbInformation, "Build kész"
End Sub

' ============================================================
' FO BUILDER
' ============================================================
Private Sub BuildControls()
    Dim f   As Object
    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents("frmRendelesExprt")
    Set f = vbc.Designer

    ' Form méret és háttérszín
    vbc.Properties("Width") = FORM_W + 12
    vbc.Properties("Height") = KY_FOOTER + 60
    vbc.Properties("BackColor") = CLR_FORM_BG
    vbc.Properties("Caption") = "Rendelés export – átadásra váró küldemények"

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

    ' 4. LISTVIEW
    BuildListView f

    ' 5. SZEPARÁTOR — alsó
    With f.Controls.Add("Forms.Label.1", "lblSepBot")
        .Left = 0
        .Top = KY_SEP_BOT
        .Width = FORM_W
        .Height = 2
        .BackColor = CLR_LINE
        .Caption = ""
    End With

    ' 6. FOOTER
    BuildFooter f
End Sub

' ============================================================
' HEADER — szuro mezok és gombok
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

    ' --- Vevokód ---
    With f.Controls.Add("Forms.Label.1", "lblVevokodFix")
        .Left = KX_LEFT
        .Top = KY_HEADER
        .Width = 55
        .Height = LBL_H
        .Caption = "Vevõkód:"
        .Font.size = FONT_SIZE_DEFAULT
        .ForeColor = CLR_INK2
        .BackColor = CLR_G050
        .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", "txtVevokod")
        .Left = KX_LEFT
        .Top = KY_INPUT
        .Width = 90
        .Height = INPUT_H
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_SURF
        .ForeColor = CLR_INK
    End With

    ' --- Azonosító ---
    With f.Controls.Add("Forms.Label.1", "lblAzonFix")
        .Left = KX_LEFT + 98
        .Top = KY_HEADER
        .Width = 60
        .Height = LBL_H
        .Caption = "Azonosító:"
        .Font.size = FONT_SIZE_DEFAULT
        .ForeColor = CLR_INK2
        .BackColor = CLR_G050
        .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", "txtAzon")
        .Left = KX_LEFT + 98
        .Top = KY_INPUT
        .Width = 110
        .Height = INPUT_H
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_SURF
        .ForeColor = CLR_INK
    End With

    ' --- Dátumtól ---
    With f.Controls.Add("Forms.Label.1", "lblDatumFix")
        .Left = KX_LEFT + 216
        .Top = KY_HEADER
        .Width = 55
        .Height = LBL_H
        .Caption = "Dátumtól:"
        .Font.size = FONT_SIZE_DEFAULT
        .ForeColor = CLR_INK2
        .BackColor = CLR_G050
        .BackStyle = 1
    End With
    With f.Controls.Add("Forms.TextBox.1", "txtDatum")
        .Left = KX_LEFT + 216
        .Top = KY_INPUT
        .Width = 80
        .Height = INPUT_H
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_SURF
        .ForeColor = CLR_INK
    End With

    ' --- Szurés gomb ---
    With f.Controls.Add("Forms.CommandButton.1", "btnSzures")
        .Left = KX_LEFT + 308
        .Top = KY_INPUT - 3
        .Width = 55
        .Height = BTN_H
        .Caption = "Szûrés"
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_G500
        .ForeColor = CLR_SURF
    End With

    ' --- Törlés gomb ---
    With f.Controls.Add("Forms.CommandButton.1", "btnTorol")
        .Left = KX_LEFT + 369
        .Top = KY_INPUT - 3
        .Width = 55
        .Height = BTN_H
        .Caption = "Törlés"
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_G050
        .ForeColor = CLR_INK2
    End With

    ' --- Elválasztó: szurok / akció gombok ---
    With f.Controls.Add("Forms.Label.1", "lblHeaderDivider")
        .Left = KX_LEFT + 436
        .Top = KY_HEADER
        .Width = 2
        .Height = BTN_H + LBL_H + 2
        .BackColor = CLR_LINE
        .Caption = ""
    End With

    ' --- Mind kijelöl ---
    With f.Controls.Add("Forms.CommandButton.1", "btnSelectAll")
        .Left = KX_LEFT + 446
        .Top = KY_INPUT - 3
        .Width = 80
        .Height = BTN_H
        .Caption = "Mind kijelöl"
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_G100
        .ForeColor = CLR_INK2
    End With

    ' --- Kijelölés törlése ---
    With f.Controls.Add("Forms.CommandButton.1", "btnSelectNone")
        .Left = KX_LEFT + 532
        .Top = KY_INPUT - 3
        .Width = 80
        .Height = BTN_H
        .Caption = "Kijelölés törl."
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_G100
        .ForeColor = CLR_INK2
    End With

    ' --- Kijelölve: N db ---
    With f.Controls.Add("Forms.Label.1", "lblKijeloltDb")
        .Left = KX_LEFT + 620
        .Top = KY_INPUT
        .Width = 120
        .Height = LBL_H
        .Caption = "Kijelölve: 0 db"
        .Font.size = FONT_SIZE_DEFAULT
        .ForeColor = CLR_INK2
        .BackColor = CLR_G050
        .BackStyle = 1
    End With

    ' --- Export Excel gomb ---
    With f.Controls.Add("Forms.CommandButton.1", "btnExportExcel")
        .Left = KX_LEFT + 748
        .Top = KY_INPUT - 3
        .Width = 110
        .Height = BTN_H
        .Caption = "Export Excelbe"
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_G700
        .ForeColor = CLR_SURF
        .Enabled = False
    End With

    ' --- Átadva (kiszállítónak) gomb ---
    With f.Controls.Add("Forms.CommandButton.1", "btnAtadva")
        .Left = KX_LEFT + 864
        .Top = KY_INPUT - 3
        .Width = 130
        .Height = BTN_H
        .Caption = "Átadva (kiszállítónak)"
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_ORANGE
        .ForeColor = CLR_SURF
        .Enabled = False
    End With
End Sub

' ============================================================
' LISTVIEW
' Csak pozíció és méret — oszlopok frmRendelesExprt.Initialize-ban!
' ============================================================
Private Sub BuildListView(f As Object)
    f.Controls.Add "MSComctlLib.ListViewCtrl.2", "lstExport"
    With f.Controls("lstExport")
        .Left = KX_LEFT
        .Top = KY_LIST
        .Width = FORM_W - KX_LEFT * 2
        .Height = LIST_H
    End With
End Sub

' ============================================================
' FOOTER
' ============================================================
Private Sub BuildFooter(f As Object)
    With f.Controls.Add("Forms.Label.1", "lblFooterBg")
        .Left = 0
        .Top = KY_SEP_BOT + 2
        .Width = FORM_W
        .Height = 36
        .BackColor = CLR_G050
        .Caption = ""
    End With

    With f.Controls.Add("Forms.CommandButton.1", "btnBezaras")
        .Left = FORM_W - 90
        .Top = KY_FOOTER
        .Width = 80
        .Height = BTN_H
        .Caption = "Bezárás"
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_G700
        .ForeColor = CLR_SURF
    End With
End Sub
