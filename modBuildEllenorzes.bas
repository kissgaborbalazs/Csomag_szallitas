Attribute VB_Name = "modBuildEllenorzes"
' ============================================================
' modBuildEllenorzes
' Feltételezi hogy frmEllenorzes már létezik.
' Kontrollokat törli és újraépíti.
'
' Form: frmEllenorzes
' Struktúra:
'   Header (szurok + gombok)
'   Szeparátor (felso)
'   ListView (MSComctlLib) — 25 oszlop, vízszintes görgetéssel
'   Szeparátor (alsó)
'   Footer (gombok)
'
' ListView oszlopok:
'   0  ?  (CheckBox szimuláció — Item.Text)
'   1  Küldemény azonosító
'   2  Vevokód
'   3  Áruérték
'   4  Felrakó neve
'   5  Felrakó cím           (irsz + varos + kozterulet + tipus + hazszam összefuzve)
'   6  Felrakó megjegyzés
'   7  Felrakó kontakt
'   8  Felrakó tel.
'   9  Felrakó email
'  10  Felrakó idokapu
'  11  Címzett neve
'  12  Címzett cím           (irsz + varos + kozterulet_nev + jelleg + hazszam összefuzve)
'  13  Címzett megjegyzés
'  14  Címzett kontakt
'  15  Címzett tel.
'  16  Címzett email
'  17  Lerakás idokapu
'  18  Raklap H×Sz×M (cm)   (hossza + szelessege + magassaga összefuzve)
'  19  Bruttó tömeg (kg)
'  20  Utánvét (Ft)
'  21  EKAER
'  22  ADR jelölo
'  23  ADR pont
'  24  Szolgáltatás
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
Private Const CLR_ORANGE    As Long = &H66CC     ' átadhatónak jelöl gomb (BGR: #CC6600)

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
Public Sub BuildEllenorzes_Run()
    BuildControls
    MsgBox "frmEllenorzes kontrollok elkészültek.", vbInformation, "Build kész"
End Sub

' ============================================================
' FO BUILDER
' ============================================================
Private Sub BuildControls()
    Dim f   As Object
    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents("frmEllenorzes")
    Set f = vbc.Designer

    ' Form méret és háttérszín
    vbc.Properties("Width") = FORM_W + 12
    vbc.Properties("Height") = KY_FOOTER + 60
    vbc.Properties("BackColor") = CLR_FORM_BG
    vbc.Properties("Caption") = "Ellenõrzés – szállítónak átadásra váró küldemények"

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

    ' --- Dátum ---
    With f.Controls.Add("Forms.Label.1", "lblDatumFix")
        .Left = KX_LEFT + 216
        .Top = KY_HEADER
        .Width = 50
        .Height = LBL_H
        .Caption = "Dátum:"
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
        .Top = KY_INPUT - 1
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
        .Top = KY_INPUT - 1
        .Width = 55
        .Height = BTN_H
        .Caption = "Törlés"
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_G050
        .ForeColor = CLR_INK2
    End With

    ' --- Elválasztó a szurok és akció gombok között ---
    With f.Controls.Add("Forms.Label.1", "lblHeaderDivider")
        .Left = KX_LEFT + 436
        .Top = KY_HEADER
        .Width = 2
        .Height = BTN_H + LBL_H + 2
        .BackColor = CLR_LINE
        .Caption = ""
    End With

    ' --- Összes kijelölése / törlése ---
    With f.Controls.Add("Forms.CommandButton.1", "btnSelectAll")
        .Left = KX_LEFT + 446
        .Top = KY_INPUT - 1
        .Width = 80
        .Height = BTN_H
        .Caption = "Mind kijelöl"
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_G100
        .ForeColor = CLR_INK2
    End With
    With f.Controls.Add("Forms.CommandButton.1", "btnSelectNone")
        .Left = KX_LEFT + 532
        .Top = KY_INPUT - 1
        .Width = 80
        .Height = BTN_H
        .Caption = "Kijelölés törl."
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_G100
        .ForeColor = CLR_INK2
    End With

    ' --- Kijelöltek száma label ---
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

    ' --- Átadhatónak jelöl gomb ---
    With f.Controls.Add("Forms.CommandButton.1", "btnAtadhato")
        .Left = KX_LEFT + 748
        .Top = KY_INPUT - 1
        .Width = 120
        .Height = BTN_H
        .Caption = "Átadhatónak jelöl"
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_ORANGE
        .ForeColor = CLR_SURF
        .Enabled = False
    End With

    ' --- Stornó gomb ---
    With f.Controls.Add("Forms.CommandButton.1", "btnStornoSelected")
        .Left = KX_LEFT + 874
        .Top = KY_INPUT - 1
        .Width = 90
        .Height = BTN_H
        .Caption = "Stornó"
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_RED
        .ForeColor = CLR_SURF
        .Enabled = False
    End With
End Sub

' ============================================================
' LISTVIEW
' ============================================================
Private Sub BuildListView(f As Object)
    ' Csak pozíció és méret — oszlopok Initialize-ban!
    f.Controls.Add "MSComctlLib.ListViewCtrl.2", "lstEllenorzes"
    With f.Controls("lstEllenorzes")
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
    ' Footer háttérsáv
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

