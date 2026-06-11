VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "UserForm2"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ============================================================
' frmMainMenu
' Design-time kontrollok (designerben létrehozva):
'   lblTitle, lblVersion, lblSep1, lblSep2, lblConn
'   btnRendeles, btnEllenorzes, btnAdmin, btnExport
' ============================================================

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
Private Const CLR_DISABLED  As Long = &H8AA08A        ' --ink3  letiltott = ink3

Private Sub UserForm_Initialize()
    Me.Caption = "Csomagstátusz kezelõ"
    Me.Width = 300
    Me.Height = 340
    Me.BackColor = CLR_FORM_BG
    Me.BorderStyle = 0
    Me.StartUpPosition = 1

    With lblTitle
        .Left = 10:         .Top = 12
        .Width = 275:       .Height = 22
        .Caption = "Csomagstátusz kezelõ"
        .Font.size = 13
        .Font.Bold = True
        .ForeColor = CLR_G700
        .BackColor = CLR_FORM_BG
        .BackStyle = 1
        .TextAlign = 2
    End With

    With lblVersion
        .Left = 10:         .Top = 36
        .Width = 275:       .Height = 14
        .Caption = modConfig.APP_VERSION
        .Font.size = 8
        .ForeColor = CLR_INK3
        .BackColor = CLR_FORM_BG
        .BackStyle = 1
        .TextAlign = 2
    End With

    With lblSep1
        .Left = 15:         .Top = 56
        .Width = 265:       .Height = 2
        .Caption = ""
        .BackColor = CLR_LINE
        .BackStyle = 1
    End With

    With btnRendeles
        .Left = 25:         .Top = 68
        .Width = 245:       .Height = 34
        .Caption = "Rendelések"
        .Font.size = 11
        .BackColor = CLR_G700
        .ForeColor = CLR_SURF
    End With

    With btnEllenorzes
        .Left = 25:         .Top = 112
        .Width = 245:       .Height = 34
        .Caption = "Ellenõrzés"
        .Font.size = 11
        .BackColor = CLR_G600
        .ForeColor = CLR_SURF
    End With

    With btnAdmin
        .Left = 25:         .Top = 156
        .Width = 245:       .Height = 34
        .Caption = "Adminisztráció"
        .Font.size = 11
        .BackColor = CLR_DISABLED
        .ForeColor = CLR_SURF
        .Enabled = False
    End With

    With btnExport
        .Left = 25:         .Top = 200
        .Width = 245:       .Height = 34
        .Caption = "Export"
        .Font.size = 11
        .BackColor = CLR_DISABLED
        .ForeColor = CLR_SURF
        .Enabled = True
    End With

    With lblSep2
        .Left = 15:         .Top = 245
        .Width = 265:       .Height = 2
        .Caption = ""
        .BackColor = CLR_LINE
        .BackStyle = 1
    End With

    With lblConn
        .Left = 10:         .Top = 256
        .Width = 275:       .Height = 16
        .Font.size = 9
        .BackColor = CLR_FORM_BG
        .BackStyle = 1
        .TextAlign = 2
    End With

    ' -- Kapcsolat teszt --------------------------------------
    If modConnection.TestConnection() Then
        lblConn.Caption = "Kapcsolódva"
        lblConn.ForeColor = CLR_G700
    Else
        lblConn.Caption = "Kapcsolódási hiba"
        lblConn.ForeColor = CLR_RED
    End If
End Sub

Private Sub btnRendeles_Click()
    frmRendeles.Show
End Sub

Private Sub btnEllenorzes_Click()
    frmEllenorzes.Show
End Sub

Private Sub btnAdmin_Click()
End Sub

Private Sub btnExport_Click()
    frmRendelesExprt.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    modConnection.CloseConnection
End Sub

