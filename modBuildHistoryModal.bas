Attribute VB_Name = "modBuildHistoryModal"
Option Explicit

' ============================================================
' modBuildHistoryModal
' Feltételezi hogy frmHistory már létezik.
' Kontrollokat törli és hozza létre újra.
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

' --- Layout konstansok ---
Const FORM_W            As Integer = 1000   ' 1000 × 600 px — vbc.Width = FORM_W + 12, Height = KY_FOOTER + 38
Const KX_LEFT           As Integer = 10
Const KX_RIGHT          As Integer = 506    ' bal panel ~490px
Const KY_HEADER         As Integer = 8
Const KY_SEP_TOP        As Integer = 28
Const KY_BODY_HDR       As Integer = 36     ' szekció fejléc sora
Const KY_BODY           As Integer = 56     ' lista teteje
Const KY_SEP_BOT        As Integer = 552    ' footer elott: 600 - 48
Const KY_FOOTER         As Integer = 562
Const FONT_SIZE_DEFAULT As Integer = 9

Private Sub BuildHistoryModal_Run()
    BuildControls
    MsgBox "frmHistory kontrollok elkészültek.", vbInformation, "Build"
End Sub

Private Sub BuildControls()
    Dim f As Object
    Set f = ThisWorkbook.VBProject.VBComponents("frmHistory").Designer

    ' Form méret és háttérszín
    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents("frmHistory")
    vbc.Properties("Width") = FORM_W + 12
    vbc.Properties("Height") = KY_FOOTER + 58
    vbc.Properties("BackColor") = CLR_FORM_BG

    ' -- Meglévo kontrollok törlése ---------------------------
    Dim i As Integer
    For i = f.Controls.count - 1 To 0 Step -1
        f.Controls.Remove f.Controls(i).name
    Next i

    ' =========================================================
    ' FEJLÉC
    ' =========================================================

    ' Header háttérsáv
    With f.Controls.Add("Forms.Label.1", "lblHeaderBg")
        .Left = 0
        .Top = 0
        .Width = FORM_W
        .Height = KY_SEP_TOP
        .BackColor = CLR_G050
        .Caption = ""
    End With

    With f.Controls.Add("Forms.Label.1", "lblAzonFix")
        .Left = KX_LEFT
        .Top = KY_HEADER
        .Width = 65
        .Height = 16
        .Caption = "Küldemény:"
        .Font.Bold = True
        .Font.size = FONT_SIZE_DEFAULT
        .ForeColor = CLR_G700
        .BackColor = CLR_G050
        .BackStyle = 1
    End With

    With f.Controls.Add("Forms.Label.1", "lblAzonVal")
        .Left = 78
        .Top = KY_HEADER
        .Width = 300
        .Height = 16
        .Caption = ""
        .Font.Bold = True
        .Font.size = FONT_SIZE_DEFAULT
        .ForeColor = CLR_G700
        .BackColor = CLR_G050
        .BackStyle = 1
    End With

    ' Felso elválasztó
    With f.Controls.Add("Forms.Label.1", "lblSepTop")
        .Left = 0
        .Top = KY_SEP_TOP
        .Width = FORM_W
        .Height = 2
        .BackColor = CLR_G700
        .Caption = ""
    End With

    ' =========================================================
    ' BAL PANEL — Státuszok
    ' =========================================================

    With f.Controls.Add("Forms.Label.1", "lblStatusHdr")
        .Left = KX_LEFT
        .Top = KY_BODY_HDR
        .Width = 220
        .Height = 16
        .Caption = "STÁTUSZOK"
        .Font.Bold = True
        .Font.size = FONT_SIZE_DEFAULT
        .ForeColor = CLR_G700
        .BackColor = CLR_FORM_BG
        .BackStyle = 1
    End With

    ' lstStatus: MSComctlLib ListView
    ' Oszlopok (Initialize-ban állítandók be):
    '   Col 0 — Dátum/ido   (120 px)
    '   Col 1 — Státusz     (160 px)
    '   Col 2 — Ki          (90 px)
    '   Col 3 — Megjegyzés  (fill)
    f.Controls.Add "MSComctlLib.ListViewCtrl.2", "lstStatus"
    With f.Controls("lstStatus")
        .Left = KX_LEFT
        .Top = KY_BODY
        .Width = 478
        .Height = 488
    End With

    ' Függoleges elválasztó
    With f.Controls.Add("Forms.Label.1", "lblVDiv")
        .Left = 496
        .Top = KY_BODY_HDR
        .Width = 2
        .Height = KY_SEP_BOT - KY_BODY_HDR
        .BackColor = CLR_LINE
        .Caption = ""
    End With

    ' =========================================================
    ' JOBB PANEL — Mezováltozások
    ' =========================================================

    With f.Controls.Add("Forms.Label.1", "lblFieldsHdr")
        .Left = KX_RIGHT
        .Top = KY_BODY_HDR
        .Width = 240
        .Height = 16
        .Caption = "MEZOVÁLTOZÁSOK"
        .Font.Bold = True
        .Font.size = FONT_SIZE_DEFAULT
        .ForeColor = CLR_G700
        .BackColor = CLR_FORM_BG
        .BackStyle = 1
    End With

    ' lstFields: MSComctlLib ListView
    ' Oszlopok (Initialize-ban állítandók be):
    '   Col 0 — Dátum/ido  (120 px)
    '   Col 1 — Ki         (90 px)
    '   Col 2 — Mezo       (150 px)
    '   Col 3 — Elotte     (120 px)
    '   Col 4 — Utána      (fill)
    f.Controls.Add "MSComctlLib.ListViewCtrl.2", "lstFields"
    With f.Controls("lstFields")
        .Left = KX_RIGHT
        .Top = KY_BODY
        .Width = 482
        .Height = 488
    End With

    ' =========================================================
    ' FOOTER
    ' =========================================================

    ' Alsó elválasztó
    With f.Controls.Add("Forms.Label.1", "lblSepBot")
        .Left = 0
        .Top = KY_SEP_BOT
        .Width = FORM_W
        .Height = 2
        .BackColor = CLR_G700
        .Caption = ""
    End With

    ' Footer háttérsáv
    With f.Controls.Add("Forms.Label.1", "lblFooterBg")
        .Left = 0
        .Top = KY_SEP_BOT + 2
        .Width = FORM_W
        .Height = 40
        .BackColor = CLR_G050
        .Caption = ""
    End With

    With f.Controls.Add("Forms.CommandButton.1", "btnClose")
        .Left = 440
        .Top = KY_FOOTER
        .Width = 100
        .Height = 24
        .Caption = "Bezárás"
        .Font.size = FONT_SIZE_DEFAULT
        .BackColor = CLR_G700
        .ForeColor = CLR_SURF
    End With

End Sub


' ============================================================
' SetupListColumns — frmHistory UserForm_Initialize-ból hívni
' ============================================================
'
' Private Sub SetupListColumns()
'
'     With lstStatus
'         .View = lvwReport
'         .FullRowSelect = True
'         .GridLines = True
'         .ColumnHeaders.Add , , "Dátum/ido",   120
'         .ColumnHeaders.Add , , "Státusz",      160
'         .ColumnHeaders.Add , , "Ki",            90
'         .ColumnHeaders.Add , , "Megjegyzés",    -1   ' fill
'     End With
'
'     With lstFields
'         .View = lvwReport
'         .FullRowSelect = True
'         .GridLines = True
'         .ColumnHeaders.Add , , "Dátum/ido",  120
'         .ColumnHeaders.Add , , "Ki",           90
'         .ColumnHeaders.Add , , "Mezo",        150
'         .ColumnHeaders.Add , , "Elotte",      120
'         .ColumnHeaders.Add , , "Utána",        -1   ' fill
'     End With
'
' End Sub


' ============================================================
' LoadHistory — frmHistory betöltési logika (vázlat)
' ============================================================
'
' Public Sub LoadOrder(azon13 As String)
'
'     lblAzonVal.Caption = azon13
'
'     Dim params(0) As Variant
'     params(0) = Array("@kuldemeny_azon_13", azon13)
'     Dim rs As ADODB.Recordset
'     Set rs = modConnection.ExecuteSP("usp_GetPackageHistory", params)
'
'     lstStatus.ListItems.Clear
'     lstFields.ListItems.Clear
'
'     Dim statusCount As Integer: statusCount = 0
'     Dim fieldCount  As Integer: fieldCount  = 0
'
'     Do While Not rs.EOF
'
'         Dim evtTime  As String
'         Dim evtActor As String
'         evtTime  = Format(rs("event_time").Value, "yyyy-mm-dd hh:nn")
'         evtActor = Nz(rs("actor").Value, "")
'
'         Select Case rs("event_type").Value
'
'             Case "STATUS_CHANGE"
'                 Dim sItem As ListItem
'                 Set sItem = lstStatus.ListItems.Add
'                 sItem.Text        = evtTime
'                 sItem.SubItems(1) = Nz(rs("new_value").Value, "")
'                 sItem.SubItems(2) = evtActor
'                 sItem.SubItems(3) = Nz(rs("note").Value, "")
'                 statusCount = statusCount + 1
'
'             Case "FIELD_CHANGE"
'                 Dim parts() As String
'                 parts = Split(Nz(rs("event_label").Value, "/"), " / ")
'                 Dim fieldName As String
'                 fieldName = parts(UBound(parts))   ' csak a mezo neve, táblanév nélkül
'
'                 Dim fItem As ListItem
'                 Set fItem = lstFields.ListItems.Add
'                 fItem.Text        = evtTime
'                 fItem.SubItems(1) = evtActor
'                 fItem.SubItems(2) = fieldName
'                 fItem.SubItems(3) = Nz(rs("old_value").Value, Chr(8709))  ' Ø ha üres
'                 fItem.SubItems(4) = Nz(rs("new_value").Value, Chr(8709))
'                 fieldCount = fieldCount + 1
'
'         End Select
'
'         rs.MoveNext
'     Loop
'
'     lblStatusHdr.Caption = "STÁTUSZOK (" & statusCount & ")"
'     lblFieldsHdr.Caption = "MEZOVÁLTOZÁSOK (" & fieldCount & ")"
'
' End Sub

