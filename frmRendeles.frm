VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRendeles 
   Caption         =   "Rendelések"
   ClientHeight    =   5985
   ClientLeft      =   -390
   ClientTop       =   -2610
   ClientWidth     =   2550
   OleObjectBlob   =   "frmRendeles.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmRendeles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmRendeles — Rendelések áttekintõ
' ListView: MSComctlLib.ListViewCtrl (Microsoft ListView Control 6.0)
' Reference szükséges: Microsoft Windows Common Controls 6.0
'
' Design-time kontrollok:
'   fraStats
'   lblNyitottFix, lblNyitottVal
'   lblHianyzoFix, lblHianyzoEFJVal
'   lblEllenorzesreFix, lblEllenorzesreVal
'   lblVevokodFix, txtVevokod
'   lblAzonFix, txtAzon
'   lblDatumFix, txtDatum
'   lblStatusFix, cboStatus
'   btnSzures, btnTorol
'   lvRendeles  (ListView kontroll)
'
' Szinezesi logika:
'   Narancs  — EFJ meg nem erkezett (raw_efj_id IS NULL)
'   Kek      — EFJ megvan, szallitonak_atadhato = 0
'   Zold     — szallitonak_atadhato = 1
'   Piros    — CANCELLED
'   Szurke   — lezart (status_is_locked = 1)
' ============================================================

' SP SELECT index map:
'   0=raw_order_id, 1=kuldemeny_azonosito, 2=kuldemeny_azonosito_13
'   3=megrendelo_vevokod, 4=megrendelo_cegnev, 5=megbizas_ideje (yyyy-mm-dd hh:mm)
'   6=felrako_kontakt_szemely, 7=felrako_telefonszam, 8=felrako_email
'   9=current_status_code, 10=current_status_label, 11=status_is_locked
'   12=operative_data_active_id, 13=proc_data_active_id
'   14=raw_efj_id, 15=szallitonak_atadhato

' ListView ForeColor szinezés (BackColor nem tamogatott minden verzioban)
Private Const CLR_NARANCS  As Long = &H66CC     ' narancs FG (EFJ hianyzik)
Private Const CLR_KEK      As Long = &H996600   ' kek FG (ellenorzesre var)
Private Const CLR_ZOLD     As Long = &H6633     ' zold FG (atadhato)
Private Const CLR_PIROS    As Long = &HAA       ' piros FG (cancelled)
Private Const CLR_SZURKE   As Long = &H888888   ' szurke FG (lezart)
Private Const CLR_NORMAL   As Long = &H0        ' fekete FG (normal)
Private Const LIST_FONT_SIZE As Long = 10

Private Sub cboStatus_Change()
    LoadStats
    LoadRendeles
End Sub

Private Sub lvRendeles_DblClick()
    If lvRendeles.SelectedItem Is Nothing Then Exit Sub
    Dim azon As String
    azon = lvRendeles.SelectedItem.Text
    frmEllenorzesModal.LoadOrder (Left(azon, 13))
    frmEllenorzesModal.Show vbModal
    'frmHistory.LoadOrder(Left(lvRendeles.SelectedItem.Text, 13))
    'frmHistory.Show vbModal
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Rendelések"
    Me.Width = 1000
    Me.Height = 600
    Me.BackColor = RGB(245, 245, 245)
    Me.StartUpPosition = 1

    ' -- Statisztika frame ------------------------------------
    With fraStats
        .Left = 8: .Top = 6: .Width = Me.Width - 28: .Height = 36
        .Caption = ""
        .BackColor = RGB(235, 235, 235)
    End With

    With lblNyitottFix
        .Left = 16: .Top = 10: .Width = 55: .Height = 14
        .Caption = "Nyitott:": .Font.Bold = True: .BackColor = RGB(235, 235, 235)
    End With
    With lblNyitottVal
        .Left = 74: .Top = 10: .Width = 35: .Height = 14
        .Caption = "-": .Font.Bold = True: .ForeColor = RGB(0, 100, 180)
        .BackColor = RGB(235, 235, 235)
    End With

    With lblHianyzoFix
        .Left = 120: .Top = 10: .Width = 85: .Height = 14
        .Caption = "Hiányzó EFJ:": .Font.Bold = True
        .BackColor = RGB(235, 235, 235)
    End With
    With lblHianyzoEFJVal
        .Left = 208: .Top = 10: .Width = 35: .Height = 14
        .Caption = "-": .Font.Bold = True: .ForeColor = RGB(180, 100, 0)
    End With

    With lblEllenorzesreFix
        .Left = 255: .Top = 10: .Width = 110: .Height = 14
        .Caption = "Ellenõrzésre vár:": .Font.Bold = True
        .BackColor = RGB(235, 235, 235)
    End With
    With lblEllenorzesreVal
        .Left = 368: .Top = 10: .Width = 35: .Height = 14
        .Caption = "-": .Font.Bold = True: .ForeColor = RGB(0, 130, 60)
        .BackColor = RGB(235, 235, 235)
    End With

    ' -- Szûrõ sor --------------------------------------------
    With lblVevokodFix
        .Left = 8: .Top = 50: .Width = 55: .Height = 14
        .Caption = "Vevõkód:"
    End With
    With txtVevokod
        .Left = 8: .Top = 64: .Width = 90: .Height = 18
    End With

    With lblAzonFix
        .Left = 106: .Top = 50: .Width = 60: .Height = 14
        .Caption = "Azonosító:"
    End With
    With txtAzon
        .Left = 106: .Top = 64: .Width = 140: .Height = 18
    End With

    With lblDatumFix
        .Left = 254: .Top = 50: .Width = 50: .Height = 14
        .Caption = "Dátum:"
    End With
    With txtDatum
        .Left = 254: .Top = 64: .Width = 80: .Height = 18
    End With

    With lblStatusFix
        .Left = 342: .Top = 50: .Width = 50: .Height = 14
        .Caption = "Státusz:"
    End With
    With cboStatus
        .Left = 342: .Top = 64: .Width = 150: .Height = 18
        .Style = 2
    End With

    With btnSzures
        .Left = 500: .Top = 61: .Width = 58: .Height = 22
        .Caption = "Szûrés"
        .BackColor = RGB(0, 112, 192): .ForeColor = RGB(255, 255, 255)
    End With
    With btnTorol
        .Left = 564: .Top = 61: .Width = 58: .Height = 22
        .Caption = "Törlés"
    End With

    ' -- ListView ---------------------------------------------
    With lvRendeles
        .Left = 8: .Top = 90: .Width = Me.Width - 28: .Height = Me.Height - 130
        .View = 3               ' lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = 1          ' lvwManual — nem szerkesztheto
        .HideSelection = False

        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Azonosító", 155
        .ColumnHeaders.Add , , "Vevõkód", 65
        .ColumnHeaders.Add , , "Vevõ neve", 150
        .ColumnHeaders.Add , , "Beérkezés", 90
        .ColumnHeaders.Add , , "Kapcsolattartó", 150
        .ColumnHeaders.Add , , "Telefon", 75
        .ColumnHeaders.Add , , "E-mail", 150
        .ColumnHeaders.Add , , "Státusz", 120
    End With

    ' -- Adatok betöltése ------------------------------------
    LoadStatusCombo
    LoadStats
    LoadRendeles
End Sub

' ============================================================
' Betöltõk
' ============================================================

Private Sub LoadStatusCombo()
    cboStatus.Clear
    cboStatus.AddItem "(összes)"
    On Error GoTo ErrHnd
    Dim rs As Object
    Set rs = modConnection.ExecuteSP("usp_GetStatusList")
    Do While Not rs.EOF
        cboStatus.AddItem CStr(rs.Fields("status_label_hu").Value) & _
                          " (" & CStr(rs.Fields("status_code").Value) & ")"
        rs.MoveNext
    Loop
    cboStatus.ListIndex = 0
    Exit Sub
ErrHnd:
    modUtils.ShowError "Státuszlista", Err.Description
End Sub

Private Sub LoadStats()
    On Error GoTo ErrHnd
    Dim rs As Object
    Set rs = modConnection.ExecuteSP("usp_GetDashboardStats")
    If Not (rs.BOF And rs.EOF) Then
        lblNyitottVal.Caption = CStr(rs.Fields("nyitott_count").Value)
        lblHianyzoEFJVal.Caption = CStr(rs.Fields("hianyzo_efj_count").Value)
        lblEllenorzesreVal.Caption = CStr(rs.Fields("ellenorzesre_count").Value)
    End If
    Exit Sub
ErrHnd:
    lblNyitottVal.Caption = "-"
    lblHianyzoEFJVal.Caption = "-"
    lblEllenorzesreVal.Caption = "-"
End Sub

Private Sub LoadRendeles()
    On Error GoTo ErrHnd
    Dim params(3) As Variant
    params(0) = Array("@vevokod_filter", Trim(txtVevokod.Text))
    params(1) = Array("@kuldemeny_filter", Trim(txtAzon.Text))
    params(2) = Array("@date_from", modUtils.SafeDate(txtDatum.Text))
    params(3) = Array("@status_code", GetSelectedStatusCode())
    Dim rs As Object
    Set rs = modConnection.ExecuteSP("usp_GetPackageList", params)
    FillListView rs
    Exit Sub
ErrHnd:
    modUtils.ShowError "Rendelések betöltés", Err.Description
End Sub

Private Sub FillListView(rs As Object)
    lvRendeles.ListItems.Clear
    If rs Is Nothing Then Exit Sub
    If rs.BOF And rs.EOF Then Exit Sub

    Do While Not rs.EOF
        ' Adatok kiolvasása
        Dim azon        As String: azon = NS(rs.Fields("kuldemeny_azonosito").Value)
        Dim vevokod     As String: vevokod = NS(rs.Fields("megrendelo_vevokod").Value)
        Dim cegnev      As String: cegnev = NS(rs.Fields("megrendelo_cegnev").Value)
        Dim beerkezett  As String: beerkezett = NS(rs.Fields("megbizas_datetime").Value)
        Dim kontakt     As String: kontakt = NS(rs.Fields("felrako_kontakt_szemely").Value)
        Dim telefon     As String: telefon = NS(rs.Fields("felrako_telefonszam").Value)
        Dim email       As String: email = NS(rs.Fields("felrako_email").Value)
        Dim statusLabel As String: statusLabel = NS(rs.Fields("current_status_label").Value)
        Dim statusCode  As String: statusCode = NS(rs.Fields("current_status_code").Value)
        Dim isLocked    As Boolean
        Dim hasEfj      As Boolean
        Dim atadhato    As Boolean

        If IsNull(rs.Fields("status_is_locked").Value) Then isLocked = False Else isLocked = CBool(rs.Fields("status_is_locked").Value)
        hasEfj = Not IsNull(rs.Fields("raw_efj_id").Value)
        If IsNull(rs.Fields("szallitonak_atadhato").Value) Then atadhato = False Else atadhato = CBool(rs.Fields("szallitonak_atadhato").Value)

        ' ListItem létrehozása
        Dim li As Object
        Set li = lvRendeles.ListItems.Add(, , azon)
        li.SubItems(1) = vevokod
        li.SubItems(2) = cegnev
        li.SubItems(3) = beerkezett
        li.SubItems(4) = kontakt
        li.SubItems(5) = telefon
        li.SubItems(6) = email
        li.SubItems(7) = statusLabel

        ' Szinezés ForeColor-ral (BackColor nem tamogatott)
        If statusCode = "CANCELLED" Then
            li.ForeColor = CLR_SZURKE
        ElseIf atadhato Then
            li.ForeColor = CLR_ZOLD
        ElseIf Not atadhato Then
            li.ForeColor = CLR_NARANCS
        End If

        rs.MoveNext
    Loop
End Sub

Private Function NS(val As Variant) As String
    If IsNull(val) Then NS = "" Else NS = CStr(val)
End Function

Private Function FormatRsDate(val As Variant) As String
    If IsNull(val) Then FormatRsDate = "": Exit Function
    On Error Resume Next
    FormatRsDate = Format(CDate(val), "yyyy-mm-dd")
End Function

Private Function GetSelectedStatusCode() As String
    If cboStatus.ListIndex <= 0 Then GetSelectedStatusCode = "": Exit Function
    Dim s As String: s = cboStatus.Text
    Dim p1 As Integer: p1 = InStrRev(s, "(")
    Dim p2 As Integer: p2 = InStrRev(s, ")")
    If p1 > 0 And p2 > p1 Then
        GetSelectedStatusCode = Mid(s, p1 + 1, p2 - p1 - 1)
    End If
End Function

' ============================================================
' Eseménykezelõk
' ============================================================

Private Sub btnSzures_Click()
    LoadStats
    LoadRendeles
End Sub

Private Sub btnTorol_Click()
    txtVevokod.Text = ""
    txtAzon.Text = ""
    txtDatum.Text = ""
    cboStatus.ListIndex = 0
    LoadStats
    LoadRendeles
End Sub


