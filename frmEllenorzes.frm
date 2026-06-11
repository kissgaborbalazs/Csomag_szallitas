VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEllenorzes 
   Caption         =   "Ellenõrzés – szállítónak átadásra váró küldemények"
   ClientHeight    =   9930.001
   ClientLeft      =   30
   ClientTop       =   -105
   ClientWidth     =   28005
   OleObjectBlob   =   "frmEllenorzes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEllenorzes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmEllenorzes — form kód
'
' SubItems térkép (0-bazisú SubItems index, Item.Text = col 0):
'   Item.Text   = "" (checkbox szimul., col 0)
'   SubItems(1) = kuldemeny_azonosito       (col 1)
'   SubItems(2) = megrendelo_vevokod        (col 2)
'   SubItems(3) = kuldemeny_aruertek        (col 3)
'   SubItems(4) = felrako_nev               (col 4)
'   SubItems(5) = felrako_cim              (col 5, összefuzve)
'   SubItems(6) = felrako_megjegyzes        (col 6)
'   SubItems(7) = felrako_kontakt_szemely   (col 7)
'   SubItems(8) = felrako_telefonszam       (col 8)
'   SubItems(9) = felrako_email             (col 9)
'   SubItems(10)= felrako_idokapu           (col 10)
'   SubItems(11)= cimzett_nev              (col 11)
'   SubItems(12)= cimzett_cim             (col 12, összefuzve)
'   SubItems(13)= cimzett_megjegyzes        (col 13)
'   SubItems(14)= cimzett_kontakt_szemely   (col 14)
'   SubItems(15)= cimzett_telefonszam       (col 15)
'   SubItems(16)= cimzett_email             (col 16)
'   SubItems(17)= lerakas_idokapu           (col 17)
'   SubItems(18)= raklap H×Sz×M            (col 18, összefuzve)
'   SubItems(19)= kuldemeny_brutto_tomege_kg(col 19)
'   SubItems(20)= utanvet_osszege_ft        (col 20)
'   SubItems(21)= ekaer                     (col 21)
'   SubItems(22)= adr_jelolo                (col 22)
'   SubItems(23)= adr_pont                  (col 23)
'   SubItems(24)= szolgaltatas              (col 24)
'   SubItems(25)= proc_data_id              (rejtett — ütközésdetektálás)
' ============================================================

Private Sub lstEllenorzes_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    UpdateKijeloltLabel
End Sub

Private Sub UserForm_Initialize()
    SetupListColumns
    LoadEllenorzes
End Sub

Private Sub SetupListColumns()
    With lstEllenorzes
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .CheckBoxes = True
        .ColumnHeaders.Add , , "OK", 20
        .ColumnHeaders.Add , , "Küldemény azonosító", 130
        .ColumnHeaders.Add , , "Vevokód", 70
        .ColumnHeaders.Add , , "Áruérték", 70
        .ColumnHeaders.Add , , "Felrakó neve", 120
        .ColumnHeaders.Add , , "Felrakó cím", 180
        .ColumnHeaders.Add , , "Felrakó megj.", 100
        .ColumnHeaders.Add , , "Felrakó kontakt", 100
        .ColumnHeaders.Add , , "Felrakó tel.", 90
        .ColumnHeaders.Add , , "Felrakó email", 130
        .ColumnHeaders.Add , , "Felrakó idokapu", 110
        .ColumnHeaders.Add , , "Címzett neve", 120
        .ColumnHeaders.Add , , "Címzett cím", 180
        .ColumnHeaders.Add , , "Címzett megj.", 100
        .ColumnHeaders.Add , , "Címzett kontakt", 100
        .ColumnHeaders.Add , , "Címzett tel.", 90
        .ColumnHeaders.Add , , "Címzett email", 130
        .ColumnHeaders.Add , , "Lerakás idokapu", 110
        .ColumnHeaders.Add , , "Raklap H×Sz×M (cm)", 110
        .ColumnHeaders.Add , , "Bruttó tömeg (kg)", 80
        .ColumnHeaders.Add , , "Utánvét (Ft)", 80
        .ColumnHeaders.Add , , "EKAER", 80
        .ColumnHeaders.Add , , "ADR jelölo", 70
        .ColumnHeaders.Add , , "ADR pont", 60
        .ColumnHeaders.Add , , "Szolgáltatás", 90
        ' 26. oszlop: rejtett proc_data_id (szélesség = 0)
        .ColumnHeaders.Add , , "proc_data_id", 0
    End With
End Sub

' ============================================================
' Betölto
' ============================================================

Private Sub LoadEllenorzes()
    On Error GoTo ErrHnd
    Dim params(2) As Variant
    params(0) = Array("@vevokod_filter", Trim(txtVevokod.Text))
    params(1) = Array("@kuldemeny_filter", Trim(txtAzon.Text))
    params(2) = Array("@date_from", modUtils.SafeDate(txtDatum.Text))

    Dim rs As Object
    Set rs = modConnection.ExecuteSP("usp_GetEllenorzesreVaroList", params)

    lstEllenorzes.ListItems.Clear

    Dim itm As MSComctlLib.ListItem
    Do While Not rs.EOF
        Set itm = lstEllenorzes.ListItems.Add
        itm.Text = ""
        itm.SubItems(1) = NvStr(rs("kuldemeny_azonosito"))
        itm.SubItems(2) = NvStr(rs("megrendelo_vevokod"))
        itm.SubItems(3) = NvStr(rs("kuldemeny_aruertek"))
        itm.SubItems(4) = NvStr(rs("felrako_nev"))
        itm.SubItems(5) = JoinCim( _
            rs("felrako_irsz"), rs("felrako_varos"), _
            rs("felrako_kozterulet_elnevezese"), _
            rs("felrako_kozterulet_tipus"), rs("felrako_hazszam"))
        itm.SubItems(6) = NvStr(rs("felrako_megjegyzes"))
        itm.SubItems(7) = NvStr(rs("felrako_kontakt_szemely"))
        itm.SubItems(8) = NvStr(rs("felrako_telefonszam"))
        itm.SubItems(9) = NvStr(rs("felrako_email"))
        itm.SubItems(10) = NvStr(rs("felrako_idokapu"))
        itm.SubItems(11) = NvStr(rs("cimzett_nev"))
        itm.SubItems(12) = JoinCim( _
            rs("cimzett_irsz"), rs("cimzett_varos"), _
            rs("cimzett_kozterulet_nev"), _
            rs("cimzett_kozterulet_jelleg"), rs("cimzett_hazszam"))
        itm.SubItems(13) = NvStr(rs("cimzett_megjegyzes"))
        itm.SubItems(14) = NvStr(rs("cimzett_kontakt_szemely"))
        itm.SubItems(15) = NvStr(rs("cimzett_telefonszam"))
        itm.SubItems(16) = NvStr(rs("cimzett_email"))
        itm.SubItems(17) = NvStr(rs("lerakas_idokapu"))
        itm.SubItems(18) = JoinRaklap( _
            rs("raklap_hossza_cm"), _
            rs("raklap_szelessege_cm"), _
            rs("raklap_magassaga_cm"))
        itm.SubItems(19) = NvStr(rs("kuldemeny_brutto_tomege_kg"))
        itm.SubItems(20) = NvStr(rs("utanvet_osszege_ft"))
        itm.SubItems(21) = NvStr(rs("ekaer"))
        itm.SubItems(22) = NvStr(rs("adr_jelolo"))
        itm.SubItems(23) = NvStr(rs("adr_pont"))
        itm.SubItems(24) = NvStr(rs("szolgaltatas"))
        ' Rejtett: proc_data_id — ütközésdetektáláshoz
        itm.SubItems(25) = NvStr(rs("proc_data_id"))
        rs.MoveNext
    Loop

    UpdateKijeloltLabel
    Exit Sub
ErrHnd:
    modUtils.ShowError "Ellenorzés lista betöltés", Err.Description
End Sub

' ============================================================
' Kijelölés kezelok
' ============================================================

Private Sub lstEllenorzes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    UpdateKijeloltLabel
End Sub

Private Sub UpdateKijeloltLabel()
    Dim n As Integer
    Dim itm As MSComctlLib.ListItem
    For Each itm In lstEllenorzes.ListItems
        If itm.Checked Then n = n + 1
    Next itm
    lblKijeloltDb.Caption = "Kijelölve: " & n & " db"
    btnAtadhato.Enabled = (n > 0)
    btnStornoSelected.Enabled = (n > 0)
End Sub

Private Sub btnSelectAll_Click()
    Dim itm As MSComctlLib.ListItem
    For Each itm In lstEllenorzes.ListItems
        itm.Checked = True
    Next itm
    UpdateKijeloltLabel
End Sub

Private Sub btnSelectNone_Click()
    Dim itm As MSComctlLib.ListItem
    For Each itm In lstEllenorzes.ListItems
        itm.Checked = False
    Next itm
    UpdateKijeloltLabel
End Sub

' ============================================================
' Átadhatónak jelöl — szallitonak_atadhato = 1 (tömeges)
' ============================================================

Private Sub btnAtadhato_Click()
    ' Kijelölt sorok összegyujtése
    Dim n As Integer
    Dim itm As MSComctlLib.ListItem
    For Each itm In lstEllenorzes.ListItems
        If itm.Checked Then n = n + 1
    Next itm
    If n = 0 Then Exit Sub

    If MsgBox(n & " küldemény átadhatónak jelölése?" & vbCrLf & _
              "A szallitonak_atadhato flag 1-re kerül.", _
              vbYesNo + vbQuestion, "Átadhatónak jelöl") <> vbYes Then Exit Sub

    On Error GoTo ErrHnd

    Dim usr As String
    usr = modConfig.GetCurrentUser()

    ' JSON: csak a szallitonak_atadhato mezo változik
    Dim procDict As Object
    Set procDict = CreateObject("Scripting.Dictionary")
    procDict("szallitonak_atadhato") = "1"
    
    
    'Dim json As String
    'json = "{""szallitonak_atadhato"":""1""}"

    Dim konflikt As Integer   ' ütközések száma — tájékoztató célból
    Dim k As Integer

    For Each itm In lstEllenorzes.ListItems
        If itm.Checked Then
            Dim azon13 As String
            azon13 = Left(itm.SubItems(1), 13)

            ' proc_data_id: "" = nincs még sor (NULL-t adunk át)
            Dim activeId As Long
            If itm.SubItems(25) = "" Then
                activeId = 0 'Null
            Else
                activeId = CLng(itm.SubItems(25))
            End If

            Dim conflict As Boolean
            modConnection.ExecuteSPUpsert "usp_UpsertProcData", _
                azon13, activeId, modUtils.DictToJson(procDict), conflict

            If conflict Then
                konflikt = konflikt + 1
            End If
        End If
    Next itm

    If konflikt > 0 Then
        MsgBox konflikt & " tétel ütközést jelzett (közben módosult) — lista frissítve.", _
               vbExclamation, "Ütközés"
    End If

    LoadEllenorzes
    Exit Sub
ErrHnd:
    modUtils.ShowError "Átadhatónak jelölés", Err.Description
End Sub

' ============================================================
' Stornó — CANCELLED státusz tömeges rögzítés
' ============================================================

Private Sub btnStornoSelected_Click()
    Dim n As Integer
    Dim itm As MSComctlLib.ListItem
    For Each itm In lstEllenorzes.ListItems
        If itm.Checked Then n = n + 1
    Next itm
    If n = 0 Then Exit Sub

    If MsgBox(n & " küldemény stornózása?" & vbCrLf & _
              "CANCELLED státusz kerül rögzítésre. Ez nem vonható vissza!", _
              vbYesNo + vbExclamation, "Stornó megerõsítés") <> vbYes Then Exit Sub

    On Error GoTo ErrHnd

    For Each itm In lstEllenorzes.ListItems
    If itm.Checked Then
        modConnection.ExecuteSPSetStatus _
            Left(itm.SubItems(1), 13), _
            "CANCELLED", _
            modConfig.GetCurrentUser(), _
            ""
    End If
    Next itm

    LoadEllenorzes
    Exit Sub
ErrHnd:
    modUtils.ShowError "Stornó", Err.Description
End Sub

' ============================================================
' Szuro eseménykezelok
' ============================================================

Private Sub txtVevokod_Change():
    'LoadEllenorzes
End Sub
Private Sub txtAzon_Change():
    'LoadEllenorzes
End Sub
Private Sub txtDatum_Change():
    'LoadEllenorzes
End Sub
Private Sub btnSzures_Click():   LoadEllenorzes: End Sub

Private Sub btnTorol_Click()
    txtVevokod.Text = ""
    txtAzon.Text = ""
    txtDatum.Text = ""
    LoadEllenorzes
End Sub

Private Sub lstEllenorzes_DblClick()
    If lstEllenorzes.SelectedItem Is Nothing Then Exit Sub
    Dim azon As String
    azon = lstEllenorzes.SelectedItem.SubItems(1)
    If azon = "" Then Exit Sub
    frmEllenorzesModal.LoadOrder azon
    frmEllenorzesModal.Show vbModal
    LoadEllenorzes
End Sub

Private Sub btnBezaras_Click()
    Unload Me
End Sub

' ============================================================
' Segédfüggvények
' ============================================================

Private Function JoinCim(irsz As Variant, varos As Variant, _
                         kozterulet As Variant, tipus As Variant, _
                         hazszam As Variant) As String
    Dim s As String
    s = NvStr(irsz) & " " & NvStr(varos)
    Dim utca As String
    utca = Trim(NvStr(kozterulet) & " " & NvStr(tipus) & " " & NvStr(hazszam))
    If utca <> "" Then s = s & ", " & utca
    JoinCim = Trim(s)
End Function

Private Function JoinRaklap(h As Variant, sz As Variant, m As Variant) As String
    If IsNull(h) And IsNull(sz) And IsNull(m) Then
        JoinRaklap = ""
        Exit Function
    End If
    JoinRaklap = NvStr(h) & "×" & NvStr(sz) & "×" & NvStr(m)
End Function

Private Function NvStr(v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        NvStr = ""
    Else
        NvStr = CStr(v)
    End If
End Function


