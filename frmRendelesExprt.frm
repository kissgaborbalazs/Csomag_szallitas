VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRendelesExprt 
   Caption         =   "Rendelés export – átadásra váró küldemények"
   ClientHeight    =   9930.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   28005
   OleObjectBlob   =   "frmRendelesExprt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRendelesExprt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' frmRendelesExprt — Eseménykezelok és metódusok
'
' Ide másolandó kód a frmRendelesExprt UserForm kódmodulba.
'
' Függoségek:
'   modConnection.ExecuteSP(spName, params)   ? ADODB.Recordset
'   modConnection.ExecuteSPNonQuery(spName, params)
'   modConfig.GetCurrentUser()                ? String
'   modUtils.SafeDate(val)                    ? Variant
'   modUtils.NvStr(val)                       ? String
'   modExport.ExportToExcel(rs, path)
'   modExport.GetExportFilePath()             ? String
'
' SP-k:
'   usp_GetRendelesExportList — @vevokod_filter, @kuldemeny_filter, @date_from
'   usp_SetPackageStatus      — @kuldemeny_azon_13, @status_code, @changed_by
'
' Megjegyzés a CheckBox-ról:
'   Az lstExport ListView .Checkboxes = True property-t használ.
'   Nincs önálló CheckBox oszlop — a jelölonégyzet a sor bal oldalán jelenik meg
'   (.ListItem.Checked). Az oszlop 0 a Build kommentben erre utal.
' ============================================================
Option Explicit

' --- Privát állapot ---
Private m_totalLoaded As Long   ' betöltött sorok száma

' ============================================================
' INITIALIZE — ListView setup + alapértelmezett betöltés
' ============================================================
Private Sub UserForm_Initialize()
    SetupListColumns
    LoadList
End Sub

' ============================================================
' SetupListColumns
' Csak itt definiálj ListView oszlopokat — a build scriptben
' szándékosan nincs oszlop-definíció.
' ============================================================
Private Sub SetupListColumns()
    With lstExport
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .CheckBoxes = True      ' sor szintu jelölonégyzet (nem oszlop!)
        .LabelEdit = lvwManual

        ' Oszlopfejlécek — sorrendben a build script kommentje szerint
        ' (a Checkboxes=True maga adja a jelölonégyzetet, ez az 1. adat-oszlop)
        .ColumnHeaders.Add , , "Küldemény azonosító", 130
        .ColumnHeaders.Add , , "Vevõkód", 70
        .ColumnHeaders.Add , , "Áruérték", 70
        .ColumnHeaders.Add , , "Küldemény megj.", 100
        .ColumnHeaders.Add , , "Felrakó neve", 100
        .ColumnHeaders.Add , , "Felrakó irsz", 55
        .ColumnHeaders.Add , , "Felrakó város", 90
        .ColumnHeaders.Add , , "Felrakó közterület", 110
        .ColumnHeaders.Add , , "Felrakó típus", 80
        .ColumnHeaders.Add , , "Felrakó hsz", 55
        .ColumnHeaders.Add , , "Felrakó megj.", 100
        .ColumnHeaders.Add , , "Felrakó kontakt", 100
        .ColumnHeaders.Add , , "Felrakó tel.", 85
        .ColumnHeaders.Add , , "Felrakó email", 110
        .ColumnHeaders.Add , , "Felrakó idokapu", 75
        .ColumnHeaders.Add , , "Címzett neve", 100
        .ColumnHeaders.Add , , "Címzett irsz", 55
        .ColumnHeaders.Add , , "Címzett város", 90
        .ColumnHeaders.Add , , "Címzett közterület", 110
        .ColumnHeaders.Add , , "Címzett jelleg", 80
        .ColumnHeaders.Add , , "Címzett hsz", 55
        .ColumnHeaders.Add , , "Címzett megj.", 100
        .ColumnHeaders.Add , , "Címzett kontakt", 100
        .ColumnHeaders.Add , , "Címzett tel.", 85
        .ColumnHeaders.Add , , "Címzett email", 110
        .ColumnHeaders.Add , , "Lerakás idokapu", 75
        .ColumnHeaders.Add , , "Raklap hossza (cm)", 75
        .ColumnHeaders.Add , , "Raklap széless. (cm)", 80
        .ColumnHeaders.Add , , "Raklap magass. (cm)", 80
        .ColumnHeaders.Add , , "Bruttó tömeg (kg)", 80
        .ColumnHeaders.Add , , "Utánvét (Ft)", 75
        .ColumnHeaders.Add , , "EKAER", 80
        .ColumnHeaders.Add , , "ADR jelölo", 70
        .ColumnHeaders.Add , , "ADR pont", 60
        .ColumnHeaders.Add , , "Szolgáltatás", 90
        .ColumnHeaders.Add , , "Lerakó ref.", 90
        .ColumnHeaders.Add , , "Felvételi ref.", 90
    End With
End Sub

' ============================================================
' loadList — SP hívás + lista feltöltés
' ============================================================
Private Sub LoadList()
    On Error GoTo ErrHandler

    lstExport.ListItems.Clear
    m_totalLoaded = 0
    updateSelectionCounter

    Dim params(2) As Variant
    params(0) = Array("@vevokod_filter", NvStr(txtVevokod.Text))
    params(1) = Array("@kuldemeny_filter", NvStr(txtAzon.Text))
    params(2) = Array("@date_from", modUtils.SafeDate(NvStr(txtDatum.Text)))

    Dim rs As Object
    Set rs = modConnection.ExecuteSP("usp_GetRendelesExportList", params)

    If rs Is Nothing Then GoTo ErrHandler
    If rs.EOF Then
        Set rs = Nothing
        Exit Sub
    End If

    Dim li As Object
    Do While Not rs.EOF
        Set li = lstExport.ListItems.Add()
        li.Checked = False

        ' Tag = proc_data_id rejtett értékként
        ' MEGJEGYZÉS: ha az SP nem ad vissza proc_data_id-t, hagyhatod üresen
        li.tag = NvStr(rs.Fields("proc_data_id").Value)

        ' Szöveg (1. adat-oszlop = ColumnHeader 1)
        li.Text = NvStr(rs.Fields("kuldemeny_azonosito").Value)

        ' SubItems (2..n ? ColumnHeader 2..n)
        li.SubItems(1) = NvStr(rs.Fields("megrendelo_vevokod").Value)
        li.SubItems(2) = NvStr(rs.Fields("kuldemeny_aruertek").Value)
        li.SubItems(3) = NvStr(rs.Fields("kuldemeny_megjegyzes").Value)
        li.SubItems(4) = NvStr(rs.Fields("felrako_nev").Value)
        li.SubItems(5) = NvStr(rs.Fields("felrako_irsz").Value)
        li.SubItems(6) = NvStr(rs.Fields("felrako_varos").Value)
        li.SubItems(7) = NvStr(rs.Fields("felrako_kozterulet_elnevezese").Value)
        li.SubItems(8) = NvStr(rs.Fields("felrako_kozterulet_tipus").Value)
        li.SubItems(9) = NvStr(rs.Fields("felrako_hazszam").Value)
        li.SubItems(10) = NvStr(rs.Fields("felrako_megjegyzes").Value)
        li.SubItems(11) = NvStr(rs.Fields("felrako_kontakt_szemely").Value)
        li.SubItems(12) = NvStr(rs.Fields("felrako_telefonszam").Value)
        li.SubItems(13) = NvStr(rs.Fields("felrako_email").Value)
        li.SubItems(14) = NvStr(rs.Fields("felrako_idokapu").Value)
        li.SubItems(15) = NvStr(rs.Fields("cimzett_nev").Value)
        li.SubItems(16) = NvStr(rs.Fields("cimzett_irsz").Value)
        li.SubItems(17) = NvStr(rs.Fields("cimzett_varos").Value)
        li.SubItems(18) = NvStr(rs.Fields("cimzett_kozterulet_nev").Value)
        li.SubItems(19) = NvStr(rs.Fields("cimzett_kozterulet_jelleg").Value)
        li.SubItems(20) = NvStr(rs.Fields("cimzett_hazszam").Value)
        li.SubItems(21) = NvStr(rs.Fields("cimzett_megjegyzes").Value)
        li.SubItems(22) = NvStr(rs.Fields("cimzett_kontakt_szemely").Value)
        li.SubItems(23) = NvStr(rs.Fields("cimzett_telefonszam").Value)
        li.SubItems(24) = NvStr(rs.Fields("cimzett_email").Value)
        li.SubItems(25) = NvStr(rs.Fields("lerakas_idokapu").Value)
        li.SubItems(26) = NvStr(rs.Fields("raklap_hossza_cm").Value)
        li.SubItems(27) = NvStr(rs.Fields("raklap_szelessege_cm").Value)
        li.SubItems(28) = NvStr(rs.Fields("raklap_magassaga_cm").Value)
        li.SubItems(29) = NvStr(rs.Fields("kuldemeny_brutto_tomege_kg").Value)
        li.SubItems(30) = NvStr(rs.Fields("utanvet_osszege_ft").Value)
        li.SubItems(31) = NvStr(rs.Fields("ekaer").Value)
        li.SubItems(32) = NvStr(rs.Fields("adr_jelolo").Value)
        li.SubItems(33) = NvStr(rs.Fields("adr_pont").Value)
        li.SubItems(34) = NvStr(rs.Fields("szolgaltatas").Value)
        li.SubItems(35) = NvStr(rs.Fields("lerako_referencia").Value)
        li.SubItems(36) = NvStr(rs.Fields("felveteli_referencia").Value)

        m_totalLoaded = m_totalLoaded + 1
        rs.MoveNext
    Loop

    Set rs = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Hiba a lista betöltésekor: " & Err.Description, vbCritical, "Betöltési hiba"
End Sub

' ============================================================
' SZURÉS GOMBOK
' ============================================================
Private Sub btnSzures_Click()
    LoadList
End Sub

Private Sub btnTorol_Click()
    txtVevokod.Text = ""
    txtAzon.Text = ""
    txtDatum.Text = ""
    LoadList
End Sub

' ============================================================
' KIJELÖLÉS KEZELÉS
' ============================================================
Private Sub btnSelectAll_Click()
    Dim li As Object
    For Each li In lstExport.ListItems
        li.Checked = True
    Next li
    updateSelectionCounter
    updateActionButtons
End Sub

Private Sub btnSelectNone_Click()
    Dim li As Object
    For Each li In lstExport.ListItems
        li.Checked = False
    Next li
    updateSelectionCounter
    updateActionButtons
End Sub

' ListView checkbox változás esemény
' Megjegyzés: MSComCtl ListView nem tüzel automatikus eseményt Checked változásra —
' az ItemClick eseményt használjuk, amely a checkbox kattintásakor is lefut.
Private Sub lstExport_ItemClick(ByVal Item As MSComctlLib.ListItem)
    updateSelectionCounter
    updateActionButtons
End Sub

' ============================================================
' updateSelectionCounter — "Kijelölve: N db" frissítése
' ============================================================
Private Sub updateSelectionCounter()
    Dim count As Long
    Dim li As Object
    For Each li In lstExport.ListItems
        If li.Checked Then count = count + 1
    Next li
    lblKijeloltDb.Caption = "Kijelölve: " & count & " db"
End Sub

' ============================================================
' updateActionButtons — Export és Átadva gombok engedélyezése
' ============================================================
Private Sub updateActionButtons()
    Dim hasChecked As Boolean
    Dim li As Object
    For Each li In lstExport.ListItems
        If li.Checked Then
            hasChecked = True
            Exit For
        End If
    Next li
    btnExportExcel.Enabled = hasChecked
    btnAtadva.Enabled = hasChecked
End Sub

' ============================================================
' EXPORT EXCELBE
' ============================================================
Private Sub btnExportExcel_Click()
    Dim checkedCount As Long
    checkedCount = countChecked()
    If checkedCount = 0 Then
        MsgBox "Nincs kijelölt tétel.", vbInformation, "Export"
        Exit Sub
    End If

    If MsgBox(checkedCount & " kijelölt tétel kerül exportálásra Excelbe." & vbCrLf & _
              "Folytatja?", vbYesNo + vbQuestion, "Export megerosítése") <> vbYes Then
        Exit Sub
    End If

    Dim filePath As String
    filePath = modExport.GetExportFilePath()
    If filePath = "" Then Exit Sub

    ' Recordset összerakása a kijelölt sorokból
    ' MEGJEGYZÉS: itt egy in-memory RS-t építünk fel a listából,
    ' mert az eredeti RS már lezárult. Alternatíva: az SP-t újra hívni
    ' a kijelölt azonosítókkal.
    Dim rs As Object
    Set rs = buildCheckedRecordset()

    If rs Is Nothing Then
        MsgBox "Hiba az adatok összeállításakor.", vbCritical, "Export hiba"
        Exit Sub
    End If

    modExport.ExportToExcel rs, filePath
    MsgBox checkedCount & " tétel exportálva.", vbInformation, "Export kész"
End Sub

' ============================================================
' ÁTADVA (KISZÁLLÍTÓNAK)
' Minden kijelölt tételre HANDED_OVER státuszt szúr be.
' Elofeltétel SP oldalon: szallitonak_atadhato = 1 (D11)
' ============================================================
Private Sub btnAtadva_Click()
    Dim checkedCount As Long
    checkedCount = countChecked()
    If checkedCount = 0 Then
        MsgBox "Nincs kijelölt tétel.", vbInformation, "Átadás"
        Exit Sub
    End If

    If MsgBox(checkedCount & " tétel kerül HANDED_OVER státuszba." & vbCrLf & _
              "Ez a muvelet visszavonhatatlan." & vbCrLf & vbCrLf & _
              "Folytatja?", vbYesNo + vbExclamation, "Átadás megerosítése") <> vbYes Then
        Exit Sub
    End If

    Dim currentUser As String
    currentUser = modConfig.GetCurrentUser()

    Dim successCount As Long
    Dim errorCount As Long
    Dim li As Object

    For Each li In lstExport.ListItems
        If li.Checked Then
            Dim azon13 As String
            azon13 = li.Text   ' kuldemeny_azonosito (elso 13 karakter = kuldemeny_azonosito_13)
            If Len(azon13) > 13 Then azon13 = Left(azon13, 13)

            On Error Resume Next
            Dim params(2) As Variant
            params(0) = Array("@kuldemeny_azon_13", azon13)
            params(1) = Array("@status_code", "HANDED_OVER")
            params(2) = Array("@changed_by", currentUser)
            modConnection.ExecuteSPNonQuery "usp_SetPackageStatus", params
            If Err.Number = 0 Then
                successCount = successCount + 1
            Else
                errorCount = errorCount + 1
            End If
            On Error GoTo 0
        End If
    Next li

    Dim msg As String
    msg = successCount & " tétel sikeresen HANDED_OVER státuszba helyezve."
    If errorCount > 0 Then
        msg = msg & vbCrLf & errorCount & " tétel hibával — ellenorizze a szallitonak_atadhato flag-et!"
    End If
    MsgBox msg, IIf(errorCount > 0, vbExclamation, vbInformation), "Átadás eredménye"

    ' Lista frissítése — átadott tételek eltunnek (ha az SP szurofeltétel kizárja oket)
    LoadList
End Sub

' ============================================================
' BEZÁRÁS
' ============================================================
Private Sub btnBezaras_Click()
    Unload Me
End Sub

' ============================================================
' SEGÉDMETÓDUSOK
' ============================================================

' countChecked — kijelölt sorok száma
Private Function countChecked() As Long
    Dim count As Long
    Dim li As Object
    For Each li In lstExport.ListItems
        If li.Checked Then count = count + 1
    Next li
    countChecked = count
End Function

' buildCheckedRecordset — kijelölt sorokból in-memory ADODB.Recordset
' Tartalmazza a listában látható összes mezot exporthoz.
Private Function buildCheckedRecordset() As Object
    On Error GoTo ErrHandler

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    ' Mezok hozzáadása (sorrendben a lista oszlopokkal)
    With rs.Fields
        .Append "kuldemeny_azonosito", 200, 255              ' adVarChar
        .Append "megrendelo_vevokod", 200, 20
        .Append "kuldemeny_aruertek", 200, 30
        .Append "kuldemeny_megjegyzes", 200, 255
        .Append "felrako_nev", 200, 100
        .Append "felrako_irsz", 200, 15
        .Append "felrako_varos", 200, 60
        .Append "felrako_kozterulet_elnevezese", 200, 60
        .Append "felrako_kozterulet_tipus", 200, 60
        .Append "felrako_hazszam", 200, 60
        .Append "felrako_megjegyzes", 200, 255
        .Append "felrako_kontakt_szemely", 200, 60
        .Append "felrako_telefonszam", 200, 25
        .Append "felrako_email", 200, 60
        .Append "felrako_idokapu", 200, 20
        .Append "cimzett_nev", 200, 100
        .Append "cimzett_irsz", 200, 15
        .Append "cimzett_varos", 200, 60
        .Append "cimzett_kozterulet_nev", 200, 60
        .Append "cimzett_kozterulet_jelleg", 200, 60
        .Append "cimzett_hazszam", 200, 60
        .Append "cimzett_megjegyzes", 200, 255
        .Append "cimzett_kontakt_szemely", 200, 60
        .Append "cimzett_telefonszam", 200, 25
        .Append "cimzett_email", 200, 60
        .Append "lerakas_idokapu", 200, 20
        .Append "raklap_hossza_cm", 200, 10
        .Append "raklap_szelessege_cm", 200, 10
        .Append "raklap_magassaga_cm", 200, 10
        .Append "kuldemeny_brutto_tomege_kg", 200, 15
        .Append "utanvet_osszege_ft", 200, 20
        .Append "ekaer", 200, 30
        .Append "adr_jelolo", 200, 30
        .Append "adr_pont", 200, 30
        .Append "szolgaltatas", 200, 60
        .Append "lerako_referencia", 200, 60
        .Append "felveteli_referencia", 200, 60
    End With

    rs.Open

    Dim li As Object
    For Each li In lstExport.ListItems
        If li.Checked Then
            rs.AddNew
            rs.Fields("kuldemeny_azonosito").Value = li.Text
            rs.Fields("megrendelo_vevokod").Value = li.SubItems(1)
            rs.Fields("kuldemeny_aruertek").Value = li.SubItems(2)
            rs.Fields("kuldemeny_megjegyzes").Value = li.SubItems(3)
            rs.Fields("felrako_nev").Value = li.SubItems(4)
            rs.Fields("felrako_irsz").Value = li.SubItems(5)
            rs.Fields("felrako_varos").Value = li.SubItems(6)
            rs.Fields("felrako_kozterulet_elnevezese").Value = li.SubItems(7)
            rs.Fields("felrako_kozterulet_tipus").Value = li.SubItems(8)
            rs.Fields("felrako_hazszam").Value = li.SubItems(9)
            rs.Fields("felrako_megjegyzes").Value = li.SubItems(10)
            rs.Fields("felrako_kontakt_szemely").Value = li.SubItems(11)
            rs.Fields("felrako_telefonszam").Value = li.SubItems(12)
            rs.Fields("felrako_email").Value = li.SubItems(13)
            rs.Fields("felrako_idokapu").Value = li.SubItems(14)
            rs.Fields("cimzett_nev").Value = li.SubItems(15)
            rs.Fields("cimzett_irsz").Value = li.SubItems(16)
            rs.Fields("cimzett_varos").Value = li.SubItems(17)
            rs.Fields("cimzett_kozterulet_nev").Value = li.SubItems(18)
            rs.Fields("cimzett_kozterulet_jelleg").Value = li.SubItems(19)
            rs.Fields("cimzett_hazszam").Value = li.SubItems(20)
            rs.Fields("cimzett_megjegyzes").Value = li.SubItems(21)
            rs.Fields("cimzett_kontakt_szemely").Value = li.SubItems(22)
            rs.Fields("cimzett_telefonszam").Value = li.SubItems(23)
            rs.Fields("cimzett_email").Value = li.SubItems(24)
            rs.Fields("lerakas_idokapu").Value = li.SubItems(25)
            rs.Fields("raklap_hossza_cm").Value = li.SubItems(26)
            rs.Fields("raklap_szelessege_cm").Value = li.SubItems(27)
            rs.Fields("raklap_magassaga_cm").Value = li.SubItems(28)
            rs.Fields("kuldemeny_brutto_tomege_kg").Value = li.SubItems(29)
            rs.Fields("utanvet_osszege_ft").Value = li.SubItems(30)
            rs.Fields("ekaer").Value = li.SubItems(31)
            rs.Fields("adr_jelolo").Value = li.SubItems(32)
            rs.Fields("adr_pont").Value = li.SubItems(33)
            rs.Fields("szolgaltatas").Value = li.SubItems(34)
            rs.Fields("lerako_referencia").Value = li.SubItems(35)
            rs.Fields("felveteli_referencia").Value = li.SubItems(36)
            rs.Update
        End If
    Next li

    rs.MoveFirst
    Set buildCheckedRecordset = rs
    Exit Function

ErrHandler:
    Set buildCheckedRecordset = Nothing
End Function

Public Function NvStr(val As Variant) As String
    If IsNull(val) Or IsEmpty(val) Then
        NvStr = ""
    Else
        NvStr = CStr(val)
    End If
End Function

