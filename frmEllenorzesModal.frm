VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEllenorzesModal 
   Caption         =   "UserForm1"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13845
   OleObjectBlob   =   "frmEllenorzesModal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEllenorzesModal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmEllenorzesModal
'
' Builder (modBuildEllenorzesModal) hozza létre a kontrollokat
' teljes pozíció/design beállítással.
'
' Az Initialize-ban marad:
'   - cboCimirat.AddItem feltöltés
'   - .Tag értékek beállítása (DB oszlopnév mapping)
'   - Caption (form cím)
'   - BackColor (form háttér)
' ============================================================

Private Const CLR_FORM_BG As Long = &HFAFAFA

Private m_pkg As clsPackageRecord

' ============================================================
' Publikus belépõ
' ============================================================
Public Sub LoadOrder(azon13 As String)
    On Error GoTo ErrHnd
    Dim params(0) As Variant
    params(0) = Array("@kuldemeny_azonosito_13", azon13)
    Dim rs As Object
    Set rs = modConnection.ExecuteSP("usp_GetPackageDetail", params)
    If rs.BOF And rs.EOF Then
        MsgBox "A küldeményt nem találtam: " & azon13, vbExclamation
        Exit Sub
    End If
    Set m_pkg = New clsPackageRecord
    m_pkg.LoadFromRecordset rs
    FillForm
    Exit Sub
ErrHnd:
    modUtils.ShowError "Küldemény betöltés", Err.Description
End Sub

Private Sub btnEfjEletut_Click()
    If m_pkg Is Nothing Then Exit Sub
    frmHistory.LoadOrder (m_pkg.KuldemenyAzonFull)
    frmHistory.Show vbModal
End Sub

' ============================================================
' Initialize – csak ami builder-bõl NEM vihetõ át
' ============================================================
Private Sub UserForm_Initialize()
    Me.Caption = "Megrendelés adatok"
    Me.BackColor = CLR_FORM_BG
    Me.StartUpPosition = 1

    ' -- Tag mapping (DB oszlop › kontroll) --
    txtKontakt.tag = "felrako_kontakt_szemely"
    txtTelefon.tag = "felrako_telefonszam"
    txtEmail.tag = "felrako_email"
    txtFelrakoNev.tag = "felrako_nev"
    txtFelrakoIrsz.tag = "felrako_irsz"
    txtFelrakoVaros.tag = "felrako_varos"
    txtKozteruletNev.tag = "felrako_kozterulet_elnevezese"
    txtKozteruletTipus.tag = "felrako_kozterulet_tipus"
    txtHazszam.tag = "felrako_hazszam"
    txtMegjegyzes.tag = "felrako_megjegyzes"
    txtAtveteli.tag = "felrako_atveteli_datum"
    txtIdokapu.tag = "felrako_idokapu"
    txtLerakasIdokapu.tag = "lerakas_idokapu"
    txtCsomagAzon.tag = "kuldemeny_azonosito"
    txtHossz.tag = "raklap_hossza_cm"
    txtSzeles.tag = "raklap_szelessege_cm"
    txtMagas.tag = "raklap_magassaga_cm"
    txtEkaer.tag = "ekaer"
    cboCimirat.tag = "cimirat_nyomtatas"
    txtAdrJelolo.tag = "adr_jelolo"
    txtAdrPont.tag = "adr_pont"

    ' -- ComboBox feltöltés --
    cboCimirat.AddItem "igen"
    cboCimirat.AddItem "nem"
End Sub

' ============================================================
' Form feltöltés
' ============================================================
Private Sub FillForm()
    txtStatus.Value = m_pkg.CurrentStatusLabel
    txtEfjAzonInfo.Value = m_pkg.EfjId
    txtVevokod.Value = m_pkg.MegrendeloVevokod
    txtCegnev.Value = m_pkg.MegrendeloCegnev
    txtKontakt.Text = m_pkg.FelrakoKontakt
    txtTelefon.Text = m_pkg.FelrakoTelefon
    txtEmail.Text = m_pkg.FelrakoEmail
    txtRawAtveteli.Value = m_pkg.OperativeDataActiveId
    txtRawIdokapu.Value = m_pkg.ProcDataActiveId

    txtFelrakoNev.Text = m_pkg.felrakoNev
    txtFelrakoIrsz.Text = m_pkg.FelrakoIrsz
    txtFelrakoVaros.Text = m_pkg.FelrakoVaros
    txtKozteruletNev.Text = m_pkg.FelrakoKozteruletNev
    txtKozteruletTipus.Text = m_pkg.FelrakoKozteruletTipus
    txtHazszam.Text = m_pkg.FelrakoHazszam
    txtMegjegyzes.Text = m_pkg.FelrakoMegjegyzes
    txtAtveteli.Text = FormatDate(m_pkg.FelrakoAtvetelDatum)
    txtIdokapu.Text = m_pkg.FelrakoIdokapu
    txtLerakasIdokapu.Text = m_pkg.LerakasIdokapu

    txtCsomagAzon.Value = m_pkg.KuldemenyAzonFull
    txtHossz.Text = m_pkg.RaklapHossz
    txtSzeles.Text = m_pkg.RaklapSzeles
    txtMagas.Text = m_pkg.RaklapMagas
    txtEkaer.Value = m_pkg.Ekaer
    SetComboValue cboCimirat, m_pkg.CimiratNyomtatas
    txtAdrJelolo.Text = m_pkg.AdrJelolo
    txtAdrPont.Text = m_pkg.AdrPont

    chkEllenorizve.Value = m_pkg.SzallitonaAtadhato
    ApplyLockState
End Sub

Private Sub ApplyLockState()
    modUtils.ApplyFieldConfig Me, m_pkg.StatusIsLocked
    chkEllenorizve.Enabled = Not m_pkg.StatusIsLocked
    btnStornoOrder.Enabled = Not m_pkg.StatusIsLocked
    btnMentes.Enabled = Not m_pkg.StatusIsLocked
    btnEfjReszletek.Enabled = True    ' EFJ gomb mindig aktív
    If m_pkg.StatusIsLocked Then
        txtStatus.BackColor = RGB(255, 235, 156)
    End If
End Sub

' ============================================================
' Segédek
' ============================================================
Private Sub SetComboValue(cbo As Object, val As String)
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = val Then cbo.ListIndex = i: Exit Sub
    Next i
    cbo.ListIndex = -1
End Sub

Private Function FormatDate(val As Variant) As String
    If IsNull(val) Or IsEmpty(val) Then FormatDate = "": Exit Function
    On Error Resume Next
    FormatDate = Format(CDate(val), "yyyy-mm-dd")
End Function

' ============================================================
' Eseménykezelõk
' ============================================================
Private Sub btnEfjReszletek_Click()
    If m_pkg Is Nothing Then Exit Sub
    frmRawEfjModal.LoadRecord m_pkg.KuldemenyAzonFull
    frmRawEfjModal.Show vbModal
End Sub

Private Sub btnMentes_Click()
    On Error GoTo ErrHnd
    Dim opDict As Object
    Set opDict = CreateObject("Scripting.Dictionary")
    CollectIfChanged opDict, "felrako_kontakt_szemely", txtKontakt.Text, m_pkg.FelrakoKontakt
    CollectIfChanged opDict, "felrako_telefonszam", txtTelefon.Text, m_pkg.FelrakoTelefon
    CollectIfChanged opDict, "felrako_email", txtEmail.Text, m_pkg.FelrakoEmail
    CollectIfChanged opDict, "felrako_nev", txtFelrakoNev.Text, m_pkg.felrakoNev
    CollectIfChanged opDict, "felrako_irsz", txtFelrakoIrsz.Text, m_pkg.FelrakoIrsz
    CollectIfChanged opDict, "felrako_varos", txtFelrakoVaros.Text, m_pkg.FelrakoVaros
    CollectIfChanged opDict, "felrako_kozterulet_elnevezese", txtKozteruletNev.Text, m_pkg.FelrakoKozteruletNev
    CollectIfChanged opDict, "felrako_kozterulet_tipus", txtKozteruletTipus.Text, m_pkg.FelrakoKozteruletTipus
    CollectIfChanged opDict, "felrako_hazszam", txtHazszam.Text, m_pkg.FelrakoHazszam
    CollectIfChanged opDict, "felrako_megjegyzes", txtMegjegyzes.Text, m_pkg.FelrakoMegjegyzes
    CollectIfChanged opDict, "felrako_atveteli_datum", txtAtveteli.Text, FormatDate(m_pkg.FelrakoAtvetelDatum)
    CollectIfChanged opDict, "felrako_idokapu", txtIdokapu.Text, m_pkg.FelrakoIdokapu
    CollectIfChanged opDict, "lerakas_idokapu", txtLerakasIdokapu.Text, m_pkg.LerakasIdokapu
    CollectIfChanged opDict, "raklap_hossza_cm", txtHossz.Text, m_pkg.RaklapHossz
    CollectIfChanged opDict, "raklap_szelessege_cm", txtSzeles.Text, m_pkg.RaklapSzeles
    CollectIfChanged opDict, "raklap_magassaga_cm", txtMagas.Text, m_pkg.RaklapMagas
    CollectIfChanged opDict, "cimirat_nyomtatas", cboCimirat.Text, m_pkg.CimiratNyomtatas
    CollectIfChanged opDict, "adr_jelolo", txtAdrJelolo.Text, m_pkg.AdrJelolo
    CollectIfChanged opDict, "adr_pont", txtAdrPont.Text, m_pkg.AdrPont

    If opDict.count > 0 Then
        Dim opConflict As Boolean
        modConnection.ExecuteSPUpsert "usp_UpsertOperativeData", _
            m_pkg.KuldemenyAzon13, m_pkg.OperativeDataActiveId, _
            modUtils.DictToJson(opDict), opConflict
        If opConflict Then
            modUtils.ShowConflictWarning
            LoadOrder m_pkg.KuldemenyAzon13
            Exit Sub
        End If
    End If

    If m_pkg.SzallitonaAtadhato = chkEllenorizve.Value Then
        Me.Hide
        Exit Sub
    End If

    Dim procDict As Object
    Set procDict = CreateObject("Scripting.Dictionary")
    procDict("szallitonak_atadhato") = IIf(chkEllenorizve.Value, "1", "0")

    Dim procConflict As Boolean
    modConnection.ExecuteSPUpsert "usp_UpsertProcData", _
        m_pkg.KuldemenyAzon13, m_pkg.ProcDataActiveId, _
        modUtils.DictToJson(procDict), procConflict
    If procConflict Then
        modUtils.ShowConflictWarning
        LoadOrder m_pkg.KuldemenyAzon13
        Exit Sub
    End If

    Me.Hide
    Exit Sub
ErrHnd:
    modUtils.ShowError "Mentés", Err.Description
End Sub

Private Sub CollectIfChanged(dict As Object, key As String, _
                              cur As String, orig As String)
    If cur <> orig Then dict(key) = cur
End Sub

Private Sub btnStornoOrder_Click()
    If m_pkg Is Nothing Then Exit Sub
    If MsgBox("Biztosan stornózza: " & m_pkg.KuldemenyAzon13 & "?" & vbCrLf & _
              "CANCELLED státusz kerül rögzítésre.", _
              vbYesNo + vbExclamation, "Stornó megerõsítés") <> vbYes Then Exit Sub
    On Error GoTo ErrHnd
    Dim params(1) As Variant
    params(0) = Array("@kuldemeny_azonosito_13", m_pkg.KuldemenyAzon13)
    params(1) = Array("@status_code", "CANCELLED")
    modConnection.ExecuteSPNonQuery "usp_SetPackageStatus", params
    Me.Hide
    Exit Sub
ErrHnd:
    modUtils.ShowError "Stornó", Err.Description
End Sub

Private Sub btnMegse_Click(): Me.Hide: End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Me.Hide
    Cancel = True
End Sub




