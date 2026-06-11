VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHistory 
   Caption         =   "UserForm1"
   ClientHeight    =   11820
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   19995
   OleObjectBlob   =   "frmHistory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmHistory — Életút modal
' Hívás: frmHistory.LoadOrder "DHL-2026-001234"
'         frmHistory.Show vbModal
' ============================================================

Private m_azon13 As String

Private Sub lblAzonVal_Click()

End Sub

Private Sub lstStatus_BeforeLabelEdit(Cancel As Integer)

End Sub

' ============================================================
' INITIALIZE
' ============================================================

Private Sub UserForm_Initialize()
    Me.Caption = "Életút"
    'Me.Width = 1000
    'Me.Height = 600
    
    SetupListColumns
End Sub

Private Sub SetupListColumns()

    With lstStatus
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Dátum/idõ", 85
        .ColumnHeaders.Add , , "Módosító", 85
        .ColumnHeaders.Add , , "Státusz", 160
        .ColumnHeaders.Add , , "Megjegyzés", 150     ' nyújtható — utolsó oszlop
    End With

    With lstFields
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Dátum/idõ", 85
        .ColumnHeaders.Add , , "Módosító", 84
        .ColumnHeaders.Add , , "Mezõ", 150
        .ColumnHeaders.Add , , "Elõtte", 130
        .ColumnHeaders.Add , , "Utána", 130          ' nyújtható — utolsó oszlop
    End With

End Sub

' ============================================================
' PUBLIC BELÉPÉSI PONT
' ============================================================

Public Sub LoadOrder(azon As String)
    m_azon13 = Left(azon, 13)
    lblAzonVal.Caption = azon
    RefreshHistory
End Sub

' ============================================================
' BETÖLTÉS
' ============================================================

Private Sub RefreshHistory()

    Dim params(0) As Variant
    params(0) = Array("@kuldemeny_azon_13", m_azon13)

    Dim rs As Object
    Set rs = modConnection.ExecuteSP("usp_GetPackageHistory", params)

    lstStatus.ListItems.Clear
    lstFields.ListItems.Clear

    Dim statusCount As Integer: statusCount = 0
    Dim fieldCount  As Integer: fieldCount = 0

    Do While Not rs.EOF

        Dim evtTime  As String: evtTime = Left(Format(rs("event_time").Value, "yyyy-mm-dd hh:nn"), 19)
        Dim evtActor As String: evtActor = Nz(rs("actor").Value, "—")

        Select Case rs("event_type").Value

            Case "STATUS_CHANGE"
                Dim sItem As ListItem
                Set sItem = lstStatus.ListItems.Add
                sItem.Text = evtTime
                sItem.SubItems(1) = evtActor
                sItem.SubItems(2) = Nz(rs("new_value").Value, "")
                sItem.SubItems(3) = Nz(rs("note").Value, "")
                statusCount = statusCount + 1

            Case "FIELD_CHANGE"
                Dim parts() As String
                parts = Split(Nz(rs("event_label").Value, " / "), " / ")

                Dim fItem As ListItem
                Set fItem = lstFields.ListItems.Add
                fItem.Text = evtTime
                fItem.SubItems(1) = evtActor
                fItem.SubItems(2) = parts(UBound(parts))
                fItem.SubItems(3) = NullOrEmpty(rs("old_value").Value)
                fItem.SubItems(4) = NullOrEmpty(rs("new_value").Value)
                fieldCount = fieldCount + 1

        End Select

        rs.MoveNext
    Loop

    rs.Close

    lblStatusHdr.Caption = "STÁTUSZOK (" & statusCount & ")"
    lblFieldsHdr.Caption = "MEZÕVÁLTOZÁSOK (" & fieldCount & ")"

End Sub

' ============================================================
' SEGÉDFÜGGVÉNY
' ============================================================

Private Function NullOrEmpty(val As Variant) As String
    If IsNull(val) Or val = "" Then
        NullOrEmpty = "-"
    Else
        NullOrEmpty = CStr(val)
    End If
End Function

' ============================================================
' ESEMÉNYKEZELÕK
' ============================================================

Private Sub btnClose_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' X gomb = ugyanolyan mint Bezárás, nem Unload
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Function Nz(val As Variant, Optional default As String = "") As String
    
    If IsNull(val) Or IsEmpty(val) Then
        Nz = default
    Else
        Nz = CStr(val)
    End If
End Function

