Attribute VB_Name = "modUtils"
Option Explicit

' ============================================================
' modUtils — Általános segédfüggvények
' ============================================================

Private Const CLR_EDITABLE  As Long = &HFFFFFF   ' fehér
Private Const CLR_READONLY  As Long = &HF0F0F0   ' világosszürke
Private Const CLR_HDR_BG    As Long = &HD0D0D0   ' fejléc szürke

' ============================================================
' JSON delta builder (D04)
' ============================================================
Public Function DictToJson(dict As Object) As String
    If dict.count = 0 Then DictToJson = "{}": Exit Function

    Dim parts() As String
    ReDim parts(dict.count - 1)
    Dim i As Integer: i = 0
    Dim key As Variant

    For Each key In dict.Keys
        Dim val As Variant: val = dict(key)
        If IsNull(val) Or CStr(val) = "" Then
            parts(i) = """" & key & """:null"
        ElseIf IsDate(val) Then
            parts(i) = """" & key & """:""" & Format(val, "yyyy-mm-dd") & """"
        Else
            Dim s As String: s = CStr(val)
            s = Replace(s, "\", "\\")
            s = Replace(s, """", "\""")
            parts(i) = """" & key & """:""" & s & """"
        End If
        i = i + 1
    Next key

    DictToJson = "{" & Join(parts, ",") & "}"
End Function

' ============================================================
' Dátum null-safe konverzió
' ============================================================
Public Function SafeDate(val As Variant) As Variant
    If IsNull(val) Or Trim(CStr(val)) = "" Then
        SafeDate = Null
    ElseIf IsDate(val) Then
        SafeDate = CDate(val)
    Else
        SafeDate = Null
    End If
End Function

' ============================================================
' Hibaüzenetek
' ============================================================
Public Sub ShowError(context As String, errMsg As String)
    MsgBox "Hiba: " & context & vbCrLf & vbCrLf & errMsg, _
           vbCritical, "Csomagstátusz Kezelo"
End Sub

Public Sub ShowConflictWarning()
    MsgBox "Az adatokat közben más felhasználó módosította." & vbCrLf & _
           "Az urlap újratöltésre kerül — kérjük ismételje meg a módosítást.", _
           vbExclamation, "Mentési ütközés"
End Sub

' ============================================================
' FIELD_CONFIG alkalmazása UserForm összes kontrolljára (D14)
' locked = True ? minden mezo passzív (pl. HANDED_OVER után)
' ============================================================
Public Sub ApplyFieldConfig(frm As Object, locked As Boolean)
    Dim cfg As Object
    Set cfg = modConfig.GetFieldConfig()
    Dim ctrl As Object

    For Each ctrl In frm.Controls
        Dim ctrlType As String: ctrlType = TypeName(ctrl)
        If ctrlType = "TextBox" Or ctrlType = "ComboBox" Then
            If cfg.Exists(ctrl.tag) Then
                Dim editable As Boolean
                editable = CBool(cfg(ctrl.tag)) And Not locked
                ctrl.Enabled = editable
                ctrl.locked = Not editable
                ctrl.BackColor = IIf(editable, CLR_EDITABLE, CLR_READONLY)
            End If
        End If
    Next ctrl
End Sub

' ============================================================
' ListBox fejléc beállítás
' lstHdr  = disabled ListBox a fejléc sornak
' headers = String tömb az oszlopnevekkel
' colWidths = pl. "120 pt;80 pt;150 pt;80 pt;120 pt"
' ============================================================
Public Sub SetListBoxHeader(lstHdr As Object, headers() As String, colWidths As String)
    With lstHdr
        .ColumnCount = UBound(headers) + 1
        .ColumnWidths = colWidths
        .Enabled = False
        .BackColor = CLR_HDR_BG
        .ForeColor = RGB(50, 50, 50)
        .Clear
        .AddItem headers(0)
        Dim i As Integer
        For i = 1 To UBound(headers)
            .List(0, i) = headers(i)
        Next i
    End With
End Sub

' ============================================================
' ListBox feltöltés Recordset-bol
' A ListBox ColumnCount-ját a hívó kód állítja be.
' ============================================================
Public Sub FillListBox(lst As Object, rs As Object)
    lst.Clear
    If rs Is Nothing Then Exit Sub
    If rs.BOF And rs.EOF Then Exit Sub

    'rs.MoveFirst
    Dim r As Integer: r = 0

    Do While Not rs.EOF
        Dim firstVal As Variant: firstVal = rs.Fields(0).Value
        If IsNull(firstVal) Then lst.AddItem "" Else lst.AddItem CStr(firstVal)

        Dim c As Integer
        For c = 1 To rs.Fields.count - 1
            Dim v As Variant: v = rs.Fields(c).Value
            If IsNull(v) Then lst.List(r, c) = "" Else lst.List(r, c) = CStr(v)
        Next c

        r = r + 1
        rs.MoveNext
    Loop
End Sub
