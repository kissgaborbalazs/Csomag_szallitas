Attribute VB_Name = "modConnection"
Option Explicit

' ============================================================
' modConnection — ADODB kapcsolatkezelés (late binding)
' Nincs szükség References beállításra.
' ============================================================

' -- ADO konstansok (late binding miatt szükséges) ------------
Private Const ADO_STATE_OPEN    As Integer = 1
Private Const ADO_STORED_PROC   As Integer = 4
Private Const ADO_PARAM_INPUT   As Integer = 1
Private Const ADO_PARAM_OUTPUT  As Integer = 2
Private Const ADO_VARCHAR       As Integer = 200
Private Const ADO_VARWCHAR      As Integer = 202   ' NVARCHAR (JSON delta)
Private Const ADO_INTEGER       As Integer = 3
Private Const ADO_DATE          As Integer = 7
Private Const ADO_BOOLEAN       As Integer = 11

Private m_conn As Object  ' ADODB.Connection singleton

' ============================================================
' Kapcsolat
' ============================================================

Public Function GetConnection() As Object
    If m_conn Is Nothing Then Set m_conn = CreateObject("ADODB.Connection")
    If m_conn.State <> ADO_STATE_OPEN Then
        m_conn.ConnectionString = "Provider=SQLOLEDB;" & _
            "Data Source=" & modConfig.DB_SERVER & ";" & _
            "Initial Catalog=" & modConfig.DB_NAME & ";" & _
            "Integrated Security=SSPI;"
        m_conn.Open
    End If
    Set GetConnection = m_conn
End Function

Public Sub CloseConnection()
    If Not m_conn Is Nothing Then
        If m_conn.State = ADO_STATE_OPEN Then m_conn.Close
        Set m_conn = Nothing
    End If
End Sub

Public Function TestConnection() As Boolean
    On Error GoTo Fail
    TestConnection = (GetConnection().State = ADO_STATE_OPEN)
    Exit Function
Fail:
    TestConnection = False
End Function

' ============================================================
' Belso segédek
' ============================================================

' params() = tömbök tömbje: Array("@nev", ertek), ...
Private Function BuildCmd(spName As String, params As Variant) As Object
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        Set .ActiveConnection = GetConnection()
        .CommandType = ADO_STORED_PROC
        .CommandText = spName
        If Not IsMissing(params) And Not IsEmpty(params) Then
            Dim i As Integer
            For i = 0 To UBound(params)
                .Parameters.Append BuildParam(cmd, CStr(params(i)(0)), params(i)(1))
            Next i
        End If
    End With
    Set BuildCmd = cmd
End Function

Private Function BuildParam(cmd As Object, name As String, val As Variant) As Object
    Dim adoType As Integer
    Dim size    As Integer

    If IsNull(val) Or IsEmpty(val) Then
        adoType = ADO_VARCHAR: size = 1
    ElseIf VarType(val) = vbDate Then
        adoType = ADO_DATE:    size = 0
    ElseIf VarType(val) = vbBoolean Then
        adoType = ADO_BOOLEAN: size = 0
    ElseIf VarType(val) = vbInteger Or VarType(val) = vbLong Then
        adoType = ADO_INTEGER: size = 0
    Else
        adoType = ADO_VARCHAR: size = IIf(Len(CStr(val)) < 255, 255, Len(CStr(val)) + 1)
    End If

    Set BuildParam = cmd.CreateParameter(name, adoType, ADO_PARAM_INPUT, size, val)
End Function

' ============================================================
' Publikus SP hívók
' ============================================================

' Olvasó SP ? Recordset
Public Function ExecuteSP(spName As String, Optional params As Variant) As Object
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open BuildCmd(spName, params)
    Set ExecuteSP = rs
End Function

' Író SP (nincs visszatéro RS)
Public Sub ExecuteSPNonQuery(spName As String, Optional params As Variant)
    BuildCmd(spName, params).Execute
End Sub

' Upsert SP — @conflict OUT paraméterrel
' SP: @kuldemeny_azonosito_13, @active_id, @json_delta, @conflict OUT
Public Sub ExecuteSPUpsert(spName As String, _
                            azon13 As String, _
                            activeId As Long, _
                            jsonDelta As String, _
                            ByRef conflict As Boolean)
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        Set .ActiveConnection = GetConnection()
        .CommandType = ADO_STORED_PROC
        .CommandText = spName
        .Parameters.Append .CreateParameter("@kuldemeny_azon_13", ADO_VARCHAR, ADO_PARAM_INPUT, 50, azon13)
        If activeId = 0 Then
            .Parameters.Append .CreateParameter("@current_active_id", ADO_INTEGER, ADO_PARAM_INPUT, 0, Null)
        Else
            .Parameters.Append .CreateParameter("@current_active_id", ADO_INTEGER, ADO_PARAM_INPUT, 0, activeId)
        End If
        .Parameters.Append .CreateParameter("@changed_by", ADO_VARCHAR, ADO_PARAM_INPUT, 10, GetCurrentUser())
        .Parameters.Append .CreateParameter("@changes_json", ADO_VARWCHAR, ADO_PARAM_INPUT, 4000, jsonDelta)
        .Parameters.Append .CreateParameter("@conflict", ADO_BOOLEAN, ADO_PARAM_OUTPUT, 0, False)
        .Execute
        conflict = CBool(.Parameters("@conflict").Value)
    End With
End Sub

' Export SP — @exported_count OUT + visszatéro RS
Public Function ExecuteSPExport(spName As String, _
                                 params As Variant, _
                                 ByRef exportedCount As Integer) As Object
    Dim cmd As Object
    Set cmd = BuildCmd(spName, params)
    cmd.Parameters.Append cmd.CreateParameter("@exported_count", ADO_INTEGER, ADO_PARAM_OUTPUT, 0, 0)

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd
    exportedCount = CInt(cmd.Parameters("@exported_count").Value)
    Set ExecuteSPExport = rs
End Function

Public Function ExecuteSPSetStatus(azon13 As String, _
                               statusCode As String, _
                               changedBy As String, _
                               note As String)
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        Set .ActiveConnection = GetConnection()
        .CommandType = ADO_STORED_PROC
        .CommandText = "usp_SetPackageStatus"
        .Parameters.Append .CreateParameter("@kuldemeny_azon_13", ADO_VARCHAR, ADO_PARAM_INPUT, 13, azon13)
        .Parameters.Append .CreateParameter("@status_code", ADO_VARCHAR, ADO_PARAM_INPUT, 50, statusCode)
        .Parameters.Append .CreateParameter("@changed_by", ADO_VARCHAR, ADO_PARAM_INPUT, 100, changedBy)
        .Parameters.Append .CreateParameter("@note", ADO_VARCHAR, ADO_PARAM_INPUT, 500, note)
        .Parameters.Append .CreateParameter("@new_status_id", ADO_INTEGER, ADO_PARAM_OUTPUT, 0, 0)
        .Execute
    End With
End Function

