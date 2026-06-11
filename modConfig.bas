Attribute VB_Name = "modConfig"
Option Explicit

' ============================================================
' modConfig — Konfiguráció és FIELD_CONFIG
' ============================================================

' -- Kapcsolat (BEÁLLÍTANDÓ) ----------------------------------
Public Const DB_SERVER  As String = "DEVSQL19PRIO2\DEVSQL19DB02"
Public Const DB_NAME    As String = "RPA_Processes_DEV"

' -- Alkalmazás -----------------------------------------------
Public Const APP_VERSION As String = "v0.1"

' -- Aktuális Windows-felhasználó -----------------------------
Public Function GetCurrentUser() As String
    GetCurrentUser = Environ("USERDOMAIN") & "\" & Environ("USERNAME")
End Function

' -- Mezo szerkeszthetoségi konfiguráció ----------------------
' True  = aktív (fehér háttér, szerkesztheto)
' False = passzív (szürke háttér, readonly)
Public Function GetFieldConfig() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    ' Szerkesztheto mezok
    d("felrako_kontakt_szemely") = True
    d("felrako_telefonszam") = True
    d("felrako_email") = True
    d("felrako_nev") = True
    d("felrako_irsz") = True
    d("felrako_varos") = True
    d("felrako_kozterulet_elnevezese") = True
    d("felrako_kozterulet_tipus") = True
    d("felrako_hazszam") = True
    d("felrako_megjegyzes") = True
    d("felrako_atveteli_datum") = True
    d("felrako_idokapu") = True
    d("lerakas_idokapu") = True
    d("raklap_hossza_cm") = True
    d("raklap_szelessege_cm") = True
    d("raklap_magassaga_cm") = True
    d("adr_jelolo") = True
    d("adr_pont") = True
    d("cimirat_nyomtatas") = True

    ' Mindig readonly
    d("megrendelo_vevokod") = False
    d("megrendelo_cegnev") = False
    d("kuldemeny_azonosito") = False
    d("ekaer") = False

    Set GetFieldConfig = d
End Function
