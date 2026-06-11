VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPackageRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
' MultiUse = -1  'True
'End
Option Explicit

' ============================================================
' clsPackageRecord — Egy küldemény VIEW adatainak VBA tárolója
' Ütközésdetektáláshoz szükséges ID-k is tárolva.
' ============================================================

' -- Azonosítók ----------------------------------------------
Public KuldemenyAzon13          As String   ' join kulcs (D05)
Public KuldemenyAzonFull        As String   ' teljes azonosító

' -- Megrendelo ----------------------------------------------
Public MegrendeloVevokod        As String
Public MegrendeloCegnev         As String
Public MegbizasNapja            As Variant  ' Date or Null

' -- Státusz -------------------------------------------------
Public CurrentStatusCode        As String
Public CurrentStatusLabel       As String
Public StatusIsLocked           As Boolean  ' is_lock_trigger = 1
Public SzallitonaAtadhato       As Boolean  ' wszl_atadas_jelolve (D15)

' -- Ütközésdetektálás (D02) ---------------------------------
Public OperativeDataActiveId    As Long     ' aktív operative_data sor ID
Public ProcDataActiveId         As Long     ' aktív proc_data sor ID

' -- Felrakó adatok (op ?? RAW COALESCE) ---------------------
Public felrakoNev               As String
Public FelrakoIrsz              As String
Public FelrakoVaros             As String
Public FelrakoKozteruletNev     As String
Public FelrakoKozteruletTipus   As String
Public FelrakoHazszam           As String
Public FelrakoKontakt           As String
Public FelrakoTelefon           As String
Public FelrakoEmail             As String
Public FelrakoAtvetelDatum      As Variant  ' Date or Null
Public FelrakoIdokapu           As String
Public LerakasIdokapu           As String
Public FelrakoMegjegyzes        As String

' -- RAW értékek (csak tájékoztató, readonly) -----------------
Public RawAtvetelDatum          As String
Public RawIdokapu               As String

' -- Küldemény adatok ----------------------------------------
Public RaklapHossz              As String
Public RaklapSzeles             As String
Public RaklapMagas              As String
Public Ekaer                    As String
Public CimiratNyomtatas         As String
Public AdrJelolo                As String
Public AdrPont                  As String

' -- EFJ adatok (panel, readonly) ----------------------------
Public EfjId                    As String
Public EfeladojegyzekAzon       As String
Public BetoltesDatum            As String
Public CimzettNev               As String
Public CimzettIrsz              As String
Public CimzettVaros             As String
Public CimzettKozterulet        As String
Public CimzettHazszam           As String
Public CimzettKontakt           As String
Public CimzettTelefon           As String
Public CimzettEmail             As String
Public CimzettMegjegyzes        As String
Public KuldemenyBruttoTomeg     As String
Public UtanvetOsszege           As String
Public KuldemenyAruertek        As String
Public Szolgaltatas             As String
Public KezbesitesModja          As String
Public LerakoReferencia         As String
Public VezerAzon                As String

' ============================================================
' Feltöltés VIEW Recordset-bol
' ============================================================
Public Sub LoadFromRecordset(rs As Object)
    Dim f As Object: Set f = rs.Fields

    KuldemenyAzon13 = NS(f, "kuldemeny_azonosito_13")
    KuldemenyAzonFull = NS(f, "kuldemeny_azonosito")
    MegrendeloVevokod = NS(f, "megrendelo_vevokod")
    MegrendeloCegnev = NS(f, "megrendelo_cegnev")
    MegbizasNapja = NSD(f, "megbizas_napja")
    CurrentStatusCode = NS(f, "current_status_code")
    CurrentStatusLabel = NS(f, "current_status_label")
    StatusIsLocked = NSB(f, "status_is_locked")
    SzallitonaAtadhato = NSB(f, "szallitonak_atadhato")
    OperativeDataActiveId = NSL(f, "operative_data_active_id")
    ProcDataActiveId = NSL(f, "proc_data_active_id")
    felrakoNev = NS(f, "felrako_nev")
    FelrakoIrsz = NS(f, "felrako_irsz")
    FelrakoVaros = NS(f, "felrako_varos")
    FelrakoKozteruletNev = NS(f, "felrako_kozterulet_elnevezese")
    FelrakoKozteruletTipus = NS(f, "felrako_kozterulet_tipus")
    FelrakoHazszam = NS(f, "felrako_hazszam")
    FelrakoKontakt = NS(f, "felrako_kontakt_szemely")
    FelrakoTelefon = NS(f, "felrako_telefonszam")
    FelrakoEmail = NS(f, "felrako_email")
    FelrakoAtvetelDatum = NSD(f, "felrako_atveteli_datum")
    FelrakoIdokapu = NS(f, "felrako_idokapu")
    LerakasIdokapu = NS(f, "lerakas_idokapu")
    FelrakoMegjegyzes = NS(f, "felrako_megjegyzes")
    RawAtvetelDatum = NS(f, "raw_atveteli_datum")
    RawIdokapu = NS(f, "raw_idokapu")
    RaklapHossz = NS(f, "raklap_hossza_cm")
    RaklapSzeles = NS(f, "raklap_szelessege_cm")
    RaklapMagas = NS(f, "raklap_magassaga_cm")
    Ekaer = NS(f, "ekaer")
    CimiratNyomtatas = BitToIgenNem(f, "cimirat_nyomtatas")
    AdrJelolo = NS(f, "adr_jelolo")
    AdrPont = NS(f, "adr_pont")
    EfjId = NS(f, "efj_id")
    EfeladojegyzekAzon = NS(f, "efeladojegyzek_azon")
    BetoltesDatum = NS(f, "betoltes_datum")
    CimzettNev = NS(f, "cimzett_nev")
    CimzettIrsz = NS(f, "cimzett_irsz")
    CimzettVaros = NS(f, "cimzett_varos")
    CimzettKozterulet = NS(f, "cimzett_kozterulet")
    CimzettHazszam = NS(f, "cimzett_hazszam")
    CimzettKontakt = NS(f, "cimzett_kontakt_szemely")
    CimzettTelefon = NS(f, "cimzett_telefonszam")
    CimzettEmail = NS(f, "cimzett_email")
    CimzettMegjegyzes = NS(f, "cimzett_megjegyzes")
    KuldemenyBruttoTomeg = NS(f, "kuldemeny_brutto_tomege_kg")
    UtanvetOsszege = NS(f, "utanvet_osszege_ft")
    KuldemenyAruertek = NS(f, "kuldemeny_aruertek")
    Szolgaltatas = NS(f, "szolgaltatas")
    KezbesitesModja = NS(f, "kezbesites_modja")
    LerakoReferencia = NS(f, "lerako_referencia_szam")
    VezerAzon = NS(f, "vezer_azon")
End Sub

Private Function BitToIgenNem(f As Object, name As String) As String
    On Error Resume Next
    Dim v As Variant: v = f(name).Value
    If IsNull(v) Or IsEmpty(v) Then
        BitToIgenNem = ""
    ElseIf CBool(v) Then
        BitToIgenNem = "igen"
    Else
        BitToIgenNem = "nem"
    End If
End Function

' -- Null-safe olvasók ----------------------------------------
Private Function NS(f As Object, name As String) As String
    On Error Resume Next
    Dim v As Variant: v = f(name).Value
    NS = IIf(IsNull(v), "", CStr(v))
End Function

Private Function NSD(f As Object, name As String) As Variant
    On Error Resume Next
    Dim v As Variant: v = f(name).Value
    NSD = IIf(IsNull(v), Null, v)
End Function

Private Function NSB(f As Object, name As String) As Boolean
    On Error Resume Next
    Dim v As Variant: v = f(name).Value
    NSB = IIf(IsNull(v), False, CBool(v))
End Function

Private Function NSL(f As Object, name As String) As Long
    On Error Resume Next
    Dim v As Variant: v = f(name).Value
    NSL = IIf(IsNull(v), 0, CLng(v))
End Function
