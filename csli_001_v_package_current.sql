USE [RPA_Processes_DEV]
GO

/****** Object:  View [dbo].[csli_001_v_package_current]    Script Date: 2026. 06. 11. 11:07:37 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



-- A feldolgozási adatok blokkhoz add hozzá:
--
--     pd.wszl_atadas_jelolve,
--
-- A meglévő pd.* mezők mellé, például:

CREATE   VIEW [dbo].[csli_001_v_package_current] AS
SELECT
    -- Azonosítók
    ro.id                                           AS raw_order_id,
    COALESCE(efj.kuldemeny_azonosito ,			ro.kuldemeny_azonosito)						AS kuldemeny_azonosito,
    ro.kuldemeny_azonosito_13,
    -- Megrendelő
    ro.megrendelo_vevokod,
    ro.megrendelo_cegnev,
    -- Megbízás
    ro.megbizas_napja,
    ro.megbizas_ideje,
    -- Felvételi adatok — COALESCE: operative_data felülírja RAW-t ha nem NULL
    COALESCE(od.felrako_nev,                    ro.felrako_nev)                             AS felrako_nev,
    COALESCE(od.felrako_irsz,                   ro.felrako_irsz)                            AS felrako_irsz,
    COALESCE(od.felrako_varos,                  ro.felrako_varos)                           AS felrako_varos,
    COALESCE(od.felrako_kozterulet_elnevezese,  ro.felrako_kozterulet_elnevezese)           AS felrako_kozterulet_elnevezese,
    COALESCE(od.felrako_kozterulet_tipus,       ro.felrako_kozterulet_tipus_utca_ter_ut)    AS felrako_kozterulet_tipus,
    COALESCE(od.felrako_hazszam,                ro.felrako_hazszam)                         AS felrako_hazszam,
    COALESCE(od.felrako_megjegyzes,             ro.felrako_megjegyzes)                      AS felrako_megjegyzes,
    COALESCE(od.felrako_atveteli_datum,         ro.felrako_atveteli_datum)                  AS felrako_atveteli_datum,
    COALESCE(od.felrako_idokapu,                ro.felrako_idokapu)                         AS felrako_idokapu,
    COALESCE(od.felrako_kontakt_szemely,        ro.felrako_kontakt_szemely)                 AS felrako_kontakt_szemely,
    COALESCE(od.felrako_telefonszam,            ro.felrako_telefonszam)                     AS felrako_telefonszam,
    COALESCE(od.felrako_email,                  ro.felrako_email)                           AS felrako_email,
    -- Küldemény adatok — COALESCE
    COALESCE(od.raklap_hossza_cm,               ro.raklap_hossza_cm)                        AS raklap_hossza_cm,
    COALESCE(od.raklap_szelessege_cm,           ro.raklap_szelessege_cm)                    AS raklap_szelessege_cm,
    COALESCE(od.raklap_magassaga_cm,            ro.raklap_magassaga_cm)                     AS raklap_magassaga_cm,
    COALESCE(od.ekaer,                          ro.ekaer)                                   AS ekaer,
    COALESCE(od.cimirat_nyomtatas,              ro.cimirat_nyomtatas)                       AS cimirat_nyomtatas,
    COALESCE(od.adr_jelolo,                     ro.adr_jelolo)                              AS adr_jelolo,
    COALESCE(od.adr_pont,                       ro.adr_pont)                                AS adr_pont,
    COALESCE(od.lerakas_idokapu,                ro.lerakas_idokapu)                         AS lerakas_idokapu,
    -- EFJ / Cimzett adatok
    efj.id                                          AS raw_efj_id,
    efj.efj_id,
    efj.efeladojegyzek_azon,
    efj.betoltes_datum                              AS efj_betoltes_datum,
    efj.varhato_feladas_datum,
    efj.kuldemeny_brutto_tomege_kg,
	efj.szolgaltatas,
    efj.utanvet_osszege_ft,
    efj.kuldemeny_aruertek,
    efj.kezbesites_modja,
    efj.lerako_referencia_szam,
    efj.cimzett_nev,
    efj.cimzett_irsz,
    efj.cimzett_varos,
    efj.cimzett_kozterulet_nev,
    efj.cimzett_kozterulet_jelleg,
    efj.cimzett_hazszam,
    efj.cimzett_megjegyzes,
    efj.cimzett_kontakt_szemely,
    efj.cimzett_email,
    efj.cimzett_telefonszam,
    -- Feldolgozási adatok
    pd.tervezett_felvetel_datum,
    pd.sikertelen_felvetel_megj,
    pd.ekaer_kuldheto,
    pd.szamlazhato,
    pd.indokolatlan_kiallas,
    COALESCE(pd.szallitonak_atadhato, 0)			AS szallitonak_atadhato,
    -- Aktuális státusz
    sl.status_code                                  AS current_status_code,
    sl.status_label_hu                              AS current_status_label,
    sl.is_lock_trigger                              AS status_is_locked,
    ps_latest.changed_at                            AS status_changed_at,
    ps_latest.changed_by                            AS status_changed_by,
    -- Ütközésdetektáláshoz
    od.id                                           AS operative_data_active_id,
    pd.id                                           AS proc_data_active_id
FROM dbo.csli_001_raw_orders ro
LEFT JOIN dbo.csli_001_raw_efj efj
    ON  efj.kuldemeny_azonosito_13 = ro.kuldemeny_azonosito_13
    AND efj.obsolete = 0
LEFT JOIN dbo.csli_001_operative_data_orders od
    ON  od.kuldemeny_azonosito_13 = ro.kuldemeny_azonosito_13
    AND od.obsolete = 0
LEFT JOIN dbo.csli_001_proc_data pd
    ON  pd.kuldemeny_azonosito_13 = ro.kuldemeny_azonosito_13
    AND pd.obsolete = 0
LEFT JOIN (
    SELECT *,
           ROW_NUMBER() OVER (
               PARTITION BY kuldemeny_azonosito_13
               ORDER BY package_status_id DESC
           ) AS rn
    FROM dbo.csli_001_package_status
) ps_latest
    ON  ps_latest.kuldemeny_azonosito_13 = ro.kuldemeny_azonosito_13
    AND ps_latest.rn = 1
LEFT JOIN dbo.csli_001_status_lookup sl
    ON  sl.status_id = ps_latest.status_id
WHERE ro.obsolete = 0;
GO

