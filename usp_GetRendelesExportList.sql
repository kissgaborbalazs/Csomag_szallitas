USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_GetRendelesExportList]    Script Date: 2026. 06. 11. 11:11:03 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE OR ALTER     PROCEDURE [dbo].[usp_GetRendelesExportList]
    @vevokod_filter     NVARCHAR(10)  = NULL,
    @kuldemeny_filter   NVARCHAR(26)  = NULL,
    @date_from          DATE          = NULL
AS
BEGIN
    SET NOCOUNT ON;

    IF @vevokod_filter   = '' SET @vevokod_filter   = NULL;
    IF @kuldemeny_filter = '' SET @kuldemeny_filter = NULL;

    SELECT
        v.kuldemeny_azonosito,
        v.kuldemeny_aruertek,
        CAST(NULL AS NVARCHAR(50))  AS kuldemeny_megjegyzes,
        v.felrako_nev,
        v.felrako_irsz,
        v.felrako_varos,
        v.felrako_kozterulet_elnevezese,
        v.felrako_kozterulet_tipus,
        v.felrako_hazszam,
        v.felrako_megjegyzes,
        v.felrako_kontakt_szemely,
        v.felrako_telefonszam,
        v.felrako_email,
        v.cimzett_nev,
        v.cimzett_irsz,
        v.cimzett_varos,
        v.cimzett_kozterulet_nev,
        v.cimzett_kozterulet_jelleg,
        v.cimzett_hazszam,
        v.cimzett_megjegyzes,
        v.cimzett_kontakt_szemely,
        v.cimzett_telefonszam,
        v.cimzett_email,
        v.raklap_hossza_cm,
        v.raklap_szelessege_cm,
        v.raklap_magassaga_cm,
        v.kuldemeny_brutto_tomege_kg,
        v.utanvet_osszege_ft,
        v.ekaer,
        v.adr_jelolo,
        v.adr_pont,
        v.szolgaltatas,
        v.megrendelo_vevokod,
        CAST(NULL AS NVARCHAR(50))  AS lerako_referencia,
        CAST(NULL AS NVARCHAR(50))  AS felveteli_referencia,
        v.felrako_idokapu,
        v.lerakas_idokapu,
        -- VBA ütközésdetektáláshoz: aktív proc_data sor ID
        -- NULL ha még nincs proc_data sor (első rögzítés lesz)
        pd.id                       AS proc_data_id
    FROM dbo.csli_001_v_package_current v
    LEFT JOIN dbo.csli_001_proc_data pd
           ON pd.kuldemeny_azonosito_13 = v.kuldemeny_azonosito_13
          AND pd.obsolete = 0
    WHERE
        COALESCE(v.szallitonak_atadhato, 0) = 1
        AND v.current_status_code  = 'EFJ_RECEIVED'
        AND v.raw_efj_id            IS NOT NULL
        AND v.status_is_locked      = 0
        AND (@vevokod_filter   IS NULL OR v.megrendelo_vevokod  =      @vevokod_filter)
        AND (@kuldemeny_filter IS NULL OR v.kuldemeny_azonosito LIKE '%' + @kuldemeny_filter + '%')
        AND (@date_from        IS NULL OR v.felrako_atveteli_datum >= @date_from)
    ORDER BY v.megbizas_napja DESC;
END;
GO

