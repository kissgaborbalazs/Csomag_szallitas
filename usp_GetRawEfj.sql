USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_GetRawEfj]    Script Date: 2026. 06. 11. 11:10:39 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE OR ALTER PROCEDURE [dbo].[usp_GetRawEfj]
    @kuldemeny_azonosito_13  NVARCHAR(13)
AS
BEGIN
    SET NOCOUNT ON;

    SELECT TOP 1
        -- Audit
        e.id,
        e.efj_id,
        e.efeladojegyzek_azon,
        e.betoltes_datum,
        e.efj_zaras,
        e.efj_szoftver,
        e.varhato_feladas_datum,
        e.created_at,
        e.created_by,
        e.modified_at,
        e.modified_by,
        e.obsolete,

        -- Feladó
        e.felado_vevokod,
        e.felado_megallapodas,
        e.felado_nev,
        e.felado_irsz,
        e.felado_hely,
        e.felado_kozelebbi_cim,
        e.felado_kozterulet_nev,
        e.felado_kozterulet_jelleg,
        e.felado_hazszam,
        e.felado_megjegyzes,
        e.felado_email,
        e.felado_telefon,

        -- Küldemény
        e.sorszam,
        e.alapszolg,
        e.szolgaltatas,
        e.kuldemeny_brutto_tomege_kg,
        e.kezbesites_modja,
        e.utanvet_osszege_ft,
        e.kuldemeny_aruertek,
        e.lerako_referencia_szam,
        e.vezer_azon,
        e.ugyfadat1,
        e.ugyfadat2,

        -- Címzett
        e.cimzett_nev,
        e.cimzett_irsz,
        e.cimzett_varos,
        e.cimzett_kozerlebbi_cim,
        e.cimzett_kozterulet_nev,
        e.cimzett_kozterulet_jelleg,
        e.cimzett_hazszam,
        e.cimzett_megjegyzes,
        e.cimzett_kontakt_szemely,
        e.cimzett_email,
        e.cimzett_telefonszam

    FROM dbo.csli_001_raw_efj AS e
    WHERE e.kuldemeny_azonosito_13 = @kuldemeny_azonosito_13
      AND e.obsolete = 0
    ORDER BY e.id DESC;   -- ha több nem-obsolete sor lenne, a legfrissebbet adja

END;

GO

