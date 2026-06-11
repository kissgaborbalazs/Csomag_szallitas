USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_InsertRawEFJ]    Script Date: 2026. 06. 11. 11:11:34 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE OR ALTER   PROCEDURE [dbo].[usp_InsertRawEFJ]
    @json           NVARCHAR(MAX),
    @updated_rows   INT = 0 OUTPUT,    -- meglévő EFJ sorok frissítve (INS+OBS)
    @new_rows       INT = 0 OUTPUT     -- újonnan érkezett EFJ sorok
AS
BEGIN
    SET NOCOUNT ON;
    SET @updated_rows = 0;
    SET @new_rows     = 0;

    DECLARE @service_user NVARCHAR(100) = SUSER_SNAME();

    -- =========================================================
    -- JSON array → temp tábla
    -- =========================================================
    SELECT
        LEFT(kuldemeny_azonosito, 13)           AS k13,
        efj_id, efeladojegyzek_azon, betoltes_datum, efj_zaras, efj_szoftver,
        felado_vevokod, felado_megallapodas, felado_nev, felado_irsz, felado_hely,
        felado_kozelebbi_cim, felado_kozterulet_nev, felado_kozterulet_jelleg,
        felado_hazszam, felado_megjegyzes, felado_email, felado_telefon,
        varhato_feladas_datum, sorszam, kuldemeny_azonosito, vezer_azon, alapszolg,
        kuldemeny_brutto_tomege_kg, kezbesites_modja, szolgaltatas,
        utanvet_osszege_ft, kuldemeny_aruertek, lerako_referencia_szam,
        cimzett_nev, cimzett_irsz, cimzett_varos, cimzett_kozerlebbi_cim,
        cimzett_kozterulet_nev, cimzett_kozterulet_jelleg, cimzett_hazszam,
        cimzett_megjegyzes, cimzett_kontakt_szemely, cimzett_email, cimzett_telefonszam,
        ugyfadat1, ugyfadat2
    INTO #efj
    FROM OPENJSON(@json) WITH (
        efj_id                      NVARCHAR(255)   '$.EFJ_ID',
        efeladojegyzek_azon         NVARCHAR(255)   '$.EFELADOJEGYZEK_AZON',
        betoltes_datum              DATETIME2(0)    '$.BETOLTES_DATUM',
        efj_zaras                   NVARCHAR(255)   '$.EFJ_ZARAS',
        efj_szoftver                NVARCHAR(255)   '$.EFJ_SZOFTVER',
        felado_vevokod              NVARCHAR(255)   '$.FELADO_VEVOKOD',
        felado_megallapodas         NVARCHAR(255)   '$.FELADO_MEGALLAPODAS',
        felado_nev                  NVARCHAR(255)   '$.FELADO_NEV',
        felado_irsz                 NVARCHAR(4)     '$.FELADO_IRSZ',
        felado_hely                 NVARCHAR(255)   '$.FELADO_HELY',
        felado_kozelebbi_cim        NVARCHAR(255)   '$.FELADO_KOZELEBBI_CIM',
        felado_kozterulet_nev       NVARCHAR(255)   '$.FELADO_KOZTERULET_NEV',
        felado_kozterulet_jelleg    NVARCHAR(255)   '$.FELADO_KOZTERULET_JELLEG',
        felado_hazszam              NVARCHAR(255)   '$.FELADO_HAZSZAM',
        felado_megjegyzes           NVARCHAR(255)   '$.FELADO_MEGJEGYZES',
        felado_email                NVARCHAR(255)   '$.FELADO_EMAIL',
        felado_telefon              NVARCHAR(255)   '$.FELADO_TELEFON',
        varhato_feladas_datum       NVARCHAR(255)   '$.VARHATO_FELADAS_DATUM',
        sorszam                     NVARCHAR(255)   '$.SORSZAM',
        kuldemeny_azonosito         NVARCHAR(26)    '$.KULDEMENY_AZONOSITO',
        vezer_azon                  NVARCHAR(26)    '$.VEZER_AZON',
        alapszolg                   NVARCHAR(255)   '$.ALAPSZOLG',
        kuldemeny_brutto_tomege_kg  INT             '$.KULDEMENY_BRUTTO_TOMEGE_KG',
        kezbesites_modja            NVARCHAR(10)    '$.KEZBESITES_MODJA',
        szolgaltatas                NVARCHAR(255)   '$.SZOLGALTATAS',
        utanvet_osszege_ft          DECIMAL(10,2)   '$.UTANVET_OSSZEGE_FT',
        kuldemeny_aruertek          DECIMAL(18,2)   '$.KULDEMENY_ARUERTEK',
        lerako_referencia_szam      NVARCHAR(20)    '$.LERAKO_REFERENCIA_SZAM',
        cimzett_nev                 NVARCHAR(200)   '$.LERAKO_NEV',
        cimzett_irsz                NVARCHAR(10)    '$.LERAKO_IRSZ',
        cimzett_varos               NVARCHAR(60)    '$.LERAKO_VAROS',
        cimzett_kozerlebbi_cim      NVARCHAR(255)   '$.cimzett_kozerlebbi_cim',
        cimzett_kozterulet_nev      NVARCHAR(60)    '$.LERAKO_KOZTERULET_ELNEVEZESE',
        cimzett_kozterulet_jelleg   NVARCHAR(60)    '$.LERAKO_KOZTERULET_MEGNEVEZESE_UTCA_TER_UT',
        cimzett_hazszam             NVARCHAR(60)    '$.LERAKO_HAZSZAM',
        cimzett_megjegyzes          NVARCHAR(255)   '$.LERAKO_MEGJEGYZES',
        cimzett_kontakt_szemely     NVARCHAR(60)    '$.LERAKO_KONTAKT_SZEMELY',
        cimzett_email               NVARCHAR(60)    '$.LERAKO_EMAIL',
        cimzett_telefonszam         NVARCHAR(20)    '$.LERAKO_TELEFONSZAM',
        ugyfadat1                   NVARCHAR(255)   '$.UGYFADAT1',
        ugyfadat2                   NVARCHAR(255)   '$.UGYFADAT2'
    );

    -- =========================================================
    -- Lock check: engedélyezett azonosítók jelölése
    -- NULL státusz (nincs még rendelés), ORDER_RECEIVED, EFJ_RECEIVED → OK
    -- Minden más → skip
    -- =========================================================
    ALTER TABLE #efj ADD allowed BIT NOT NULL DEFAULT 0;

    UPDATE e
    SET allowed = 1
    FROM #efj e
    WHERE ISNULL((
        SELECT TOP 1 sl.status_code
        FROM dbo.csli_001_package_status ps
        JOIN dbo.csli_001_status_lookup sl ON sl.status_id = ps.status_id
        WHERE ps.kuldemeny_azonosito_13 = e.k13
        ORDER BY ps.package_status_id DESC
    ), 'ORDER_RECEIVED') IN ('ORDER_RECEIVED', 'EFJ_RECEIVED');

    -- =========================================================
    -- Létező aktív sorok azonosítása (INS+OBS-hoz kell)
    -- =========================================================
    ALTER TABLE #efj ADD existing_id INT NULL;

    UPDATE e
    SET existing_id = efj.id
    FROM #efj e
    JOIN dbo.csli_001_raw_efj efj
        ON efj.kuldemeny_azonosito_13 = e.k13 AND efj.obsolete = 0
    WHERE e.allowed = 1;

    BEGIN TRANSACTION;
    BEGIN TRY

        -- =========================================================
        -- 1. lépés: INS+OBS — meglévő aktív sorok elavulttá tétele
        -- =========================================================
        UPDATE efj
        SET obsolete    = 1,
            modified_at = GETDATE(),
            modified_by = @service_user
        FROM dbo.csli_001_raw_efj efj
        JOIN #efj e ON e.existing_id = efj.id
        WHERE e.allowed = 1;

        SET @updated_rows = @@ROWCOUNT;

        -- =========================================================
        -- 2. lépés: Új sorok INSERT (minden engedélyezett tétel)
        -- =========================================================
        INSERT INTO dbo.csli_001_raw_efj (
            efj_id, efeladojegyzek_azon, betoltes_datum, efj_zaras, efj_szoftver,
            felado_vevokod, felado_megallapodas, felado_nev, felado_irsz, felado_hely,
            felado_kozelebbi_cim, felado_kozterulet_nev, felado_kozterulet_jelleg,
            felado_hazszam, felado_megjegyzes, felado_email, felado_telefon,
            varhato_feladas_datum, sorszam, kuldemeny_azonosito, vezer_azon, alapszolg,
            kuldemeny_brutto_tomege_kg, kezbesites_modja, szolgaltatas,
            utanvet_osszege_ft, kuldemeny_aruertek, lerako_referencia_szam,
            cimzett_nev, cimzett_irsz, cimzett_varos, cimzett_kozerlebbi_cim,
            cimzett_kozterulet_nev, cimzett_kozterulet_jelleg, cimzett_hazszam,
            cimzett_megjegyzes, cimzett_kontakt_szemely, cimzett_email, cimzett_telefonszam,
            ugyfadat1, ugyfadat2,
            created_by, modified_by, obsolete
        )
        SELECT
            efj_id, efeladojegyzek_azon, betoltes_datum, efj_zaras, efj_szoftver,
            felado_vevokod, felado_megallapodas, felado_nev, felado_irsz, felado_hely,
            felado_kozelebbi_cim, felado_kozterulet_nev, felado_kozterulet_jelleg,
            felado_hazszam, felado_megjegyzes, felado_email, felado_telefon,
            varhato_feladas_datum, sorszam, kuldemeny_azonosito, vezer_azon, alapszolg,
            kuldemeny_brutto_tomege_kg, kezbesites_modja, szolgaltatas,
            utanvet_osszege_ft, kuldemeny_aruertek, lerako_referencia_szam,
            cimzett_nev, cimzett_irsz, cimzett_varos, cimzett_kozerlebbi_cim,
            cimzett_kozterulet_nev, cimzett_kozterulet_jelleg, cimzett_hazszam,
            cimzett_megjegyzes, cimzett_kontakt_szemely, cimzett_email, cimzett_telefonszam,
            ugyfadat1, ugyfadat2,
            @service_user, @service_user, 0
        FROM #efj
        WHERE allowed = 1;

        SET @new_rows = @@ROWCOUNT - @updated_rows;

        -- =========================================================
        -- 3. lépés: EFJ_RECEIVED státusz INSERT
        -- Csak ha van aktív raw_orders ÉS jelenlegi státusz = ORDER_RECEIVED
        -- Ha nincs raw_orders: InsertRawOrders fogja kezelni betöltéskor
        -- =========================================================
        INSERT INTO dbo.csli_001_package_status (kuldemeny_azonosito_13, status_id, changed_by)
        SELECT DISTINCT e.k13, sl.status_id, @service_user
        FROM #efj e
        JOIN dbo.csli_001_raw_orders ro
            ON ro.kuldemeny_azonosito_13 = e.k13 AND ro.obsolete = 0
        JOIN dbo.csli_001_status_lookup sl
            ON sl.status_code = 'EFJ_RECEIVED'
        WHERE e.allowed = 1
          AND (
              SELECT TOP 1 sl2.status_code
              FROM dbo.csli_001_package_status ps2
              JOIN dbo.csli_001_status_lookup sl2 ON sl2.status_id = ps2.status_id
              WHERE ps2.kuldemeny_azonosito_13 = e.k13
              ORDER BY ps2.package_status_id DESC
          ) = 'ORDER_RECEIVED';

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH;
END;
GO

