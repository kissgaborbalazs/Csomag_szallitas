USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_UpsertOperativeData]    Script Date: 2026. 06. 11. 11:12:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE OR ALTER     PROCEDURE [dbo].[usp_UpsertOperativeData]
    @kuldemeny_azon_13  NVARCHAR(13),
    @current_active_id  INT             = NULL,
    @changed_by         NVARCHAR(100),
    @changes_json       NVARCHAR(MAX),
    @conflict           BIT             OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SET @conflict = 0;

	IF @current_active_id = 0 SET @current_active_id = NULL;

    -- Ütközés ellenőrzés
    IF @current_active_id IS NOT NULL
       AND NOT EXISTS (
           SELECT 1 FROM dbo.csli_001_operative_data_orders
           WHERE id = @current_active_id AND obsolete = 0
       )
    BEGIN
        SET @conflict = 1;
        RETURN;
    END;

    -- JSON mezők kiolvasása
    DECLARE @j_felrako_nev                      NVARCHAR(100);
    DECLARE @j_felrako_irsz                     NVARCHAR(10);
    DECLARE @j_felrako_varos                    NVARCHAR(60);
    DECLARE @j_felrako_kozterulet_elnevezese    NVARCHAR(60);
    DECLARE @j_felrako_kozterulet_tipus         NVARCHAR(60);
    DECLARE @j_felrako_hazszam                  NVARCHAR(60);
    DECLARE @j_felrako_megjegyzes               NVARCHAR(255);
    DECLARE @j_felrako_atveteli_datum           NVARCHAR(20);
    DECLARE @j_felrako_idokapu                  NVARCHAR(20);
    DECLARE @j_felrako_kontakt_szemely          NVARCHAR(60);
    DECLARE @j_felrako_telefonszam              NVARCHAR(20);
    DECLARE @j_felrako_email                    NVARCHAR(60);
    DECLARE @j_raklap_hossza_cm                 NVARCHAR(10);
    DECLARE @j_raklap_szelessege_cm             NVARCHAR(10);
    DECLARE @j_raklap_magassaga_cm              NVARCHAR(10);
    DECLARE @j_ekaer                            NVARCHAR(30);
    DECLARE @j_cimirat_nyomtatas                NVARCHAR(5);
    DECLARE @j_adr_jelolo                       NVARCHAR(30);
    DECLARE @j_adr_pont                         NVARCHAR(30);
    DECLARE @j_lerakas_idokapu                  NVARCHAR(20);

    SELECT
        @j_felrako_nev                   = JSON_VALUE(@changes_json, '$.felrako_nev'),
        @j_felrako_irsz                  = JSON_VALUE(@changes_json, '$.felrako_irsz'),
        @j_felrako_varos                 = JSON_VALUE(@changes_json, '$.felrako_varos'),
        @j_felrako_kozterulet_elnevezese = JSON_VALUE(@changes_json, '$.felrako_kozterulet_elnevezese'),
        @j_felrako_kozterulet_tipus      = JSON_VALUE(@changes_json, '$.felrako_kozterulet_tipus'),
        @j_felrako_hazszam               = JSON_VALUE(@changes_json, '$.felrako_hazszam'),
        @j_felrako_megjegyzes            = JSON_VALUE(@changes_json, '$.felrako_megjegyzes'),
        @j_felrako_atveteli_datum        = JSON_VALUE(@changes_json, '$.felrako_atveteli_datum'),
        @j_felrako_idokapu               = JSON_VALUE(@changes_json, '$.felrako_idokapu'),
        @j_felrako_kontakt_szemely       = JSON_VALUE(@changes_json, '$.felrako_kontakt_szemely'),
        @j_felrako_telefonszam           = JSON_VALUE(@changes_json, '$.felrako_telefonszam'),
        @j_felrako_email                 = JSON_VALUE(@changes_json, '$.felrako_email'),
        @j_raklap_hossza_cm              = JSON_VALUE(@changes_json, '$.raklap_hossza_cm'),
        @j_raklap_szelessege_cm          = JSON_VALUE(@changes_json, '$.raklap_szelessege_cm'),
        @j_raklap_magassaga_cm           = JSON_VALUE(@changes_json, '$.raklap_magassaga_cm'),
        @j_ekaer                         = JSON_VALUE(@changes_json, '$.ekaer'),
        @j_cimirat_nyomtatas             = JSON_VALUE(@changes_json, '$.cimirat_nyomtatas'),
        @j_adr_jelolo                    = JSON_VALUE(@changes_json, '$.adr_jelolo'),
        @j_adr_pont                      = JSON_VALUE(@changes_json, '$.adr_pont'),
        @j_lerakas_idokapu               = JSON_VALUE(@changes_json, '$.lerakas_idokapu');

    -- Audit alapértékek feltöltése (old_value forrás)
    DECLARE @audit_felrako_nev                   NVARCHAR(100);
    DECLARE @audit_felrako_irsz                  NVARCHAR(10);
    DECLARE @audit_felrako_varos                 NVARCHAR(60);
    DECLARE @audit_felrako_kozterulet_elnevezese NVARCHAR(60);
    DECLARE @audit_felrako_kozterulet_tipus      NVARCHAR(60);
    DECLARE @audit_felrako_hazszam               NVARCHAR(60);
    DECLARE @audit_felrako_megjegyzes            NVARCHAR(255);
    DECLARE @audit_felrako_atveteli_datum        DATE;
    DECLARE @audit_felrako_idokapu               NVARCHAR(20);
    DECLARE @audit_felrako_kontakt_szemely       NVARCHAR(60);
    DECLARE @audit_felrako_telefonszam           NVARCHAR(20);
    DECLARE @audit_felrako_email                 NVARCHAR(60);
    DECLARE @audit_raklap_hossza_cm              INT;
    DECLARE @audit_raklap_szelessege_cm          INT;
    DECLARE @audit_raklap_magassaga_cm           INT;
    DECLARE @audit_ekaer                         NVARCHAR(30);
    DECLARE @audit_cimirat_nyomtatas             BIT;
    DECLARE @audit_adr_jelolo                    NVARCHAR(30);
    DECLARE @audit_adr_pont                      NVARCHAR(30);
    DECLARE @audit_lerakas_idokapu               NVARCHAR(20);

    IF @current_active_id IS NOT NULL
        SELECT
            @audit_felrako_nev                   = felrako_nev,
            @audit_felrako_irsz                  = felrako_irsz,
            @audit_felrako_varos                 = felrako_varos,
            @audit_felrako_kozterulet_elnevezese = felrako_kozterulet_elnevezese,
            @audit_felrako_kozterulet_tipus      = felrako_kozterulet_tipus,
            @audit_felrako_hazszam               = felrako_hazszam,
            @audit_felrako_megjegyzes            = felrako_megjegyzes,
            @audit_felrako_atveteli_datum        = felrako_atveteli_datum,
            @audit_felrako_idokapu               = felrako_idokapu,
            @audit_felrako_kontakt_szemely       = felrako_kontakt_szemely,
            @audit_felrako_telefonszam           = felrako_telefonszam,
            @audit_felrako_email                 = felrako_email,
            @audit_raklap_hossza_cm              = raklap_hossza_cm,
            @audit_raklap_szelessege_cm          = raklap_szelessege_cm,
            @audit_raklap_magassaga_cm           = raklap_magassaga_cm,
            @audit_ekaer                         = ekaer,
            @audit_cimirat_nyomtatas             = cimirat_nyomtatas,
            @audit_adr_jelolo                    = adr_jelolo,
            @audit_adr_pont                      = adr_pont,
            @audit_lerakas_idokapu               = lerakas_idokapu
        FROM dbo.csli_001_operative_data_orders
        WHERE id = @current_active_id;
    ELSE
        SELECT
            @audit_felrako_nev                   = felrako_nev,
            @audit_felrako_irsz                  = felrako_irsz,
            @audit_felrako_varos                 = felrako_varos,
            @audit_felrako_kozterulet_elnevezese = felrako_kozterulet_elnevezese,
            @audit_felrako_kozterulet_tipus      = felrako_kozterulet_tipus_utca_ter_ut,
            @audit_felrako_hazszam               = felrako_hazszam,
            @audit_felrako_megjegyzes            = felrako_megjegyzes,
            @audit_felrako_atveteli_datum        = felrako_atveteli_datum,
            @audit_felrako_idokapu               = felrako_idokapu,
            @audit_felrako_kontakt_szemely       = felrako_kontakt_szemely,
            @audit_felrako_telefonszam           = felrako_telefonszam,
            @audit_felrako_email                 = felrako_email,
            @audit_raklap_hossza_cm              = raklap_hossza_cm,
            @audit_raklap_szelessege_cm          = raklap_szelessege_cm,
            @audit_raklap_magassaga_cm           = raklap_magassaga_cm,
            @audit_ekaer                         = ekaer,
            @audit_cimirat_nyomtatas             = cimirat_nyomtatas,
            @audit_adr_jelolo                    = adr_jelolo,
            @audit_adr_pont                      = adr_pont,
            @audit_lerakas_idokapu               = lerakas_idokapu
        FROM dbo.csli_001_raw_orders
        WHERE kuldemeny_azonosito_13 = @kuldemeny_azon_13
          AND obsolete = 0;

    BEGIN TRANSACTION;
    BEGIN TRY
        -- Régi sor elavulttá tétele
        IF @current_active_id IS NOT NULL
            UPDATE dbo.csli_001_operative_data_orders
            SET obsolete = 1
            WHERE id = @current_active_id;

        -- Új sor INSERT
        INSERT INTO dbo.csli_001_operative_data_orders (
            kuldemeny_azonosito_13,
            felrako_nev,
            felrako_irsz,
            felrako_varos,
            felrako_kozterulet_elnevezese,
            felrako_kozterulet_tipus,
            felrako_hazszam,
            felrako_megjegyzes,
            felrako_atveteli_datum,
            felrako_idokapu,
            felrako_kontakt_szemely,
            felrako_telefonszam,
            felrako_email,
            raklap_hossza_cm,
            raklap_szelessege_cm,
            raklap_magassaga_cm,
            ekaer,
            cimirat_nyomtatas,
            adr_jelolo,
            adr_pont,
            lerakas_idokapu,
            created_by,
            obsolete
        )
        SELECT
            @kuldemeny_azon_13,
            CASE WHEN JSON_VALUE(@changes_json, '$.felrako_nev')                   IS NOT NULL OR @changes_json LIKE '%"felrako_nev"%'                   THEN @j_felrako_nev                   ELSE ISNULL(base.felrako_nev, NULL)                   END,
            CASE WHEN JSON_VALUE(@changes_json, '$.felrako_irsz')                  IS NOT NULL OR @changes_json LIKE '%"felrako_irsz"%'                  THEN @j_felrako_irsz                  ELSE ISNULL(base.felrako_irsz, NULL)                  END,
            CASE WHEN JSON_VALUE(@changes_json, '$.felrako_varos')                 IS NOT NULL OR @changes_json LIKE '%"felrako_varos"%'                 THEN @j_felrako_varos                 ELSE ISNULL(base.felrako_varos, NULL)                 END,
            CASE WHEN JSON_VALUE(@changes_json, '$.felrako_kozterulet_elnevezese') IS NOT NULL OR @changes_json LIKE '%"felrako_kozterulet_elnevezese"%' THEN @j_felrako_kozterulet_elnevezese ELSE ISNULL(base.felrako_kozterulet_elnevezese, NULL) END,
            CASE WHEN JSON_VALUE(@changes_json, '$.felrako_kozterulet_tipus')      IS NOT NULL OR @changes_json LIKE '%"felrako_kozterulet_tipus"%'      THEN @j_felrako_kozterulet_tipus      ELSE ISNULL(base.felrako_kozterulet_tipus, NULL)      END,
            CASE WHEN JSON_VALUE(@changes_json, '$.felrako_hazszam')               IS NOT NULL OR @changes_json LIKE '%"felrako_hazszam"%'               THEN @j_felrako_hazszam               ELSE ISNULL(base.felrako_hazszam, NULL)               END,
            CASE WHEN JSON_VALUE(@changes_json, '$.felrako_megjegyzes')            IS NOT NULL OR @changes_json LIKE '%"felrako_megjegyzes"%'            THEN @j_felrako_megjegyzes            ELSE ISNULL(base.felrako_megjegyzes, NULL)            END,
            CASE WHEN JSON_VALUE(@changes_json, '$.felrako_atveteli_datum')        IS NOT NULL OR @changes_json LIKE '%"felrako_atveteli_datum"%'        THEN TRY_CAST(@j_felrako_atveteli_datum AS DATE)         ELSE base.felrako_atveteli_datum         END,
            CASE WHEN JSON_VALUE(@changes_json, '$.felrako_idokapu')               IS NOT NULL OR @changes_json LIKE '%"felrako_idokapu"%'               THEN @j_felrako_idokapu               ELSE ISNULL(base.felrako_idokapu, NULL)               END,
            CASE WHEN JSON_VALUE(@changes_json, '$.felrako_kontakt_szemely')       IS NOT NULL OR @changes_json LIKE '%"felrako_kontakt_szemely"%'       THEN @j_felrako_kontakt_szemely       ELSE ISNULL(base.felrako_kontakt_szemely, NULL)       END,
            CASE WHEN JSON_VALUE(@changes_json, '$.felrako_telefonszam')           IS NOT NULL OR @changes_json LIKE '%"felrako_telefonszam"%'           THEN @j_felrako_telefonszam           ELSE ISNULL(base.felrako_telefonszam, NULL)           END,
            CASE WHEN JSON_VALUE(@changes_json, '$.felrako_email')                 IS NOT NULL OR @changes_json LIKE '%"felrako_email"%'                 THEN @j_felrako_email                 ELSE ISNULL(base.felrako_email, NULL)                 END,
            CASE WHEN JSON_VALUE(@changes_json, '$.raklap_hossza_cm')              IS NOT NULL OR @changes_json LIKE '%"raklap_hossza_cm"%'              THEN TRY_CAST(@j_raklap_hossza_cm AS INT)               ELSE base.raklap_hossza_cm               END,
            CASE WHEN JSON_VALUE(@changes_json, '$.raklap_szelessege_cm')          IS NOT NULL OR @changes_json LIKE '%"raklap_szelessege_cm"%'          THEN TRY_CAST(@j_raklap_szelessege_cm AS INT)           ELSE base.raklap_szelessege_cm           END,
            CASE WHEN JSON_VALUE(@changes_json, '$.raklap_magassaga_cm')           IS NOT NULL OR @changes_json LIKE '%"raklap_magassaga_cm"%'           THEN TRY_CAST(@j_raklap_magassaga_cm AS INT)            ELSE base.raklap_magassaga_cm            END,
            CASE WHEN JSON_VALUE(@changes_json, '$.ekaer')                         IS NOT NULL OR @changes_json LIKE '%"ekaer"%'                         THEN @j_ekaer                         ELSE ISNULL(base.ekaer, NULL)                         END,
            CASE WHEN JSON_VALUE(@changes_json, '$.cimirat_nyomtatas')             IS NOT NULL OR @changes_json LIKE '%"cimirat_nyomtatas"%'             THEN TRY_CAST(@j_cimirat_nyomtatas AS BIT)              ELSE base.cimirat_nyomtatas              END,
            CASE WHEN JSON_VALUE(@changes_json, '$.adr_jelolo')                    IS NOT NULL OR @changes_json LIKE '%"adr_jelolo"%'                    THEN @j_adr_jelolo                    ELSE ISNULL(base.adr_jelolo, NULL)                    END,
            CASE WHEN JSON_VALUE(@changes_json, '$.adr_pont')                      IS NOT NULL OR @changes_json LIKE '%"adr_pont"%'                      THEN @j_adr_pont                      ELSE ISNULL(base.adr_pont, NULL)                      END,
            CASE WHEN JSON_VALUE(@changes_json, '$.lerakas_idokapu')               IS NOT NULL OR @changes_json LIKE '%"lerakas_idokapu"%'               THEN @j_lerakas_idokapu               ELSE ISNULL(base.lerakas_idokapu, NULL)               END,
            @changed_by,
            0
        FROM (
            SELECT
                felrako_nev, felrako_irsz, felrako_varos,
                felrako_kozterulet_elnevezese, felrako_kozterulet_tipus,
                felrako_hazszam, felrako_megjegyzes, felrako_atveteli_datum,
                felrako_idokapu, felrako_kontakt_szemely, felrako_telefonszam,
                felrako_email, raklap_hossza_cm, raklap_szelessege_cm, raklap_magassaga_cm,
                ekaer, cimirat_nyomtatas, adr_jelolo, adr_pont, lerakas_idokapu
            FROM dbo.csli_001_operative_data_orders
            WHERE id = @current_active_id
            UNION ALL
            SELECT NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL
            WHERE @current_active_id IS NULL
        ) base;

        -- Audit loggolás
        INSERT INTO dbo.csli_001_operative_changes
            (kuldemeny_azonosito_13, source_table, field_name, old_value, new_value, changed_by)
        SELECT
            @kuldemeny_azon_13,
            'csli_001_operative_data_orders',
            j.[key],
            CASE j.[key]
                WHEN 'felrako_nev'                   THEN CAST(@audit_felrako_nev AS NVARCHAR(MAX))
                WHEN 'felrako_irsz'                  THEN CAST(@audit_felrako_irsz AS NVARCHAR(MAX))
                WHEN 'felrako_varos'                 THEN CAST(@audit_felrako_varos AS NVARCHAR(MAX))
                WHEN 'felrako_kozterulet_elnevezese' THEN CAST(@audit_felrako_kozterulet_elnevezese AS NVARCHAR(MAX))
                WHEN 'felrako_kozterulet_tipus'      THEN CAST(@audit_felrako_kozterulet_tipus AS NVARCHAR(MAX))
                WHEN 'felrako_hazszam'               THEN CAST(@audit_felrako_hazszam AS NVARCHAR(MAX))
                WHEN 'felrako_megjegyzes'            THEN CAST(@audit_felrako_megjegyzes AS NVARCHAR(MAX))
                WHEN 'felrako_atveteli_datum'        THEN CAST(@audit_felrako_atveteli_datum AS NVARCHAR(MAX))
                WHEN 'felrako_idokapu'               THEN CAST(@audit_felrako_idokapu AS NVARCHAR(MAX))
                WHEN 'felrako_kontakt_szemely'       THEN CAST(@audit_felrako_kontakt_szemely AS NVARCHAR(MAX))
                WHEN 'felrako_telefonszam'           THEN CAST(@audit_felrako_telefonszam AS NVARCHAR(MAX))
                WHEN 'felrako_email'                 THEN CAST(@audit_felrako_email AS NVARCHAR(MAX))
                WHEN 'raklap_hossza_cm'              THEN CAST(@audit_raklap_hossza_cm AS NVARCHAR(MAX))
                WHEN 'raklap_szelessege_cm'          THEN CAST(@audit_raklap_szelessege_cm AS NVARCHAR(MAX))
                WHEN 'raklap_magassaga_cm'           THEN CAST(@audit_raklap_magassaga_cm AS NVARCHAR(MAX))
                WHEN 'ekaer'                         THEN CAST(@audit_ekaer AS NVARCHAR(MAX))
                WHEN 'cimirat_nyomtatas'             THEN CAST(@audit_cimirat_nyomtatas AS NVARCHAR(MAX))
                WHEN 'adr_jelolo'                    THEN CAST(@audit_adr_jelolo AS NVARCHAR(MAX))
                WHEN 'adr_pont'                      THEN CAST(@audit_adr_pont AS NVARCHAR(MAX))
                WHEN 'lerakas_idokapu'               THEN CAST(@audit_lerakas_idokapu AS NVARCHAR(MAX))
                ELSE NULL
            END,
            j.value,
            @changed_by
        FROM OPENJSON(@changes_json) j;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH;
END;
GO

