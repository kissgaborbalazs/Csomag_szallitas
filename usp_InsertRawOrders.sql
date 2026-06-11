USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_InsertRawOrders]    Script Date: 2026. 06. 11. 11:11:47 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE OR ALTER   PROCEDURE [dbo].[usp_InsertRawOrders]
    @orders_json    NVARCHAR(MAX),
    @inserted_count INT = 0 OUTPUT,
    @skipped_count  INT = 0 OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SET @inserted_count = 0;
    SET @skipped_count  = 0;

    DECLARE @caller NVARCHAR(100) = SUSER_SNAME();  -- D07

    -- =========================================================
    -- JSON array → temp tábla
    -- =========================================================
    SELECT
        megrendelo_vevokod,
        megrendelo_cegnev,
        TRY_CAST(megbizas_napja          AS DATE)    AS megbizas_napja,
        TRY_CAST(megbizas_ideje          AS TIME(7)) AS megbizas_ideje,
        felrako_nev,
        felrako_irsz,
        felrako_varos,
        felrako_kozterulet_elnevezese,
        felrako_kozterulet_tipus_utca_ter_ut,
        felrako_hazszam,
        felrako_megjegyzes,
        TRY_CAST(felrako_atveteli_datum  AS DATE)    AS felrako_atveteli_datum,
        felrako_idokapu,
        felrako_kontakt_szemely,
        felrako_telefonszam,
        felrako_email,
        kuldemeny_azonosito,
        LEFT(kuldemeny_azonosito, 13)                AS kuldemeny_azonosito_13,
        TRY_CAST(raklap_hossza_cm        AS INT)     AS raklap_hossza_cm,
        TRY_CAST(raklap_szelessege_cm    AS INT)     AS raklap_szelessege_cm,
        TRY_CAST(raklap_magassaga_cm     AS INT)     AS raklap_magassaga_cm,
        ekaer,
        ISNULL(TRY_CAST(cimirat_nyomtatas AS BIT), 0) AS cimirat_nyomtatas,
        adr_jelolo,
        adr_pont,
        lerakas_idokapu
    INTO #orders
    FROM OPENJSON(@orders_json) WITH (
        megrendelo_vevokod                    NVARCHAR(10)  '$.VEVOKOD',
        megrendelo_cegnev                     NVARCHAR(60)  '$.MEGBIZO_NEVE',
        megbizas_napja                        NVARCHAR(20)  '$.MEGBIZAS_NAPJA',
        megbizas_ideje                        NVARCHAR(20)  '$.MEGBIZAS_IDEJE',
        felrako_nev                           NVARCHAR(100) '$.FELRAKO_NEV',
        felrako_irsz                          NVARCHAR(10)  '$.FELRAKO_IRSZ',
        felrako_varos                         NVARCHAR(60)  '$.FELRAKO_VAROS',
        felrako_kozterulet_elnevezese         NVARCHAR(60)  '$.FELRAKO_KOZTERULET_ELNEVEZESE',
        felrako_kozterulet_tipus_utca_ter_ut  NVARCHAR(60)  '$.FELRAKO_KOZTERULET_TIPUS_UTCA_TER_UT',
        felrako_hazszam                       NVARCHAR(60)  '$.FELRAKO_HAZSZAM',
        felrako_megjegyzes                    NVARCHAR(255) '$.FELRAKO_MEGJEGYZES',
        felrako_atveteli_datum                NVARCHAR(20)  '$.MEGRENDELES_NAPJA',
        felrako_idokapu                       NVARCHAR(20)  '$.FELRAKAS_IDOKAPU',
        felrako_kontakt_szemely               NVARCHAR(60)  '$.FELRAKO_KONTAKT_SZEMELY',
        felrako_telefonszam                   NVARCHAR(20)  '$.FELRAKO_TELEFONSZAM',
        felrako_email                         NVARCHAR(60)  '$.FELRAKO_EMAIL',
        kuldemeny_azonosito                   NVARCHAR(255) '$.KULDEMENY_AZONOSITO',
        raklap_hossza_cm                      NVARCHAR(10)  '$.RAKLAP_HOSSZA_CM',
        raklap_szelessege_cm                  NVARCHAR(10)  '$.RAKLAP_SZELESSEGE_CM',
        raklap_magassaga_cm                   NVARCHAR(10)  '$.RAKLAP_MAGASSAGA_CM',
        ekaer                                 NVARCHAR(30)  '$.EKAER',
        cimirat_nyomtatas                     NVARCHAR(5)   '$.CIMIRAT_NYOMTATAS',
        adr_jelolo                            NVARCHAR(30)  '$.ADR_JELOLO',
        adr_pont                              NVARCHAR(30)  '$.ADR_PONT',
        lerakas_idokapu                       NVARCHAR(20)  '$.LERAKAS_IDOKAPU'
    );

    -- =========================================================
    -- Cursor deklaráció
    -- =========================================================
    DECLARE
        @megrendelo_vevokod                  NVARCHAR(10),
        @megrendelo_cegnev                   NVARCHAR(60),
        @megbizas_napja                      DATE,
        @megbizas_ideje                      TIME(7),
        @felrako_nev                         NVARCHAR(100),
        @felrako_irsz                        NVARCHAR(10),
        @felrako_varos                       NVARCHAR(60),
        @felrako_kozterulet_elnevezese       NVARCHAR(60),
        @felrako_kozterulet_tipus_utca_ter_ut NVARCHAR(60),
        @felrako_hazszam                     NVARCHAR(60),
        @felrako_megjegyzes                  NVARCHAR(255),
        @felrako_atveteli_datum              DATE,
        @felrako_idokapu                     NVARCHAR(20),
        @felrako_kontakt_szemely             NVARCHAR(60),
        @felrako_telefonszam                 NVARCHAR(20),
        @felrako_email                       NVARCHAR(60),
        @kuldemeny_azonosito                 NVARCHAR(255),
        @kuldemeny_azon_13                   NVARCHAR(13),
        @raklap_hossza_cm                    INT,
        @raklap_szelessege_cm                INT,
        @raklap_magassaga_cm                 INT,
        @ekaer                               NVARCHAR(30),
        @cimirat_nyomtatas                   BIT,
        @adr_jelolo                          NVARCHAR(30),
        @adr_pont                            NVARCHAR(30),
        @lerakas_idokapu                     NVARCHAR(20);

    DECLARE
        @current_status  NVARCHAR(50),
        @existing_raw_id INT,
        @is_new          BIT,
        @proc_active_id  INT,
        @proc_szallit    BIT,
        @conflict        BIT;

    DECLARE order_cursor CURSOR LOCAL FAST_FORWARD FOR
        SELECT
            megrendelo_vevokod, megrendelo_cegnev,
            megbizas_napja, megbizas_ideje,
            felrako_nev, felrako_irsz, felrako_varos,
            felrako_kozterulet_elnevezese, felrako_kozterulet_tipus_utca_ter_ut,
            felrako_hazszam, felrako_megjegyzes, felrako_atveteli_datum,
            felrako_idokapu, felrako_kontakt_szemely, felrako_telefonszam, felrako_email,
            kuldemeny_azonosito, kuldemeny_azonosito_13,
            raklap_hossza_cm, raklap_szelessege_cm, raklap_magassaga_cm,
            ekaer, cimirat_nyomtatas, adr_jelolo, adr_pont, lerakas_idokapu
        FROM #orders;

    OPEN order_cursor;
    FETCH NEXT FROM order_cursor INTO
        @megrendelo_vevokod, @megrendelo_cegnev,
        @megbizas_napja, @megbizas_ideje,
        @felrako_nev, @felrako_irsz, @felrako_varos,
        @felrako_kozterulet_elnevezese, @felrako_kozterulet_tipus_utca_ter_ut,
        @felrako_hazszam, @felrako_megjegyzes, @felrako_atveteli_datum,
        @felrako_idokapu, @felrako_kontakt_szemely, @felrako_telefonszam, @felrako_email,
        @kuldemeny_azonosito, @kuldemeny_azon_13,
        @raklap_hossza_cm, @raklap_szelessege_cm, @raklap_magassaga_cm,
        @ekaer, @cimirat_nyomtatas, @adr_jelolo, @adr_pont, @lerakas_idokapu;

    -- =========================================================
    -- Tranzakció — minden sor egységes egész
    -- =========================================================
    BEGIN TRANSACTION;
    BEGIN TRY
        WHILE @@FETCH_STATUS = 0
        BEGIN
            -- ------------------------------------------------
            -- 0. lépés: Lock check
            -- ORDER_RECEIVED / EFJ_RECEIVED → frissíthető
            -- minden más státusz → skip
            -- ------------------------------------------------
            SET @current_status = NULL;

            SELECT TOP 1 @current_status = sl.status_code
            FROM dbo.csli_001_package_status ps
            JOIN dbo.csli_001_status_lookup sl ON sl.status_id = ps.status_id
            WHERE ps.kuldemeny_azonosito_13 = @kuldemeny_azon_13
            ORDER BY ps.package_status_id DESC;

            -- NULL = ismeretlen azonosító → engedélyezett (első betöltés)
            IF @current_status IS NOT NULL
               AND @current_status NOT IN ('ORDER_RECEIVED', 'EFJ_RECEIVED')
            BEGIN
                SET @skipped_count += 1;
                FETCH NEXT FROM order_cursor INTO
                    @megrendelo_vevokod, @megrendelo_cegnev,
                    @megbizas_napja, @megbizas_ideje,
                    @felrako_nev, @felrako_irsz, @felrako_varos,
                    @felrako_kozterulet_elnevezese, @felrako_kozterulet_tipus_utca_ter_ut,
                    @felrako_hazszam, @felrako_megjegyzes, @felrako_atveteli_datum,
                    @felrako_idokapu, @felrako_kontakt_szemely, @felrako_telefonszam, @felrako_email,
                    @kuldemeny_azonosito, @kuldemeny_azon_13,
                    @raklap_hossza_cm, @raklap_szelessege_cm, @raklap_magassaga_cm,
                    @ekaer, @cimirat_nyomtatas, @adr_jelolo, @adr_pont, @lerakas_idokapu;
                CONTINUE;
            END;

            -- ------------------------------------------------
            -- 1. lépés: INS+OBS — csli_001_raw_orders
            -- ------------------------------------------------
            SET @existing_raw_id = NULL;
            SET @is_new          = 0;

            SELECT TOP 1 @existing_raw_id = id
            FROM dbo.csli_001_raw_orders
            WHERE kuldemeny_azonosito_13 = @kuldemeny_azon_13
              AND obsolete = 0
            ORDER BY id DESC;

            IF @existing_raw_id IS NULL
                SET @is_new = 1;
            ELSE
                UPDATE dbo.csli_001_raw_orders
                SET    obsolete = 1, modified_at = GETDATE(), modified_by = @caller
                WHERE  id = @existing_raw_id;

            INSERT INTO dbo.csli_001_raw_orders (
                megrendelo_vevokod, megrendelo_cegnev,
                megbizas_napja, megbizas_ideje,
                felrako_nev, felrako_irsz, felrako_varos,
                felrako_kozterulet_elnevezese, felrako_kozterulet_tipus_utca_ter_ut,
                felrako_hazszam, felrako_megjegyzes, felrako_atveteli_datum,
                felrako_idokapu, felrako_kontakt_szemely, felrako_telefonszam, felrako_email,
                kuldemeny_azonosito,
                raklap_hossza_cm, raklap_szelessege_cm, raklap_magassaga_cm,
                ekaer, cimirat_nyomtatas, adr_jelolo, adr_pont, lerakas_idokapu,
                created_by, modified_by, obsolete
            )
            VALUES (
                @megrendelo_vevokod, @megrendelo_cegnev,
                @megbizas_napja, @megbizas_ideje,
                @felrako_nev, @felrako_irsz, @felrako_varos,
                @felrako_kozterulet_elnevezese, @felrako_kozterulet_tipus_utca_ter_ut,
                @felrako_hazszam, @felrako_megjegyzes, @felrako_atveteli_datum,
                @felrako_idokapu, @felrako_kontakt_szemely, @felrako_telefonszam, @felrako_email,
                @kuldemeny_azonosito,
                @raklap_hossza_cm, @raklap_szelessege_cm, @raklap_magassaga_cm,
                @ekaer, @cimirat_nyomtatas, @adr_jelolo, @adr_pont, @lerakas_idokapu,
                @caller, @caller, 0
            );

            -- ------------------------------------------------
            -- 2. lépés: státuszok — csak új azonosítónál (D06)
            -- ORDER_RECEIVED mindig; ha EFJ már megérkezett → EFJ_RECEIVED is
            -- (EFJ betöltéskor státusz nem keletkezik ha nincs még rendelés)
            -- ------------------------------------------------
            IF @is_new = 1
            BEGIN
                INSERT INTO dbo.csli_001_package_status (kuldemeny_azonosito_13, status_id, changed_by)
                SELECT @kuldemeny_azon_13, status_id, @caller
                FROM dbo.csli_001_status_lookup
                WHERE status_code = 'ORDER_RECEIVED';

                IF EXISTS (
                    SELECT 1 FROM dbo.csli_001_raw_efj
                    WHERE kuldemeny_azonosito_13 = @kuldemeny_azon_13 AND obsolete = 0
                )
                    INSERT INTO dbo.csli_001_package_status (kuldemeny_azonosito_13, status_id, changed_by)
                    SELECT @kuldemeny_azon_13, status_id, @caller
                    FROM dbo.csli_001_status_lookup
                    WHERE status_code = 'EFJ_RECEIVED';
            END;

            -- ------------------------------------------------
            -- 3–4. lépés: proc_data
            -- ------------------------------------------------
            SET @proc_active_id = NULL;
            SET @proc_szallit   = NULL;

            SELECT @proc_active_id = id, @proc_szallit = szallitonak_atadhato
            FROM dbo.csli_001_proc_data
            WHERE kuldemeny_azonosito_13 = @kuldemeny_azon_13
              AND obsolete = 0;

            IF @proc_active_id IS NULL
            BEGIN
                -- 3. lépés: első sor létrehozása (D09)
                EXEC dbo.usp_UpsertProcData
                    @kuldemeny_azon_13 = @kuldemeny_azon_13,
                    @current_active_id = NULL,
                    @changed_by        = @caller,
                    @changes_json      = N'{}',
                    @conflict          = @conflict OUTPUT;
            END
            ELSE IF @proc_szallit = 1
            BEGIN
                -- 4. lépés: szallitonak_atadhato visszaállítás (D12)
                EXEC dbo.usp_UpsertProcData
                    @kuldemeny_azon_13 = @kuldemeny_azon_13,
                    @current_active_id = @proc_active_id,
                    @changed_by        = @caller,
                    @changes_json      = N'{"szallitonak_atadhato":0}',
                    @conflict          = @conflict OUTPUT;

                IF @conflict = 1
                    RAISERROR('szallitonak_atadhato visszaállítás sikertelen — konkurens módosítás: %s', 16, 1, @kuldemeny_azon_13);
            END;

            SET @inserted_count += 1;

            FETCH NEXT FROM order_cursor INTO
                @megrendelo_vevokod, @megrendelo_cegnev,
                @megbizas_napja, @megbizas_ideje,
                @felrako_nev, @felrako_irsz, @felrako_varos,
                @felrako_kozterulet_elnevezese, @felrako_kozterulet_tipus_utca_ter_ut,
                @felrako_hazszam, @felrako_megjegyzes, @felrako_atveteli_datum,
                @felrako_idokapu, @felrako_kontakt_szemely, @felrako_telefonszam, @felrako_email,
                @kuldemeny_azonosito, @kuldemeny_azon_13,
                @raklap_hossza_cm, @raklap_szelessege_cm, @raklap_magassaga_cm,
                @ekaer, @cimirat_nyomtatas, @adr_jelolo, @adr_pont, @lerakas_idokapu;
        END;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        CLOSE order_cursor;
        DEALLOCATE order_cursor;
        THROW;
    END CATCH;

    CLOSE order_cursor;
    DEALLOCATE order_cursor;
END;
GO

