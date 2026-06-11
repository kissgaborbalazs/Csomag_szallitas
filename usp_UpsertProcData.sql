USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_UpsertProcData]    Script Date: 2026. 06. 11. 11:12:33 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER PROCEDURE [dbo].[usp_UpsertProcData]
    @kuldemeny_azon_13  NVARCHAR(13),
    @current_active_id  INT             = NULL,
    @changed_by         NVARCHAR(100),
    @changes_json       NVARCHAR(MAX),
    @conflict           BIT             OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SET @conflict = 0;

    -- Ütközés ellenőrzés
    IF @current_active_id IS NOT NULL
       AND NOT EXISTS (
           SELECT 1 FROM dbo.csli_001_proc_data
           WHERE id = @current_active_id AND obsolete = 0
       )
    BEGIN
        SET @conflict = 1;
        RETURN;
    END;

    BEGIN TRANSACTION;
    BEGIN TRY
        -- Régi sor elavulttá tétele
        IF @current_active_id IS NOT NULL
            UPDATE dbo.csli_001_proc_data
            SET obsolete = 1
            WHERE id = @current_active_id;

        -- Új sor INSERT JSON delta alapján
        INSERT INTO dbo.csli_001_proc_data (
            kuldemeny_azonosito_13,
            szallitonak_atadhato,
            tervezett_felvetel_datum,
            sikertelen_felvetel_megj,
            ekaer_kuldheto,
            szamlazhato,
            indokolatlan_kiallas,
            created_by,
            obsolete
        )
        SELECT
            @kuldemeny_azon_13,
            CASE WHEN @changes_json LIKE '%"szallitonak_atadhato"%'
                THEN TRY_CAST(JSON_VALUE(@changes_json, '$.szallitonak_atadhato') AS BIT)
                ELSE base.szallitonak_atadhato END,
            CASE WHEN @changes_json LIKE '%"tervezett_felvetel_datum"%'
                THEN TRY_CAST(JSON_VALUE(@changes_json, '$.tervezett_felvetel_datum') AS DATE)
                ELSE base.tervezett_felvetel_datum END,
            CASE WHEN @changes_json LIKE '%"sikertelen_felvetel_megj"%'
                THEN JSON_VALUE(@changes_json, '$.sikertelen_felvetel_megj')
                ELSE base.sikertelen_felvetel_megj END,
            CASE WHEN @changes_json LIKE '%"ekaer_kuldheto"%'
                THEN TRY_CAST(JSON_VALUE(@changes_json, '$.ekaer_kuldheto') AS BIT)
                ELSE base.ekaer_kuldheto END,
            CASE WHEN @changes_json LIKE '%"szamlazhato"%'
                THEN TRY_CAST(JSON_VALUE(@changes_json, '$.szamlazhato') AS BIT)
                ELSE base.szamlazhato END,
            CASE WHEN @changes_json LIKE '%"indokolatlan_kiallas"%'
                THEN TRY_CAST(JSON_VALUE(@changes_json, '$.indokolatlan_kiallas') AS BIT)
                ELSE base.indokolatlan_kiallas END,
            @changed_by,
            0
        FROM (
            SELECT
                szallitonak_atadhato,
                tervezett_felvetel_datum, sikertelen_felvetel_megj,
                ekaer_kuldheto, szamlazhato, indokolatlan_kiallas
            FROM dbo.csli_001_proc_data
            WHERE id = @current_active_id
            UNION ALL
            SELECT NULL, NULL, NULL, NULL, NULL, NULL
            WHERE @current_active_id IS NULL
        ) base;

        -- Audit loggolás
        INSERT INTO dbo.csli_001_operative_changes
            (kuldemeny_azonosito_13, source_table, field_name, old_value, new_value, changed_by)
        SELECT
            @kuldemeny_azon_13,
            'csli_001_proc_data',
            j.[key],
            CASE j.[key]
                WHEN 'szallitonak_atadhato'    THEN CAST(base_audit.szallitonak_atadhato AS NVARCHAR(MAX))
                WHEN 'tervezett_felvetel_datum' THEN CAST(base_audit.tervezett_felvetel_datum AS NVARCHAR(MAX))
                WHEN 'sikertelen_felvetel_megj' THEN base_audit.sikertelen_felvetel_megj
                WHEN 'ekaer_kuldheto'           THEN CAST(base_audit.ekaer_kuldheto AS NVARCHAR(MAX))
                WHEN 'szamlazhato'              THEN CAST(base_audit.szamlazhato AS NVARCHAR(MAX))
                WHEN 'indokolatlan_kiallas'     THEN CAST(base_audit.indokolatlan_kiallas AS NVARCHAR(MAX))
                ELSE NULL
            END,
            j.value,
            @changed_by
        FROM OPENJSON(@changes_json) j
        CROSS JOIN (
            SELECT
                szallitonak_atadhato,
                tervezett_felvetel_datum, sikertelen_felvetel_megj,
                ekaer_kuldheto, szamlazhato, indokolatlan_kiallas
            FROM dbo.csli_001_proc_data
            WHERE id = @current_active_id
            UNION ALL
            SELECT NULL, NULL, NULL, NULL, NULL, NULL
            WHERE @current_active_id IS NULL
        ) base_audit;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH;
END;
GO

