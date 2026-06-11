USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_GetPackageHistory]    Script Date: 2026. 06. 11. 11:10:06 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE OR ALTER PROCEDURE [dbo].[usp_GetPackageHistory]
    @kuldemeny_azon_13  NVARCHAR(13)
AS
BEGIN
    SET NOCOUNT ON;
    -- Státuszváltások + mezőmódosítások — egységes nézet időrendben
    SELECT
        ps.changed_at                               AS event_time,
        'STATUS_CHANGE'                             AS event_type,
        sl.status_label_hu                          AS event_label,
        NULL                                        AS old_value,
        sl.status_label_hu                          AS new_value,
        ps.changed_by                               AS actor,
        ps.note
    FROM dbo.csli_001_package_status ps
    JOIN dbo.csli_001_status_lookup sl
        ON sl.status_id = ps.status_id
    WHERE ps.kuldemeny_azonosito_13 = @kuldemeny_azon_13

    UNION ALL

    SELECT
        oc.changed_at                               AS event_time,
        'FIELD_CHANGE'                              AS event_type,
        oc.source_table + ' / ' + oc.field_name    AS event_label,
        oc.old_value,
        oc.new_value,
        oc.changed_by                               AS actor,
        NULL                                        AS note
    FROM dbo.csli_001_operative_changes oc
    WHERE oc.kuldemeny_azonosito_13 = @kuldemeny_azon_13

    ORDER BY event_time DESC;
END;

GO

