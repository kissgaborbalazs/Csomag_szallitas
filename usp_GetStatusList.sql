USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_GetStatusList]    Script Date: 2026. 06. 11. 11:11:20 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE OR ALTER PROCEDURE [dbo].[usp_GetStatusList]
AS
BEGIN
    SET NOCOUNT ON;
    SELECT
        status_id,
        status_code,
        status_label_hu,
        sort_order,
        is_lock_trigger
    FROM dbo.csli_001_status_lookup
    WHERE is_active = 1
    ORDER BY sort_order;
END;

GO

