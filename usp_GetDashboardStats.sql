USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_GetDashboardStats]    Script Date: 2026. 06. 11. 11:09:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE OR ALTER   PROCEDURE [dbo].[usp_GetDashboardStats]
AS
BEGIN
    SET NOCOUNT ON;
 
    SELECT
        SUM(CASE WHEN status_is_locked = 0
                  AND current_status_code <> 'CANCELLED'
             THEN 1 ELSE 0 END)                             AS nyitott_count,
 
        SUM(CASE WHEN current_status_code = 'ORDER_RECEIVED'
                  AND status_is_locked = 0
             THEN 1 ELSE 0 END)                             AS hianyzo_efj_count,
 
        SUM(CASE WHEN current_status_code = 'EFJ_RECEIVED'
                  AND COALESCE(szallitonak_atadhato, 0) = 0
                  AND status_is_locked = 0
             THEN 1 ELSE 0 END)                             AS ellenorzesre_count
 
    FROM dbo.csli_001_v_package_current;
END;
GO

