USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_GetPackageList]    Script Date: 2026. 06. 11. 11:10:23 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER   PROCEDURE [dbo].[usp_GetPackageList]
    @vevokod_filter     NVARCHAR(10)  = NULL,
    @kuldemeny_filter   NVARCHAR(26)  = NULL,
    @date_from          DATE          = NULL,
    @status_code        NVARCHAR(50)  = NULL,
    @page_number        INT           = 1,
    @page_size          INT           = 500
AS
BEGIN
    SET NOCOUNT ON;
    IF @vevokod_filter   = '' SET @vevokod_filter   = NULL;
    IF @kuldemeny_filter = '' SET @kuldemeny_filter = NULL;
    IF @status_code      = '' SET @status_code      = NULL;

    SELECT
        raw_order_id,
        kuldemeny_azonosito,
        kuldemeny_azonosito_13,
        megrendelo_vevokod,
        megrendelo_cegnev,
        CASE
            WHEN megbizas_napja IS NULL THEN ''
            WHEN megbizas_ideje IS NULL THEN CONVERT(NVARCHAR(10), megbizas_napja, 120)
            ELSE CONVERT(NVARCHAR(10), megbizas_napja, 120) + ' ' + CONVERT(NVARCHAR(5), megbizas_ideje, 108)
        END                                         AS megbizas_datetime,
        felrako_kontakt_szemely,
        felrako_telefonszam,
        felrako_email,
        current_status_code,
        current_status_label,
        status_is_locked,
        operative_data_active_id,
        proc_data_active_id,
        raw_efj_id,
        COALESCE(szallitonak_atadhato, 0)           AS szallitonak_atadhato
    FROM dbo.csli_001_v_package_current
    WHERE
        (@vevokod_filter   IS NULL OR megrendelo_vevokod    =      @vevokod_filter)
        AND (@kuldemeny_filter IS NULL OR kuldemeny_azonosito LIKE '%' + @kuldemeny_filter + '%')
        AND (@date_from        IS NULL OR megbizas_napja    >= @date_from)
        AND (@status_code      IS NULL OR current_status_code =    @status_code)
    ORDER BY megbizas_napja DESC, megbizas_ideje DESC, raw_order_id DESC
    OFFSET (@page_number - 1) * @page_size ROWS
    FETCH NEXT @page_size ROWS ONLY;

END;
GO

