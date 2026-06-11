USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_ExportPackageList]    Script Date: 2026. 06. 11. 11:08:40 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE OR ALTER PROCEDURE [dbo].[usp_ExportPackageList]
    @status_code        NVARCHAR(50)    = NULL,
    @kuldemeny_filter   NVARCHAR(26)    = NULL,
    @date_from          DATE            = NULL,
    @date_to            DATE            = NULL
AS
BEGIN
    SET NOCOUNT ON;
    -- Teljes oszlopkészlet exporthoz
    SELECT *
    FROM dbo.csli_001_v_package_current
    WHERE
        (@status_code       IS NULL OR current_status_code = @status_code)
        AND (@kuldemeny_filter  IS NULL OR kuldemeny_azonosito LIKE '%' + @kuldemeny_filter + '%')
        AND (@date_from         IS NULL OR felrako_atveteli_datum >= @date_from)
        AND (@date_to           IS NULL OR felrako_atveteli_datum <= @date_to)
    ORDER BY megbizas_napja DESC, raw_order_id DESC;
END;

GO

