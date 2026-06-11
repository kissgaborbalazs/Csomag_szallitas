USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_GetPackageDetail]    Script Date: 2026. 06. 11. 11:09:49 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE OR ALTER PROCEDURE [dbo].[usp_GetPackageDetail]
    @kuldemeny_azon_13  NVARCHAR(13)
AS
BEGIN
    SET NOCOUNT ON;
    SELECT *
    FROM dbo.csli_001_v_package_current
    WHERE kuldemeny_azonosito_13 = @kuldemeny_azon_13;
END;

GO

