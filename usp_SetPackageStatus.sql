USE [RPA_Processes_DEV]
GO

/****** Object:  StoredProcedure [dbo].[usp_SetPackageStatus]    Script Date: 2026. 06. 11. 11:12:02 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE OR ALTER   PROCEDURE [dbo].[usp_SetPackageStatus]
    @kuldemeny_azon_13  NVARCHAR(13),
    @status_code        NVARCHAR(50),
    @changed_by         NVARCHAR(100)	= NULL,
    @note               NVARCHAR(500)   = NULL,
    @new_status_id      INT             OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

	If @changed_by IS NULL
		SET @changed_by = SUSER_NAME()

    DECLARE @status_id INT;
    SELECT @status_id = status_id
    FROM dbo.csli_001_status_lookup
    WHERE status_code = @status_code
      AND is_active = 1;

    IF @status_id IS NULL
        RAISERROR(N'Érvénytelen vagy inaktív státuszkód: %s', 16, 1, @status_code);

    BEGIN TRANSACTION;
    BEGIN TRY
        INSERT INTO dbo.csli_001_package_status
            (kuldemeny_azonosito_13, status_id, changed_by, note)
        VALUES
            (@kuldemeny_azon_13, @status_id, @changed_by, @note);

        SET @new_status_id = SCOPE_IDENTITY();
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH;
END;

GO

