USE [RPA_Processes_DEV]
GO

/****** Object:  Table [dbo].[csli_001_package_status]    Script Date: 2026. 06. 11. 11:05:09 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[csli_001_package_status](
	[package_status_id] [int] IDENTITY(1,1) NOT NULL,
	[kuldemeny_azonosito_13] [nvarchar](13) NOT NULL,
	[status_id] [int] NOT NULL,
	[changed_at] [datetime2](7) NOT NULL,
	[changed_by] [nvarchar](100) NOT NULL,
	[note] [nvarchar](500) NULL,
 CONSTRAINT [PK_csli_001_package_status] PRIMARY KEY CLUSTERED 
(
	[package_status_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[csli_001_package_status] ADD  CONSTRAINT [DF_pkg_status_changed_at]  DEFAULT (getdate()) FOR [changed_at]
GO

ALTER TABLE [dbo].[csli_001_package_status]  WITH CHECK ADD  CONSTRAINT [FK_csli_001_package_status_lookup] FOREIGN KEY([status_id])
REFERENCES [dbo].[csli_001_status_lookup] ([status_id])
GO

ALTER TABLE [dbo].[csli_001_package_status] CHECK CONSTRAINT [FK_csli_001_package_status_lookup]
GO

