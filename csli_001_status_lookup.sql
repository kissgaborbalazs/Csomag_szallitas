USE [RPA_Processes_DEV]
GO

/****** Object:  Table [dbo].[csli_001_status_lookup]    Script Date: 2026. 06. 11. 11:07:11 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[csli_001_status_lookup](
	[status_id] [int] IDENTITY(1,1) NOT NULL,
	[status_code] [nvarchar](50) NOT NULL,
	[status_label_hu] [nvarchar](200) NOT NULL,
	[sort_order] [int] NOT NULL,
	[is_active] [bit] NOT NULL,
	[is_lock_trigger] [bit] NOT NULL,
 CONSTRAINT [PK_csli_001_status_lookup] PRIMARY KEY CLUSTERED 
(
	[status_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [UK_csli_001_status_lookup_code] UNIQUE NONCLUSTERED 
(
	[status_code] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[csli_001_status_lookup] ADD  CONSTRAINT [DF_status_lookup_sort]  DEFAULT ((0)) FOR [sort_order]
GO

ALTER TABLE [dbo].[csli_001_status_lookup] ADD  CONSTRAINT [DF_status_lookup_active]  DEFAULT ((1)) FOR [is_active]
GO

ALTER TABLE [dbo].[csli_001_status_lookup] ADD  CONSTRAINT [DF_status_lookup_lock]  DEFAULT ((0)) FOR [is_lock_trigger]
GO

