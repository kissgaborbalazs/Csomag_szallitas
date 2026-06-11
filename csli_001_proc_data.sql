USE [RPA_Processes_DEV]
GO

/****** Object:  Table [dbo].[csli_001_proc_data]    Script Date: 2026. 06. 11. 11:05:27 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[csli_001_proc_data](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[kuldemeny_azonosito_13] [nvarchar](13) NOT NULL,
	[tervezett_felvetel_datum] [date] NULL,
	[szallitonak_atadhato] [bit] NULL,
	[sikertelen_felvetel_megj] [nvarchar](255) NULL,
	[ekaer_kuldheto] [bit] NULL,
	[szamlazhato] [bit] NULL,
	[indokolatlan_kiallas] [bit] NULL,
	[created_at] [datetime2](3) NOT NULL,
	[created_by] [nvarchar](100) NOT NULL,
	[obsolete] [bit] NOT NULL,
 CONSTRAINT [PK_csli_001_proc_data] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[csli_001_proc_data] ADD  CONSTRAINT [DF_proc_data_created_at]  DEFAULT (getdate()) FOR [created_at]
GO

ALTER TABLE [dbo].[csli_001_proc_data] ADD  CONSTRAINT [DF_proc_data_obsolete]  DEFAULT ((0)) FOR [obsolete]
GO

