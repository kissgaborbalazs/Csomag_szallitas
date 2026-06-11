USE [RPA_Processes_DEV]
GO

/****** Object:  Table [dbo].[csli_001_raw_orders]    Script Date: 2026. 06. 11. 11:06:49 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[csli_001_raw_orders](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[megrendelo_vevokod] [nvarchar](10) NOT NULL,
	[megrendelo_cegnev] [nvarchar](60) NOT NULL,
	[megbizas_napja] [date] NULL,
	[megbizas_ideje] [time](7) NULL,
	[felrako_nev] [nvarchar](100) NOT NULL,
	[felrako_irsz] [nvarchar](10) NOT NULL,
	[felrako_varos] [nvarchar](60) NOT NULL,
	[felrako_kozterulet_elnevezese] [nvarchar](60) NOT NULL,
	[felrako_kozterulet_tipus_utca_ter_ut] [nvarchar](60) NOT NULL,
	[felrako_hazszam] [nvarchar](60) NOT NULL,
	[felrako_megjegyzes] [nvarchar](255) NULL,
	[felrako_atveteli_datum] [date] NOT NULL,
	[felrako_idokapu] [nvarchar](20) NULL,
	[felrako_kontakt_szemely] [nvarchar](60) NOT NULL,
	[felrako_telefonszam] [nvarchar](20) NOT NULL,
	[felrako_email] [nvarchar](60) NULL,
	[kuldemeny_azonosito] [nvarchar](255) NULL,
	[kuldemeny_azonosito_13]  AS (left([kuldemeny_azonosito],(13))) PERSISTED,
	[raklap_hossza_cm] [int] NOT NULL,
	[raklap_szelessege_cm] [int] NOT NULL,
	[raklap_magassaga_cm] [int] NOT NULL,
	[ekaer] [nvarchar](30) NULL,
	[cimirat_nyomtatas] [bit] NOT NULL,
	[adr_jelolo] [nvarchar](30) NULL,
	[adr_pont] [nvarchar](30) NULL,
	[lerakas_idokapu] [nvarchar](20) NULL,
	[created_at] [datetime2](3) NOT NULL,
	[created_by] [nvarchar](100) NOT NULL,
	[modified_at] [datetime2](3) NOT NULL,
	[modified_by] [nvarchar](100) NOT NULL,
	[obsolete] [bit] NOT NULL,
 CONSTRAINT [PK_csli_001_raw_orders] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[csli_001_raw_orders] ADD  CONSTRAINT [DF_raw_orders_created_at]  DEFAULT (getdate()) FOR [created_at]
GO

ALTER TABLE [dbo].[csli_001_raw_orders] ADD  CONSTRAINT [DF_raw_orders_modified_at]  DEFAULT (getdate()) FOR [modified_at]
GO

ALTER TABLE [dbo].[csli_001_raw_orders] ADD  CONSTRAINT [DF_raw_orders_obsolete]  DEFAULT ((0)) FOR [obsolete]
GO

