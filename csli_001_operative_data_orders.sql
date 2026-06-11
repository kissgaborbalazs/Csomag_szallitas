USE [RPA_Processes_DEV]
GO

/****** Object:  Table [dbo].[csli_001_operative_data_orders]    Script Date: 2026. 06. 11. 11:04:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[csli_001_operative_data_orders](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[kuldemeny_azonosito_13] [nvarchar](13) NOT NULL,
	[felrako_nev] [nvarchar](100) NULL,
	[felrako_irsz] [nvarchar](10) NULL,
	[felrako_varos] [nvarchar](60) NULL,
	[felrako_kozterulet_elnevezese] [nvarchar](60) NULL,
	[felrako_kozterulet_tipus] [nvarchar](60) NULL,
	[felrako_hazszam] [nvarchar](60) NULL,
	[felrako_megjegyzes] [nvarchar](255) NULL,
	[felrako_atveteli_datum] [date] NULL,
	[felrako_idokapu] [nvarchar](20) NULL,
	[felrako_kontakt_szemely] [nvarchar](60) NULL,
	[felrako_telefonszam] [nvarchar](20) NULL,
	[felrako_email] [nvarchar](60) NULL,
	[raklap_hossza_cm] [int] NULL,
	[raklap_szelessege_cm] [int] NULL,
	[raklap_magassaga_cm] [int] NULL,
	[ekaer] [nvarchar](30) NULL,
	[cimirat_nyomtatas] [bit] NULL,
	[adr_jelolo] [nvarchar](30) NULL,
	[adr_pont] [nvarchar](30) NULL,
	[lerakas_idokapu] [nvarchar](20) NULL,
	[created_at] [datetime2](3) NOT NULL,
	[created_by] [nvarchar](100) NOT NULL,
	[obsolete] [bit] NOT NULL,
 CONSTRAINT [PK_csli_001_operative_data_orders] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[csli_001_operative_data_orders] ADD  CONSTRAINT [DF_op_data_orders_created_at]  DEFAULT (getdate()) FOR [created_at]
GO

ALTER TABLE [dbo].[csli_001_operative_data_orders] ADD  CONSTRAINT [DF_op_data_orders_obsolete]  DEFAULT ((0)) FOR [obsolete]
GO

