USE [RPA_Processes_DEV]
GO

/****** Object:  Table [dbo].[csli_001_raw_efj]    Script Date: 2026. 06. 11. 11:06:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[csli_001_raw_efj](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[efj_id] [nvarchar](255) NULL,
	[efeladojegyzek_azon] [nvarchar](255) NULL,
	[betoltes_datum] [datetime2](0) NULL,
	[efj_zaras] [nvarchar](255) NULL,
	[efj_szoftver] [nvarchar](255) NULL,
	[felado_vevokod] [nvarchar](255) NULL,
	[felado_megallapodas] [nvarchar](255) NULL,
	[felado_nev] [nvarchar](255) NULL,
	[felado_irsz] [nvarchar](4) NULL,
	[felado_hely] [nvarchar](255) NULL,
	[felado_kozelebbi_cim] [nvarchar](255) NULL,
	[felado_kozterulet_nev] [nvarchar](255) NULL,
	[felado_kozterulet_jelleg] [nvarchar](255) NULL,
	[felado_hazszam] [nvarchar](255) NULL,
	[felado_megjegyzes] [nvarchar](255) NULL,
	[felado_email] [nvarchar](255) NULL,
	[felado_telefon] [nvarchar](255) NULL,
	[varhato_feladas_datum] [nvarchar](255) NULL,
	[sorszam] [nvarchar](255) NULL,
	[kuldemeny_azonosito] [nvarchar](26) NOT NULL,
	[kuldemeny_azonosito_13]  AS (left([kuldemeny_azonosito],(13))) PERSISTED,
	[vezer_azon] [nvarchar](26) NULL,
	[alapszolg] [nvarchar](255) NULL,
	[kuldemeny_brutto_tomege_kg] [int] NOT NULL,
	[kezbesites_modja] [nvarchar](10) NULL,
	[szolgaltatas] [nvarchar](255) NULL,
	[utanvet_osszege_ft] [decimal](10, 2) NOT NULL,
	[kuldemeny_aruertek] [decimal](18, 2) NOT NULL,
	[lerako_referencia_szam] [nvarchar](20) NULL,
	[cimzett_nev] [nvarchar](200) NULL,
	[cimzett_irsz] [nvarchar](10) NULL,
	[cimzett_varos] [nvarchar](60) NULL,
	[cimzett_kozerlebbi_cim] [nvarchar](255) NULL,
	[cimzett_kozterulet_nev] [nvarchar](60) NULL,
	[cimzett_kozterulet_jelleg] [nvarchar](60) NULL,
	[cimzett_hazszam] [nvarchar](60) NULL,
	[cimzett_megjegyzes] [nvarchar](255) NULL,
	[cimzett_kontakt_szemely] [nvarchar](60) NULL,
	[cimzett_email] [nvarchar](60) NULL,
	[cimzett_telefonszam] [nvarchar](20) NULL,
	[ugyfadat1] [nvarchar](255) NULL,
	[ugyfadat2] [nvarchar](255) NULL,
	[created_at] [datetime2](3) NOT NULL,
	[created_by] [nvarchar](100) NOT NULL,
	[modified_at] [datetime2](3) NOT NULL,
	[modified_by] [nvarchar](100) NOT NULL,
	[obsolete] [bit] NOT NULL,
 CONSTRAINT [PK_csli_001_raw_efj] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[csli_001_raw_efj] ADD  CONSTRAINT [DF_raw_efj_created_at]  DEFAULT (getdate()) FOR [created_at]
GO

ALTER TABLE [dbo].[csli_001_raw_efj] ADD  CONSTRAINT [DF_raw_efj_modified_at]  DEFAULT (getdate()) FOR [modified_at]
GO

ALTER TABLE [dbo].[csli_001_raw_efj] ADD  CONSTRAINT [DF_raw_efj_obsolete]  DEFAULT ((0)) FOR [obsolete]
GO

