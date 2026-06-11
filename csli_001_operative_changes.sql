USE [RPA_Processes_DEV]
GO

/****** Object:  Table [dbo].[csli_001_operative_changes]    Script Date: 2026. 06. 11. 11:04:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[csli_001_operative_changes](
	[change_id] [int] IDENTITY(1,1) NOT NULL,
	[kuldemeny_azonosito_13] [nvarchar](13) NOT NULL,
	[source_table] [nvarchar](100) NOT NULL,
	[field_name] [nvarchar](100) NOT NULL,
	[old_value] [nvarchar](max) NULL,
	[new_value] [nvarchar](max) NULL,
	[changed_at] [datetime2](7) NOT NULL,
	[changed_by] [nvarchar](100) NOT NULL,
 CONSTRAINT [PK_csli_001_operative_changes] PRIMARY KEY CLUSTERED 
(
	[change_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [dbo].[csli_001_operative_changes] ADD  CONSTRAINT [DF_op_changes_changed_at]  DEFAULT (getdate()) FOR [changed_at]
GO

