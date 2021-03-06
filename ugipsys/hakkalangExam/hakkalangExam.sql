SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[hakkalangExamTopic]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[hakkalangExamTopic](
	[et_id] [int] IDENTITY(1,1) NOT NULL,
	[icuitem] [int] NOT NULL,
	[tune_id] [nvarchar](20) NOT NULL,
	[et_correct] [nvarchar](50) NULL,
 CONSTRAINT [PK_hakkalangExamTopic] PRIMARY KEY CLUSTERED 
(
	[et_id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[hakkalangExamOption]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[hakkalangExamOption](
	[et_id] [int] NOT NULL,
	[eo_title] [nvarchar](50) NOT NULL,
	[eo_answer] [nvarchar](50) NULL,
	[eo_sort] [tinyint] NOT NULL,
 CONSTRAINT [PK_hakkalangExamOption] PRIMARY KEY CLUSTERED 
(
	[et_id] ASC,
	[eo_sort] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
