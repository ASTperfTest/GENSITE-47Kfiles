USE [PLANT_LOG]
GO

IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_OWNER_SCORE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[OWNER] DROP CONSTRAINT [DF_OWNER_SCORE]
END

GO

USE [PLANT_LOG]
GO

/****** Object:  Table [dbo].[OWNER]    Script Date: 06/19/2009 18:17:07 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[OWNER]') AND type in (N'U'))
DROP TABLE [dbo].[OWNER]
GO

USE [PLANT_LOG]
GO

/****** Object:  Table [dbo].[OWNER]    Script Date: 06/19/2009 18:17:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[OWNER](
	[OWNER_ID] [nvarchar](255) NOT NULL,
	[DISPLAY_NAME] [nvarchar](255) NOT NULL,
	[NICKNAME] [nvarchar](255),
	[TOPIC] [nvarchar](50) NOT NULL,
	[EMAIL] [nvarchar](255),
	[DESCRIPTION] [ntext],
	[IS_APPROVE] bit DEFAULT (1) NOT NULL,
	[CREATOR_ID] [nvarchar](255) NOT NULL,
	[CREATE_DATETIME] [datetime] NOT NULL,
	[MODIFIER_ID] [nvarchar](255) NOT NULL,
	[MODIFY_DATETIME] [datetime] NOT NULL,
 CONSTRAINT [PK_OWNER] PRIMARY KEY CLUSTERED 
(
	[OWNER_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]

GO

USE [PLANT_LOG]
GO

/****** Object:  Table [dbo].[ENTRY]    Script Date: 06/19/2009 18:17:19 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ENTRY]') AND type in (N'U'))
DROP TABLE [dbo].[ENTRY]
GO

USE [PLANT_LOG]
GO

/****** Object:  Table [dbo].[ENTRY]    Script Date: 06/19/2009 18:17:19 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ENTRY](
	[ENTRY_ID] [nvarchar](255) NOT NULL,
	[OWNER_ID] [nvarchar](255) NOT NULL,
	[DATE] [datetime] NOT NULL,
	[TITLE] [nvarchar](255) NULL,
	[DESCRIPTION] [ntext] NULL,
	[IS_PUBLIC] bit DEFAULT (1) NOT NULL,
	[IS_APPROVE] bit DEFAULT (1) NOT NULL,
	[CREATOR_ID] [nvarchar](255) NOT NULL,
	[CREATE_DATETIME] [datetime] NOT NULL,
	[MODIFIER_ID] [nvarchar](255) NOT NULL,
	[MODIFY_DATETIME] [datetime] NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[ENTRY_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]

GO

USE [PLANT_LOG]
GO

/****** Object:  Table [dbo].[IMG_FILE]    Script Date: 06/19/2009 18:17:22 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[IMG_FILE]') AND type in (N'U'))
DROP TABLE [dbo].[IMG_FILE]
GO

USE [PLANT_LOG]
GO

/****** Object:  Table [dbo].[IMG_FILE]    Script Date: 06/19/2009 18:17:22 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[IMG_FILE](
	[FILE_ID] [nvarchar](255) NOT NULL,
	[ENTRY_ID] [nvarchar](255) NOT NULL,
	[NAME] [nvarchar](255) NOT NULL,
	[URI] [nvarchar](255) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[FILE_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]

GO

USE [PLANT_LOG]
GO

/****** Object:  Table [dbo].[VOTE_RECORD]    Script Date: 06/22/2009 11:13:59 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[VOTE_RECORD]') AND type in (N'U'))
DROP TABLE [dbo].[VOTE_RECORD]
GO

USE [PLANT_LOG]
GO

/****** Object:  Table [dbo].[VOTE_RECORD]    Script Date: 06/22/2009 11:13:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[VOTE_RECORD](
	[RECORD_ID] [int] IDENTITY(1,1) NOT NULL,
	[IP] [nvarchar](50) NOT NULL,
	[VOTE_DATE] [datetime] NOT NULL,
	[VOTER_ID] [nvarchar](50) NOT NULL,
	[USER_ID] [nvarchar](50) NOT NULL,
CONSTRAINT [PK_VOTE_RECORD] PRIMARY KEY CLUSTERED 
(
	[RECORD_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]

GO