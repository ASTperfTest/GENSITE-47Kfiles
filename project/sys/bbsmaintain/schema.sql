if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Article]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Article]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[talk02]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[talk02]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[talk03]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[talk03]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[talk05]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[talk05]
GO

CREATE TABLE [dbo].[Article] (
	[Id] [int] NOT NULL ,
	[MId] [int] NOT NULL ,
	[SId] [int] NOT NULL ,
	[Title] [char] (100) COLLATE Chinese_Taiwan_Stroke_CI_AS NOT NULL ,
	[Message] [text] COLLATE Chinese_Taiwan_Stroke_CI_AS NOT NULL ,
	[Date] [datetime] NOT NULL ,
	[BId] [int] NOT NULL ,
	[AuthorID] [int] NOT NULL ,
	[NickName] [char] (30) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[ip] [char] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[isethics] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[cnt] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[talk02] (
	[content] [text] COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[updatetime] [datetime] NULL ,
	[updatename] [char] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[talk03] (
	[content] [text] COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[updatetime] [datetime] NULL ,
	[updatename] [char] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[talk05] (
	[id] [int] NOT NULL ,
	[name] [varchar] (250) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[master] [varchar] (250) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[updatetime] [datetime] NULL ,
	[seq] [int] NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Article] WITH NOCHECK ADD 
	CONSTRAINT [PK_Article] PRIMARY KEY  CLUSTERED 
	(
		[Id],
		[MId],
		[SId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[talk05] WITH NOCHECK ADD 
	CONSTRAINT [PK_talk05] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Article] ADD 
	CONSTRAINT [DF_Article_cnt] DEFAULT (0) FOR [cnt]
GO

