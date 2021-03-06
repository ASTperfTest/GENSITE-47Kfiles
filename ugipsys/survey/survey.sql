if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[m011]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[m011]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[m012]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[m012]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[m013]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[m013]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[m014]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[m014]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[m015]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[m015]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[m016]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[m016]
GO

CREATE TABLE [dbo].[m011] (
	[giCuItem] [int] NOT NULL ,
	[m011_subjectid] [int] NULL ,
	[m011_subject] [nchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011_questno] [int] NULL ,
	[m011_bdate] [datetime] NULL ,
	[m011_edate] [datetime] NULL ,
	[m011_notetype] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011_haveprize] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011_jumpquestion] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011_onlyonce] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011_pflag] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[m012] (
	[m012_subjectid] [int] NOT NULL ,
	[m012_questionid] [int] NOT NULL ,
	[m012_title] [nvarchar] (3000) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m012_answerno] [int] NULL ,
	[m012_type] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m012_modifyuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m012_updatetime] [datetime] NULL ,
	[m012_createuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m012_createtime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[m013] (
	[m013_subjectid] [int] NOT NULL ,
	[m013_questionid] [int] NOT NULL ,
	[m013_answerid] [int] NOT NULL ,
	[m013_title] [nvarchar] (3000) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m013_default] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m013_nextord] [int] NULL ,
	[m013_no] [int] NULL ,
	[m013_modifyuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m013_updatetime] [datetime] NULL ,
	[m013_createuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m013_createtime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[m014] (
	[m014_id] [int] NOT NULL ,
	[m014_A] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[m014_B] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014_C] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014_D] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014_E] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014_F] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014_G] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014_H] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014_pflag] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014_reply] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014_subjectid] [int] NOT NULL ,
	[m014_polldate] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[m015] (
	[m015_id] [int] IDENTITY (1, 1) NOT NULL ,
	[m015_subjectid] [int] NOT NULL ,
	[m015_questionid] [int] NOT NULL ,
	[m015_answerid] [int] NOT NULL ,
	[m015_modifyuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[m015_updatetime] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[m016] (
	[m016_id] [int] IDENTITY (1, 1) NOT NULL ,
	[m016_subjectid] [int] NOT NULL ,
	[m016_questionid] [int] NOT NULL ,
	[m016_answerid] [int] NOT NULL ,
	[m016_userid] [int] NULL ,
	[m016_content] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[m016_updatetime] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

