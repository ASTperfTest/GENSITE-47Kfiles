if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AdRotate]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AdRotate]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Ap]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Ap]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Apcat]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Apcat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AskOut]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AskOut]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Banner]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Banner]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BaseDsd]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BaseDsd]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BaseDsdfield]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BaseDsdfield]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CatTreeNode]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CatTreeNode]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CatTreeRoot]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CatTreeRoot]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Catalogue]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Catalogue]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm01]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm01]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm02]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm02]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm03]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm03]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm04]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm04]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm05]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm05]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm06]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm06]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm07]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm07]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm08]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm08]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm09]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm09]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm10]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm10]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm11]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm11]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm12]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm12]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cm13]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cm13]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Code1]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Code1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CodeHt]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CodeHt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CodeMain]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CodeMain]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CodeMainLong]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CodeMainLong]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CodeMetaDef]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CodeMetaDef]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Counter]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Counter]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CtUnit]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CtUnit]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CtUserSet]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CtUserSet]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CtUserSet2]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CtUserSet2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CtUserSet3]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CtUserSet3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDTx9]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDTx9]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtGeneric]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtGeneric]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtattach]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtattach]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtkeyword]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtkeyword]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtkeywordWtp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtkeywordWtp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtpage]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtpage]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtwImg]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtwImg]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtx2]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtx2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtx27]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtx27]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtx3]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtx3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtx30]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtx30]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtx4]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtx4]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtx5]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtx5]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtx7]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtx7]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtx8]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtx8]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtxDformSpec]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtxDformSpec]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtxMmo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtxMmo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CuDtxPaper]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CuDtxPaper]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DataCat]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DataCat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DataContent]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DataContent]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DataUnit]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DataUnit]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dept]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dept]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DeptOrg]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DeptOrg]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[EpPub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[EpPub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[EpSend]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[EpSend]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Epaper]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Epaper]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FileUp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FileUp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Forum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Forum]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ForumArticle]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ForumArticle]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ForumEssenceArticle]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ForumEssenceArticle]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GipSites]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[GipSites]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HfAptForm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HfAptForm]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HfAptJob]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HfAptJob]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HfDeptSet]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HfDeptSet]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HtDdb]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HtDdb]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HtDentity]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HtDentity]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HtDfield]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HtDfield]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ImageFile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ImageFile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[InfoUser]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[InfoUser]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[JobOrg]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[JobOrg]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ListConten]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ListConten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[M011]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[M011]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[M012]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[M012]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[M013]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[M013]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[M014]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[M014]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[M015]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[M015]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[M016]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[M016]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[M074]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[M074]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Mall]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Mall]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MemCatTree]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MemCatTree]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MemEpaper]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MemEpaper]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Member]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Member]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Mmofolder]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Mmofolder]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Mmosite]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Mmosite]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OrgVersion]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[OrgVersion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Poll]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Poll]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Polluser]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Polluser]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt6]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Rpt6]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Ugrp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Ugrp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UgrpAp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UgrpAp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UgrpCode]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UgrpCode]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UpLoadSite]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UpLoadSite]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[WorkDay]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[WorkDay]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[X2mainToSub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[X2mainToSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[X9postCat]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[X9postCat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[X9temp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[X9temp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[XcatNew]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[XcatNew]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[XdmpList]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[XdmpList]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[XfUpload]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[XfUpload]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZcustomDbinfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ZcustomDbinfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZcustomDbstatus]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ZcustomDbstatus]
GO

CREATE TABLE [dbo].[AdRotate] (
	[giCuItem] [int] NOT NULL ,
	[adfixSeq] [tinyint] NULL ,
	[adweight] [tinyint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Ap] (
	[apcode] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[apnameE] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[apnameC] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[apcat] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[specPath] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[appath] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[aporder] [nvarchar] (3) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[apmask] [tinyint] NOT NULL ,
	[spare64] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[spare128] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xsNewWindow] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xsSubmit] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Apcat] (
	[apcatId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[apcatCname] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[apcatEname] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[apseq] [tinyint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AskOut] (
	[askOutIid] [int] NOT NULL ,
	[aoUserId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xdept] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[aoKind] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[aoDate] [smalldatetime] NOT NULL ,
	[aoCount] [tinyint] NOT NULL ,
	[aoDesc] [nvarchar] (500) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Banner] (
	[bannerid] [int] NOT NULL ,
	[seq] [int] NULL ,
	[bannerPic] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[bannerPicRealname] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[bannerUrl] [nchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[bannerTxt] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[bannerOnline] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[createtime] [datetime] NULL ,
	[createuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[modifytime] [datetime] NULL ,
	[modifyuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[BaseDsd] (
	[ibaseDsd] [int] IDENTITY (1, 1) NOT NULL ,
	[sbaseTableName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[sbaseDsdname] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[tdesc] [nvarchar] (300) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[rdsdcat] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[inUse] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[sbaseDsdxml] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BaseDsdfield] (
	[ibaseDsd] [int] NOT NULL ,
	[ibaseField] [int] IDENTITY (1, 1) NOT NULL ,
	[xfieldName] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xfieldSeq] [smallint] NULL ,
	[inUse] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xfieldLabel] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xfieldDesc] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xdataType] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xdataLen] [int] NULL ,
	[xcanNull] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xisPrimaryKey] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xidentity] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xdefaultvalue] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xinputType] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xclientDefault] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xrefLookup] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xrows] [smallint] NULL ,
	[xcols] [smallint] NULL ,
	[xinputLen] [smallint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CatTreeNode] (
	[ctRootId] [int] NOT NULL ,
	[ctNodeKind] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[ctNodeId] [int] IDENTITY (1, 1) NOT NULL ,
	[catName] [nvarchar] (70) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[ctNameLogo] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[catShowOrder] [int] NULL ,
	[dataLevel] [tinyint] NOT NULL ,
	[dataParent] [int] NULL ,
	[childCount] [int] NOT NULL ,
	[dataCount] [int] NOT NULL ,
	[ctUnitId] [int] NULL ,
	[inUse] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[editDate] [smalldatetime] NOT NULL ,
	[editUserId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[dcondition] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xslList] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xslData] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ctNodeNpkind] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[rss] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CatTreeRoot] (
	[ctRootId] [int] IDENTITY (1, 1) NOT NULL ,
	[ctRootName] [nvarchar] (40) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[vgroup] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[purpose] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[inUse] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[editDate] [smalldatetime] NOT NULL ,
	[editor] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deptId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[pvXdmp] [nvarchar] (6) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Catalogue] (
	[itemId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[catId] [int] NOT NULL ,
	[catName] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[catShowOrder] [int] NULL ,
	[editDate] [smalldatetime] NOT NULL ,
	[editUserId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[dataLevel] [int] NOT NULL ,
	[dataTop] [int] NOT NULL ,
	[dataParent] [int] NOT NULL ,
	[childCount] [int] NOT NULL ,
	[dataCount] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm01] (
	[giCuItem] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm02] (
	[giCuItem] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm03] (
	[giCuItem] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm04] (
	[giCuItem] [int] NOT NULL ,
	[cuDtx12f15] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx12f16] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx12f17] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm05] (
	[giCuItem] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm06] (
	[giCuItem] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm07] (
	[giCuItem] [int] NOT NULL ,
	[cuDtx15f18] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f19] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f20] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f21] [nvarchar] (60) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f22] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f23] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f24] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f25] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f26] [nvarchar] (800) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f27] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f28] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f29] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f30] [nvarchar] (800) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f31] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx15f32] [smalldatetime] NULL ,
	[cuDtx15f33] [smalldatetime] NULL ,
	[cuDtx15f34] [nvarchar] (800) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm08] (
	[giCuItem] [int] NOT NULL ,
	[cuDtx16f35] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx16f36] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm09] (
	[giCuItem] [int] NOT NULL ,
	[cuDtx17f37] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx17f38] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx17f39] [smalldatetime] NULL ,
	[cuDtx17f40] [int] NULL ,
	[cuDtx17f41] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm10] (
	[giCuItem] [int] NOT NULL ,
	[cuDtx18f42] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx18f43] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx18f44] [nvarchar] (800) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx18f45] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm11] (
	[giCuItem] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm12] (
	[giCuItem] [int] NOT NULL ,
	[cuDtx20f46] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cm13] (
	[giCuItem] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Code1] (
	[item] [nvarchar] (5) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[descrip] [nvarchar] (140) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[code] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[hist] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[freq] [nvarchar] (3) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[aux1] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[chPosDegree] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[numDutyWhrs] [numeric](5, 2) NULL ,
	[numDeduWhrs] [numeric](5, 2) NULL ,
	[rjob] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[qjobYn] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CodeHt] (
	[item] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[code] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[descrip] [nvarchar] (140) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[sortfld] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CodeMain] (
	[codeMetaId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[mcode] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[mvalue] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[msortValue] [nvarchar] (6) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CodeMainLong] (
	[codeMetaId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[mcode] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[mvalue] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[msortValue] [nvarchar] (6) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mlongValue] [nvarchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CodeMetaDef] (
	[codeId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[codeName] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[codeTblName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[codeSrcFld] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[codeSrcItem] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[codeValueFld] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[codeDisplayFld] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[codeSortFld] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[codeValueFldName] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[codeDisplayFldName] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[codeSortFldName] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[codeType] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[codeRank] [nvarchar] (3) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[showOrNot] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[codeXmlspec] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Counter] (
	[mp] [int] NOT NULL ,
	[counts] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CtUnit] (
	[ctUnitId] [int] IDENTITY (1, 1) NOT NULL ,
	[ctUnitName] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ctUnitLogo] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[headerPart] [int] NULL ,
	[footerPart] [int] NULL ,
	[fstyle] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[bgLogo] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[bgColor] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ctUnitKind] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[redirectUrl] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[newWindow] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ibaseDsd] [int] NULL ,
	[fctUnitOnly] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[inUse] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[unitDesc] [nvarchar] (60) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[checkYn] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deptId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ctunitexpireday] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CtUserSet] (
	[userId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[ctNodeId] [int] NOT NULL ,
	[rights] [smallint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CtUserSet2] (
	[userId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[ctNodeId] [int] NOT NULL ,
	[rights] [smallint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CtUserSet3] (
	[userId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[ctNodeId] [int] NOT NULL ,
	[rights] [smallint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDTx9] (
	[giCuItem] [int] NOT NULL ,
	[mmofolderName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mmofileName] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mmofileType] [nvarchar] (2) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mmofileSize] [smallmoney] NULL ,
	[mmofileAlt] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mmoimageHeight] [int] NULL ,
	[mmoimageWidth] [int] NULL ,
	[mmofileIcon] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mmositeId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtGeneric] (
	[icuitem] [int] IDENTITY (1, 1) NOT NULL ,
	[ibaseDsd] [int] NOT NULL ,
	[ictunit] [int] NULL ,
	[fctupublic] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[avBegin] [smalldatetime] NULL ,
	[avEnd] [smalldatetime] NULL ,
	[stitle] [nvarchar] (500) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[ieditor] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deditDate] [smalldatetime] NOT NULL ,
	[idept] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[topCat] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[vgroup] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xkeyword] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ximportant] [int] NULL ,
	[xurl] [nvarchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xnewWindow] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xpostDate] [smalldatetime] NOT NULL ,
	[xbody] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xpostDateEnd] [smalldatetime] NULL ,
	[createdDate] [smalldatetime] NULL ,
	[xabstract] [nvarchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[fileDownLoad] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[refId] [int] NULL ,
	[showType] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ximgFile] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtattach] (
	[ixCuAttach] [int] IDENTITY (1, 1) NOT NULL ,
	[xiCuItem] [int] NOT NULL ,
	[atitle] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[adesc] [nvarchar] (800) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ofileName] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[nfileName] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[aeditor] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[aeditDate] [smalldatetime] NOT NULL ,
	[blist] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[listSeq] [nvarchar] (2) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtkeyword] (
	[icuitem] [int] NOT NULL ,
	[xkeyword] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[weight] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtkeywordWtp] (
	[ikeyword] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[icount] [int] NULL ,
	[ikeywordNew] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[keywordStatus] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[editor] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[editDate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtpage] (
	[ixCuPage] [int] IDENTITY (1, 1) NOT NULL ,
	[xiCuItem] [int] NOT NULL ,
	[atitle] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[adesc] [nvarchar] (800) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[npageId] [int] NOT NULL ,
	[blist] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[listSeq] [nvarchar] (2) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtwImg] (
	[giCuItem] [int] NOT NULL ,
	[ximgFile] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtx2] (
	[giCuItem] [int] NOT NULL ,
	[cuDtx2f5] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx2f7] [nvarchar] (300) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx2f9] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtx27] (
	[giCuItem] [int] NOT NULL ,
	[cuDtx24f51] [nvarchar] (512) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx24f52] [nvarchar] (16) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx24f53] [nvarchar] (512) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx24f54] [nvarchar] (1000) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx24f55] [nvarchar] (2000) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx24f56] [nvarchar] (2000) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtx3] (
	[gicuitem] [int] NOT NULL ,
	[gicuitem123] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtx30] (
	[giCuItem] [int] NOT NULL ,
	[cuDtx9f16] [nvarchar] (250) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx9f18] [nvarchar] (250) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx9f19] [nvarchar] (250) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[cuDtx9f121] [nvarchar] (250) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtx4] (
	[giCuItem] [int] NOT NULL ,
	[cuDtx4f14] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtx5] (
	[giCuItem] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtx7] (
	[giCuItem] [int] NOT NULL ,
	[cuDtx7f8] [nvarchar] (4000) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtx8] (
	[giCuItem] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtxDformSpec] (
	[giCuItem] [int] NOT NULL ,
	[xmlSpec] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[phtml] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[expireDayNum] [int] NOT NULL ,
	[formScript] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtxMmo] (
	[icuitem] [int] NOT NULL ,
	[fctupublic] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[avBegin] [smalldatetime] NULL ,
	[avEnd] [smalldatetime] NULL ,
	[stitle] [nvarchar] (500) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[ieditor] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deditDate] [smalldatetime] NOT NULL ,
	[topCat] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[vgroup] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xkeyword] [nvarchar] (40) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ximportant] [int] NULL ,
	[xurl] [nvarchar] (160) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xnewWindow] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xpostDate] [smalldatetime] NOT NULL ,
	[mmocat] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[mmofileExt] [nvarchar] (2) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[mmofileTime] [int] NULL ,
	[mmofilePath] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mmofolderName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xiICuitem] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CuDtxPaper] (
	[giCuItem] [int] NOT NULL ,
	[author] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DataCat] (
	[catId] [int] NOT NULL ,
	[language] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[dataType] [nvarchar] (3) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[catName] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[catShowOrder] [int] NULL ,
	[editDate] [smalldatetime] NULL ,
	[editUserId] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DataContent] (
	[unitId] [int] NOT NULL ,
	[contentId] [int] NOT NULL ,
	[content] [nvarchar] (4000) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[imageFile] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[imageWay] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[position] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DataUnit] (
	[unitId] [int] NOT NULL ,
	[dataType] [nvarchar] (3) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[catId] [int] NULL ,
	[language] [nvarchar] (3) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[subject] [nvarchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[showOrder] [int] NULL ,
	[beginDate] [smalldatetime] NULL ,
	[endDate] [smalldatetime] NULL ,
	[editDate] [smalldatetime] NOT NULL ,
	[editUserId] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[extend1] [nvarchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[extend2] [nvarchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[extend3] [nvarchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dept] (
	[giCuItem] [int] NULL ,
	[deptId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deptName] [nvarchar] (70) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[abbrName] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[edeptName] [nvarchar] (60) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[eabbrName] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[orgRank] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[kind] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deptHead] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[inUse] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[tdataCat] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[servPhone] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[servAddr] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[servHour] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[eservAddr] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[eservHour] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[parent] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[nodeKind] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[plevel] [smallint] NULL ,
	[seq] [smallint] NULL ,
	[deptCode] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[codeName] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[webSite] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DeptOrg] (
	[verId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deptId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deptName] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[parent] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[kind] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[plevel] [smallint] NOT NULL ,
	[seq] [smallint] NOT NULL ,
	[child] [smallint] NOT NULL ,
	[deptHead] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[uniCod] [nvarchar] (6) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[uniAname] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[uniEname] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[uniLevel] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[uniRef] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[upUniRefDeptId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[staffCat] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[EpPub] (
	[epubId] [int] IDENTITY (1, 1) NOT NULL ,
	[pubDate] [smalldatetime] NOT NULL ,
	[title] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[dbDate] [smalldatetime] NOT NULL ,
	[deDate] [smalldatetime] NOT NULL ,
	[maxNo] [int] NOT NULL ,
	[ctRootId] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[EpSend] (
	[epubId] [int] NOT NULL ,
	[email] [nvarchar] (60) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[sendDate] [smalldatetime] NOT NULL ,
	[ctRootId] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Epaper] (
	[email] [nchar] (60) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[createtime] [datetime] NULL ,
	[ctRootId] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FileUp] (
	[itemId] [nvarchar] (3) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[parentId] [int] NOT NULL ,
	[ofileName] [nvarchar] (255) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[nfileName] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Forum] (
	[itemId] [nvarchar] (3) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[forumId] [int] NOT NULL ,
	[catId] [int] NULL ,
	[forumName] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[forumMaster] [int] NOT NULL ,
	[forumRemark] [nvarchar] (4000) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[masterPhoto] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[forumShowOrder] [int] NOT NULL ,
	[editDate] [smalldatetime] NOT NULL ,
	[editUserId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ForumArticle] (
	[forumId] [int] NOT NULL ,
	[articleId] [int] NOT NULL ,
	[subject] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[content] [nvarchar] (4000) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[postUserId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[postUserName] [nvarchar] (60) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[postEmail] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[postDate] [smalldatetime] NOT NULL ,
	[articleLevel] [int] NOT NULL ,
	[articleTop] [int] NOT NULL ,
	[articleParent] [int] NOT NULL ,
	[articleChildCount] [int] NOT NULL ,
	[childrenCount] [int] NOT NULL ,
	[lastReArticleId] [int] NULL ,
	[lastPostDate] [smalldatetime] NULL ,
	[readCount] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ForumEssenceArticle] (
	[forumId] [int] NOT NULL ,
	[articleId] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[GipSites] (
	[gipSiteId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[gipSiteName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[gipSiteDbconn] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[gipSiteUrl] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[gipSiteUrlsys] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[gipSiteInnerRoot] [int] NOT NULL ,
	[gipSiteDefaultRoot] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[HfAptForm] (
	[hfAptId] [int] NOT NULL ,
	[hyFormId] [int] NOT NULL ,
	[toDept] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[aplName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[aplId] [nvarchar] (15) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[aplType] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[billNum] [nvarchar] (14) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[caseNum] [nvarchar] (16) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[aplDate] [smalldatetime] NOT NULL ,
	[status] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[lastDate] [smalldatetime] NOT NULL ,
	[jrcvDate] [smalldatetime] NULL ,
	[jrcvNo] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[jdept] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[juser] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[jexpDate] [smalldatetime] NULL ,
	[jdoneDate] [smalldatetime] NULL ,
	[jmailText] [nvarchar] (800) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[formXml] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[aplCusId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[jrcvPhone] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[transDept] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[jdept2] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[HfAptJob] (
	[hfAptJobId] [int] NOT NULL ,
	[hfJstatus] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[hfJdate] [smalldatetime] NOT NULL ,
	[hfJuser] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[rtMail] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[hfJaptId] [int] NOT NULL ,
	[hfJorgDept] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[HfDeptSet] (
	[hyFormId] [int] NOT NULL ,
	[deptId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[userId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[editUser] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[editDate] [smalldatetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[HtDdb] (
	[dbId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[dbName] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[dbDesc] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[dbIp] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[HtDentity] (
	[dbId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[entityId] [int] IDENTITY (1, 1) NOT NULL ,
	[tableName] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[entityDesc] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[entityUri] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[apcatId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[HtDfield] (
	[entityId] [int] NOT NULL ,
	[xfieldName] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xfieldSeq] [smallint] NULL ,
	[xfieldLabel] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xfieldDesc] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xdataType] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xdataLen] [int] NULL ,
	[xcanNull] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xisPrimaryKey] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xidentity] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xdefaultvalue] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xinputType] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xclientDefault] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xrefLookup] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xrows] [smallint] NULL ,
	[xcols] [smallint] NULL ,
	[xinputLen] [smallint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ImageFile] (
	[newFileName] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[oldFileName] [nvarchar] (60) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[filePath] [nvarchar] (150) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[InfoUser] (
	[userId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[password] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[userName] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[userType] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deptId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[jobName] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[email] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ugrpId] [nvarchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[lastVisit] [smalldatetime] NULL ,
	[visitCount] [smallint] NOT NULL ,
	[ugrpName] [nchar] (500) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[telephone] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[tdataCat] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[uploadPath] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[lastIp] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[modifyPassword] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[JobOrg] (
	[verId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[jobId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[titcod] [nvarchar] (4) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[jobName] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deptId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deptManager] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[upJobId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[quotaType] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[quotaAssume] [smallint] NULL ,
	[quotaApprove] [smallint] NULL ,
	[staffDiv] [nchar] (2) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[staffKind] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[qualification] [nvarchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[chfCod] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[i23tittype] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[quotaOrg] [smallint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ListConten] (
	[giCuItem] [int] NOT NULL ,
	[data1] [nvarchar] (512) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[data2] [nvarchar] (512) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[data3] [nvarchar] (512) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[data4] [nvarchar] (512) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[data5] [nvarchar] (512) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[data6] [nvarchar] (512) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[data7] [nvarchar] (512) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[data8] [nvarchar] (512) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[data9] [nvarchar] (512) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[data10] [nvarchar] (512) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[M011] (
	[m011Subjectid] [int] NOT NULL ,
	[m011Subject] [nchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011Questno] [int] NULL ,
	[m011Bdate] [datetime] NULL ,
	[m011Edate] [datetime] NULL ,
	[m011Online] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011Notetype] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011Questionnote] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011Public] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011Haveprize] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011Pflag] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011Jumpquestion] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011Onlyonce] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011Createuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011Createtime] [datetime] NULL ,
	[m011Modifyuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m011Updatetime] [datetime] NULL ,
	[m011Dept] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mallId] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[M012] (
	[m012Subjectid] [int] NOT NULL ,
	[m012Questionid] [int] NOT NULL ,
	[m012Title] [nvarchar] (3000) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m012Answerno] [int] NULL ,
	[m012Type] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m012Createuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m012Createtime] [datetime] NULL ,
	[m012Modifyuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m012Updatetime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[M013] (
	[m013Subjectid] [int] NOT NULL ,
	[m013Questionid] [int] NOT NULL ,
	[m013Answerid] [int] NOT NULL ,
	[m013Title] [nvarchar] (3000) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m013Default] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m013Nextord] [int] NULL ,
	[m013No] [int] NULL ,
	[m013Createuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m013Createtime] [datetime] NULL ,
	[m013Modifyuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m013Updatetime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[M014] (
	[m014Id] [int] NOT NULL ,
	[m014Name] [nchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[m014Sex] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014Idnumber] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014Email] [nchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014Age] [int] NULL ,
	[m014Addrarea] [int] NULL ,
	[m014Familymember] [int] NULL ,
	[m014Money] [int] NULL ,
	[m014Job] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014Edu] [nchar] (2) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014Pflag] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014Reply] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m014Subjectid] [int] NOT NULL ,
	[m014Polldate] [datetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[M015] (
	[m015Id] [int] NOT NULL ,
	[m015Subjectid] [int] NOT NULL ,
	[m015Questionid] [int] NOT NULL ,
	[m015Answerid] [int] NOT NULL ,
	[m015Modifyuser] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[m015Updatetime] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[M016] (
	[m016Id] [int] NOT NULL ,
	[m016Subjectid] [int] NOT NULL ,
	[m016Questionid] [int] NOT NULL ,
	[m016Answerid] [int] NOT NULL ,
	[m016Userid] [int] NULL ,
	[m016Content] [nchar] (1024) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[m016Updatetime] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[M074] (
	[m074Id] [int] NOT NULL ,
	[m074Table] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m074Qstr] [nvarchar] (500) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m074Adminid] [nchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m074Action] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[m074Time] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Mall] (
	[mallId] [nvarchar] (4) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mallName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[msortValue] [nvarchar] (6) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MemCatTree] (
	[memId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[ctNodeId] [int] NOT NULL ,
	[rights] [smallint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MemEpaper] (
	[memId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[ctNodeId] [int] NOT NULL ,
	[rights] [smallint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Member] (
	[account] [nchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[passwd] [nchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[realname] [nchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[homeaddr] [nchar] (500) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[phone] [nchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mobile] [nchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[email] [nchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[mcode] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[createtime] [datetime] NULL ,
	[modifytime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Mmofolder] (
	[mmofolderId] [int] IDENTITY (1, 1) NOT NULL ,
	[mmofolderName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mmofolderDesc] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ctUnitId] [int] NULL ,
	[mmositeId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[upLoadSite] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mmofolderParent] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Mmosite] (
	[mmositeId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[mmositeName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[mmositeUrl] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mmositeUrl2] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mmositeDesc] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mmositeSeq] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mmositeSeq2] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[upLoadSiteFtpip] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[upLoadSiteFtpport] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[upLoadSiteFtpid] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[upLoadSiteFtppwd] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[upLoadSiteHttp] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[OrgVersion] (
	[verId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[verType] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[startDate] [smalldatetime] NOT NULL ,
	[expireDate] [smalldatetime] NULL ,
	[approveNo] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[remark] [nvarchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Poll] (
	[pollId] [int] NOT NULL ,
	[content] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[answer] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[type] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[option1] [nvarchar] (250) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[option2] [nvarchar] (250) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[option3] [nvarchar] (250) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[option4] [nvarchar] (250) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[hint] [nvarchar] (250) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Polluser] (
	[sid] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[username] [nvarchar] (8) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[address] [nvarchar] (250) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[phone1] [nvarchar] (3) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[phone2] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[email] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[mobile] [nvarchar] (12) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Rpt6] (
	[rpt6id] [int] NOT NULL ,
	[userId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deptId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[odate] [smalldatetime] NOT NULL ,
	[oplace] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[osDate] [smalldatetime] NOT NULL ,
	[oeDate] [smalldatetime] NOT NULL ,
	[opurpose] [nvarchar] (200) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Ugrp] (
	[ugrpId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[ugrpName] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[remark] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[regdate] [smalldatetime] NOT NULL ,
	[signer] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[isPublic] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UgrpAp] (
	[ugrpId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[apcode] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[rights] [smallint] NOT NULL ,
	[regdate] [smalldatetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UgrpCode] (
	[ugrpId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[codeId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[rights] [smallint] NULL ,
	[regdate] [datetime] NULL ,
	[startdate] [datetime] NULL ,
	[enddate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UpLoadSite] (
	[upLoadSiteId] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[upLoadSiteName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[upLoadSiteFtpip] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[upLoadSiteFtpport] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[upLoadSiteFtpid] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[upLoadSiteFtppwd] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[upLoadSiteHttp] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[WorkDay] (
	[xday] [smalldatetime] NOT NULL ,
	[offMark] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xweekDay] [tinyint] NULL ,
	[xdayNo] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[X2mainToSub] (
	[id] [int] NOT NULL ,
	[belong] [int] NULL ,
	[deptName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ename] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[newDept] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[X9postCat] (
	[orgPcat] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[pcatId] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[X9temp] (
	[ibaseDsd] [int] NOT NULL ,
	[ictUnit] [int] NULL ,
	[avEnd] [smalldatetime] NULL ,
	[stitle] [nvarchar] (500) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[ieditor] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[idept] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[topCat] [nchar] (1) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xkeyword] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xurl] [nvarchar] (160) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xpostDate] [smalldatetime] NULL ,
	[xbody] [ntext] COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[XcatNew] (
	[orgCat] [smallint] NOT NULL ,
	[newCat] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[XdmpList] (
	[xdmpId] [nvarchar] (6) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xdmpName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[purpose] [nvarchar] (80) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[editDate] [smalldatetime] NOT NULL ,
	[editor] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[deptId] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[XfUpload] (
	[xfUpPath] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xfUpName] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xfDesc] [nvarchar] (300) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[xeditor] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CS_AS NOT NULL ,
	[xeditDate] [smalldatetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ZcustomDbinfo] (
	[sysId] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[itemId] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[itemName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[itemValue] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[updateTime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ZcustomDbstatus] (
	[sysId] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[sysName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[status] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
	[updateTime] [datetime] NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[AdRotate] WITH NOCHECK ADD 
	CONSTRAINT [PK_AdRotate] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Ap] WITH NOCHECK ADD 
	CONSTRAINT [PK_Ap] PRIMARY KEY  CLUSTERED 
	(
		[apcode]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Apcat] WITH NOCHECK ADD 
	CONSTRAINT [PK_Apcat] PRIMARY KEY  CLUSTERED 
	(
		[apcatId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AskOut] WITH NOCHECK ADD 
	CONSTRAINT [PK_AskOut] PRIMARY KEY  CLUSTERED 
	(
		[askOutIid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Banner] WITH NOCHECK ADD 
	CONSTRAINT [PK_Banner] PRIMARY KEY  CLUSTERED 
	(
		[bannerid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BaseDsd] WITH NOCHECK ADD 
	CONSTRAINT [PK_BaseDsd] PRIMARY KEY  CLUSTERED 
	(
		[ibaseDsd]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BaseDsdfield] WITH NOCHECK ADD 
	CONSTRAINT [PK_BaseDsdfield] PRIMARY KEY  CLUSTERED 
	(
		[ibaseField]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CatTreeNode] WITH NOCHECK ADD 
	CONSTRAINT [PK_CatTreeNode] PRIMARY KEY  CLUSTERED 
	(
		[ctNodeId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CatTreeRoot] WITH NOCHECK ADD 
	CONSTRAINT [PK_CatTreeRoot] PRIMARY KEY  CLUSTERED 
	(
		[ctRootId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Catalogue] WITH NOCHECK ADD 
	CONSTRAINT [PK_Catalogue] PRIMARY KEY  CLUSTERED 
	(
		[catId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm01] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm01] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm02] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm02] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm03] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm03] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm04] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm04] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm05] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm05] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm06] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm06] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm07] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm07] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm08] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm08] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm09] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm09] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm10] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm10] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm11] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm11] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm12] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm12] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cm13] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cm13] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CodeHt] WITH NOCHECK ADD 
	CONSTRAINT [PK_CodeHt] PRIMARY KEY  CLUSTERED 
	(
		[code],
		[item]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CodeMain] WITH NOCHECK ADD 
	CONSTRAINT [PK_CodeMain] PRIMARY KEY  CLUSTERED 
	(
		[codeMetaId],
		[mcode]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CodeMainLong] WITH NOCHECK ADD 
	CONSTRAINT [PK_CodeMainLong] PRIMARY KEY  CLUSTERED 
	(
		[codeMetaId],
		[mcode]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CodeMetaDef] WITH NOCHECK ADD 
	CONSTRAINT [PK_CodeMetaDef] PRIMARY KEY  CLUSTERED 
	(
		[codeId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Counter] WITH NOCHECK ADD 
	CONSTRAINT [PK_Counter] PRIMARY KEY  CLUSTERED 
	(
		[mp]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CtUnit] WITH NOCHECK ADD 
	CONSTRAINT [PK_CtUnit] PRIMARY KEY  CLUSTERED 
	(
		[ctUnitId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CtUserSet] WITH NOCHECK ADD 
	CONSTRAINT [PK_CtUserSet] PRIMARY KEY  CLUSTERED 
	(
		[ctNodeId],
		[userId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CtUserSet2] WITH NOCHECK ADD 
	CONSTRAINT [PK_CtUserSet2] PRIMARY KEY  CLUSTERED 
	(
		[ctNodeId],
		[userId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CtUserSet3] WITH NOCHECK ADD 
	CONSTRAINT [PK_CtUserSet3] PRIMARY KEY  CLUSTERED 
	(
		[ctNodeId],
		[userId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDTx9] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtx9] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtGeneric] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtgeneric] PRIMARY KEY  CLUSTERED 
	(
		[icuitem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtattach] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtattach] PRIMARY KEY  CLUSTERED 
	(
		[ixCuAttach]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtkeyword] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtkeyword] PRIMARY KEY  CLUSTERED 
	(
		[icuitem],
		[xkeyword]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtkeywordWtp] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtkeywordWtp] PRIMARY KEY  CLUSTERED 
	(
		[ikeyword]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtpage] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtpage] PRIMARY KEY  CLUSTERED 
	(
		[ixCuPage]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtwImg] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtwImg] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtx2] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtx2] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtx27] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtx27] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtx30] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtx30] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtx4] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtx4] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtx5] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtx5] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtx7] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtx7] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtx8] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtx8] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtxDformSpec] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtxDformSpec] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtxMmo] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtxMmo] PRIMARY KEY  CLUSTERED 
	(
		[icuitem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtxPaper] WITH NOCHECK ADD 
	CONSTRAINT [PK_CuDtxPaper] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[DataCat] WITH NOCHECK ADD 
	CONSTRAINT [PK_DataCat] PRIMARY KEY  CLUSTERED 
	(
		[catId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[DataContent] WITH NOCHECK ADD 
	CONSTRAINT [PK_DataContent] PRIMARY KEY  CLUSTERED 
	(
		[contentId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[DataUnit] WITH NOCHECK ADD 
	CONSTRAINT [PK_DataUnit] PRIMARY KEY  CLUSTERED 
	(
		[unitId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dept] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dept] PRIMARY KEY  CLUSTERED 
	(
		[deptId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[DeptOrg] WITH NOCHECK ADD 
	CONSTRAINT [PK_DeptOrg] PRIMARY KEY  CLUSTERED 
	(
		[deptId],
		[verId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[EpPub] WITH NOCHECK ADD 
	CONSTRAINT [PK_EpPub] PRIMARY KEY  CLUSTERED 
	(
		[epubId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[EpSend] WITH NOCHECK ADD 
	CONSTRAINT [PK_EpSend] PRIMARY KEY  CLUSTERED 
	(
		[ctRootId],
		[email],
		[epubId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Epaper] WITH NOCHECK ADD 
	CONSTRAINT [PK_Epaper] PRIMARY KEY  CLUSTERED 
	(
		[ctRootId],
		[email]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Forum] WITH NOCHECK ADD 
	CONSTRAINT [PK_Forum] PRIMARY KEY  CLUSTERED 
	(
		[forumId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ForumArticle] WITH NOCHECK ADD 
	CONSTRAINT [PK_ForumArticle] PRIMARY KEY  CLUSTERED 
	(
		[articleId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[GipSites] WITH NOCHECK ADD 
	CONSTRAINT [PK_GipSites] PRIMARY KEY  CLUSTERED 
	(
		[gipSiteId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[HfAptForm] WITH NOCHECK ADD 
	CONSTRAINT [PK_HfAptForm] PRIMARY KEY  CLUSTERED 
	(
		[hfAptId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[HfAptJob] WITH NOCHECK ADD 
	CONSTRAINT [PK_HfAptJob] PRIMARY KEY  CLUSTERED 
	(
		[hfAptJobId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[HfDeptSet] WITH NOCHECK ADD 
	CONSTRAINT [PK_HfDeptSet] PRIMARY KEY  CLUSTERED 
	(
		[deptId],
		[hyFormId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[HtDdb] WITH NOCHECK ADD 
	CONSTRAINT [PK_HtDdb] PRIMARY KEY  CLUSTERED 
	(
		[dbId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[HtDentity] WITH NOCHECK ADD 
	CONSTRAINT [PK_HtDentity] PRIMARY KEY  CLUSTERED 
	(
		[entityId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[HtDfield] WITH NOCHECK ADD 
	CONSTRAINT [PK_HtDfield] PRIMARY KEY  CLUSTERED 
	(
		[entityId],
		[xfieldName]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[InfoUser] WITH NOCHECK ADD 
	CONSTRAINT [PK_InfoUser] PRIMARY KEY  CLUSTERED 
	(
		[userId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[JobOrg] WITH NOCHECK ADD 
	CONSTRAINT [PK_JobOrg] PRIMARY KEY  CLUSTERED 
	(
		[jobId],
		[verId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ListConten] WITH NOCHECK ADD 
	CONSTRAINT [PK_ListConten] PRIMARY KEY  CLUSTERED 
	(
		[giCuItem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[M074] WITH NOCHECK ADD 
	CONSTRAINT [PK_M074] PRIMARY KEY  CLUSTERED 
	(
		[m074Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MemCatTree] WITH NOCHECK ADD 
	CONSTRAINT [PK_MemCatTree] PRIMARY KEY  CLUSTERED 
	(
		[ctNodeId],
		[memId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MemEpaper] WITH NOCHECK ADD 
	CONSTRAINT [PK_MemEpaper] PRIMARY KEY  CLUSTERED 
	(
		[ctNodeId],
		[memId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Member] WITH NOCHECK ADD 
	CONSTRAINT [PK_Member] PRIMARY KEY  CLUSTERED 
	(
		[account]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Mmofolder] WITH NOCHECK ADD 
	CONSTRAINT [PK_Mmofolder] PRIMARY KEY  CLUSTERED 
	(
		[mmofolderId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Mmosite] WITH NOCHECK ADD 
	CONSTRAINT [PK_Mmosite] PRIMARY KEY  CLUSTERED 
	(
		[mmositeId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[OrgVersion] WITH NOCHECK ADD 
	CONSTRAINT [PK_OrgVersion] PRIMARY KEY  CLUSTERED 
	(
		[verId],
		[verType]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Poll] WITH NOCHECK ADD 
	CONSTRAINT [PK_Poll] PRIMARY KEY  CLUSTERED 
	(
		[pollId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Polluser] WITH NOCHECK ADD 
	CONSTRAINT [PK_Polluser] PRIMARY KEY  CLUSTERED 
	(
		[sid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Rpt6] WITH NOCHECK ADD 
	CONSTRAINT [PK_Rpt6] PRIMARY KEY  CLUSTERED 
	(
		[rpt6id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Ugrp] WITH NOCHECK ADD 
	CONSTRAINT [PK_Ugrp] PRIMARY KEY  CLUSTERED 
	(
		[ugrpId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UgrpAp] WITH NOCHECK ADD 
	CONSTRAINT [PK_UgrpAp] PRIMARY KEY  CLUSTERED 
	(
		[apcode],
		[ugrpId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UgrpCode] WITH NOCHECK ADD 
	CONSTRAINT [PK_UgrpCode] PRIMARY KEY  CLUSTERED 
	(
		[codeId],
		[ugrpId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UpLoadSite] WITH NOCHECK ADD 
	CONSTRAINT [PK_UpLoadSite] PRIMARY KEY  CLUSTERED 
	(
		[upLoadSiteId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[WorkDay] WITH NOCHECK ADD 
	CONSTRAINT [PK_WorkDay] PRIMARY KEY  CLUSTERED 
	(
		[xday]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[XdmpList] WITH NOCHECK ADD 
	CONSTRAINT [PK_XdmpList] PRIMARY KEY  CLUSTERED 
	(
		[xdmpId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[XfUpload] WITH NOCHECK ADD 
	CONSTRAINT [PK_XfUpload] PRIMARY KEY  CLUSTERED 
	(
		[xfUpName],
		[xfUpPath]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BaseDsd] ADD 
	CONSTRAINT [DF__BaseDsd__inUse__24285DB4] DEFAULT ('Y') FOR [inUse]
GO

ALTER TABLE [dbo].[BaseDsdfield] ADD 
	CONSTRAINT [DF__BaseDsdfi__inUse__2704CA5F] DEFAULT ('Y') FOR [inUse],
	CONSTRAINT [DF__BaseDsdfi__xcanN__27F8EE98] DEFAULT ('Y') FOR [xcanNull],
	CONSTRAINT [DF__BaseDsdfi__xisPr__28ED12D1] DEFAULT ('N') FOR [xisPrimaryKey],
	CONSTRAINT [DF__BaseDsdfi__xiden__29E1370A] DEFAULT ('N') FOR [xidentity]
GO

ALTER TABLE [dbo].[CatTreeNode] ADD 
	CONSTRAINT [DF__CatTreeNo__ctNod__336AA144] DEFAULT ('C') FOR [ctNodeKind],
	CONSTRAINT [DF__CatTreeNo__dataL__345EC57D] DEFAULT (0) FOR [dataLevel],
	CONSTRAINT [DF__CatTreeNo__child__3552E9B6] DEFAULT (0) FOR [childCount],
	CONSTRAINT [DF__CatTreeNo__dataC__36470DEF] DEFAULT (0) FOR [dataCount],
	CONSTRAINT [DF__CatTreeNo__inUse__373B3228] DEFAULT ('Y') FOR [inUse],
	CONSTRAINT [DF__CatTreeNo__editD__382F5661] DEFAULT (getdate()) FOR [editDate]
GO

ALTER TABLE [dbo].[CatTreeRoot] ADD 
	CONSTRAINT [DF__CatTreeRo__editD__3B0BC30C] DEFAULT (getdate()) FOR [editDate]
GO

ALTER TABLE [dbo].[Catalogue] ADD 
	CONSTRAINT [DF__Catalogue__dataL__2CBDA3B5] DEFAULT (0) FOR [dataLevel],
	CONSTRAINT [DF__Catalogue__dataT__2DB1C7EE] DEFAULT (0) FOR [dataTop],
	CONSTRAINT [DF__Catalogue__dataP__2EA5EC27] DEFAULT (0) FOR [dataParent],
	CONSTRAINT [DF__Catalogue__child__2F9A1060] DEFAULT (0) FOR [childCount],
	CONSTRAINT [DF__Catalogue__dataC__308E3499] DEFAULT (0) FOR [dataCount]
GO

ALTER TABLE [dbo].[CodeMetaDef] ADD 
	CONSTRAINT [DF__CodeMetaD__codeT__5D60DB10] DEFAULT ('CodeMain') FOR [codeTblName],
	CONSTRAINT [DF__CodeMetaD__codeS__5E54FF49] DEFAULT ('codeMetaID') FOR [codeSrcFld],
	CONSTRAINT [DF__CodeMetaD__codeV__5F492382] DEFAULT ('mCode') FOR [codeValueFld],
	CONSTRAINT [DF__CodeMetaD__codeD__603D47BB] DEFAULT ('mValue') FOR [codeDisplayFld],
	CONSTRAINT [DF__CodeMetaD__codeS__61316BF4] DEFAULT ('mSortValue') FOR [codeSortFld],
	CONSTRAINT [DF__CodeMetaD__codeV__6225902D] DEFAULT ('代碼值') FOR [codeValueFldName],
	CONSTRAINT [DF__CodeMetaD__codeD__6319B466] DEFAULT ('顯示值') FOR [codeDisplayFldName],
	CONSTRAINT [DF__CodeMetaD__codeS__640DD89F] DEFAULT ('排序值') FOR [codeSortFldName],
	CONSTRAINT [DF__CodeMetaD__codeT__6501FCD8] DEFAULT ('EMB') FOR [codeType]
GO

ALTER TABLE [dbo].[Counter] ADD 
	CONSTRAINT [DF__Counter__counts__67DE6983] DEFAULT (0) FOR [counts]
GO

ALTER TABLE [dbo].[CtUnit] ADD 
	CONSTRAINT [DF__CtUnit__fctUnitO__6ABAD62E] DEFAULT ('Y') FOR [fctUnitOnly],
	CONSTRAINT [DF__CtUnit__inUse__6BAEFA67] DEFAULT ('Y') FOR [inUse],
	CONSTRAINT [DF__CtUnit__checkYn__6CA31EA0] DEFAULT ('N') FOR [checkYn]
GO

ALTER TABLE [dbo].[CtUserSet] ADD 
	CONSTRAINT [DF__CtUserSet__right__6F7F8B4B] DEFAULT (1) FOR [rights]
GO

ALTER TABLE [dbo].[CuDtGeneric] ADD 
	CONSTRAINT [DF__CuDtgener__fctup__7908F585] DEFAULT ('P') FOR [fctupublic],
	CONSTRAINT [DF__CuDtgener__dedit__79FD19BE] DEFAULT (getdate()) FOR [deditDate],
	CONSTRAINT [DF__CuDtgener__ximpo__7AF13DF7] DEFAULT (0) FOR [ximportant],
	CONSTRAINT [DF__CuDtgener__xnewW__7BE56230] DEFAULT ('N') FOR [xnewWindow],
	CONSTRAINT [DF__CuDtgener__xpost__7CD98669] DEFAULT (getdate()) FOR [xpostDate],
	CONSTRAINT [DF__CuDtgener__creat__7DCDAAA2] DEFAULT (getdate()) FOR [createdDate],
	CONSTRAINT [DF__CuDtgener__showT__7EC1CEDB] DEFAULT ('1') FOR [showType]
GO

ALTER TABLE [dbo].[CuDtattach] ADD 
	CONSTRAINT [DF__CuDtattac__blist__762C88DA] DEFAULT ('Y') FOR [blist]
GO

ALTER TABLE [dbo].[CuDtpage] ADD 
	CONSTRAINT [DF__CuDtpage__blist__056ECC6A] DEFAULT ('Y') FOR [blist]
GO

ALTER TABLE [dbo].[CuDtx2] ADD 
	CONSTRAINT [DF__CuDtx2__cuDtx2f9__0A338187] DEFAULT ('Y') FOR [cuDtx2f9]
GO

ALTER TABLE [dbo].[CuDtx3] ADD 
	CONSTRAINT [PK_CuDtx3] PRIMARY KEY  NONCLUSTERED 
	(
		[gicuitem]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CuDtxDformSpec] ADD 
	CONSTRAINT [DF__CuDtxDfor__expir__1C5231C2] DEFAULT (5) FOR [expireDayNum]
GO

ALTER TABLE [dbo].[Dept] ADD 
	CONSTRAINT [DF__Dept__kind__28B808A7] DEFAULT ('1') FOR [kind],
	CONSTRAINT [DF__Dept__inUse__29AC2CE0] DEFAULT ('Y') FOR [inUse],
	CONSTRAINT [DF__Dept__nodeKind__2AA05119] DEFAULT ('D') FOR [nodeKind]
GO

ALTER TABLE [dbo].[DeptOrg] ADD 
	CONSTRAINT [DF__DeptOrg__uniLeve__2D7CBDC4] DEFAULT (9) FOR [uniLevel],
	CONSTRAINT [DF__DeptOrg__uniRef__2E70E1FD] DEFAULT (0) FOR [uniRef],
	CONSTRAINT [DF__DeptOrg__staffCa__2F650636] DEFAULT (0) FOR [staffCat]
GO

ALTER TABLE [dbo].[EpSend] ADD 
	CONSTRAINT [DF__EpSend__sendDate__361203C5] DEFAULT (getdate()) FOR [sendDate]
GO

ALTER TABLE [dbo].[ForumArticle] ADD 
	CONSTRAINT [DF__ForumArti__postD__3BCADD1B] DEFAULT (getdate()) FOR [postDate],
	CONSTRAINT [DF__ForumArti__artic__3CBF0154] DEFAULT (0) FOR [articleLevel],
	CONSTRAINT [DF__ForumArti__artic__3DB3258D] DEFAULT (0) FOR [articleTop],
	CONSTRAINT [DF__ForumArti__artic__3EA749C6] DEFAULT (0) FOR [articleParent],
	CONSTRAINT [DF__ForumArti__artic__3F9B6DFF] DEFAULT (0) FOR [articleChildCount],
	CONSTRAINT [DF__ForumArti__child__408F9238] DEFAULT (0) FOR [childrenCount],
	CONSTRAINT [DF__ForumArti__readC__4183B671] DEFAULT (0) FOR [readCount]
GO

ALTER TABLE [dbo].[GipSites] ADD 
	CONSTRAINT [DF__GipSites__gipSit__45544755] DEFAULT (4) FOR [gipSiteInnerRoot],
	CONSTRAINT [DF__GipSites__gipSit__46486B8E] DEFAULT (3) FOR [gipSiteDefaultRoot]
GO

ALTER TABLE [dbo].[HfAptForm] ADD 
	CONSTRAINT [DF__HfAptForm__aplTy__4924D839] DEFAULT ('1') FOR [aplType],
	CONSTRAINT [DF__HfAptForm__aplDa__4A18FC72] DEFAULT (getdate()) FOR [aplDate],
	CONSTRAINT [DF__HfAptForm__statu__4B0D20AB] DEFAULT ('1') FOR [status],
	CONSTRAINT [DF__HfAptForm__lastD__4C0144E4] DEFAULT (getdate()) FOR [lastDate]
GO

ALTER TABLE [dbo].[HfAptJob] ADD 
	CONSTRAINT [DF__HfAptJob__hfJdat__4EDDB18F] DEFAULT (getdate()) FOR [hfJdate],
	CONSTRAINT [DF__HfAptJob__rtMail__4FD1D5C8] DEFAULT ('N') FOR [rtMail]
GO

ALTER TABLE [dbo].[HfDeptSet] ADD 
	CONSTRAINT [DF__HfDeptSet__editD__52AE4273] DEFAULT (getdate()) FOR [editDate]
GO

ALTER TABLE [dbo].[HtDfield] ADD 
	CONSTRAINT [DF__HtDfield__xcanNu__595B4002] DEFAULT ('Y') FOR [xcanNull],
	CONSTRAINT [DF__HtDfield__xisPri__5A4F643B] DEFAULT ('N') FOR [xisPrimaryKey],
	CONSTRAINT [DF__HtDfield__xident__5B438874] DEFAULT ('N') FOR [xidentity]
GO

ALTER TABLE [dbo].[InfoUser] ADD 
	CONSTRAINT [DF__InfoUser__userTy__5F141958] DEFAULT ('P') FOR [userType],
	CONSTRAINT [DF__InfoUser__ugrpId__60083D91] DEFAULT ('basic') FOR [ugrpId],
	CONSTRAINT [DF__InfoUser__visitC__60FC61CA] DEFAULT (0) FOR [visitCount]
GO

ALTER TABLE [dbo].[MemCatTree] ADD 
	CONSTRAINT [DF__MemCatTre__right__7226EDCC] DEFAULT (1) FOR [rights]
GO

ALTER TABLE [dbo].[MemEpaper] ADD 
	CONSTRAINT [DF__MemEpaper__right__75035A77] DEFAULT (1) FOR [rights]
GO

ALTER TABLE [dbo].[Ugrp] ADD 
	CONSTRAINT [DF__Ugrp__regdate__035179CE] DEFAULT (getdate()) FOR [regdate],
	CONSTRAINT [DF__Ugrp__isPublic__04459E07] DEFAULT ('Y') FOR [isPublic]
GO

ALTER TABLE [dbo].[WorkDay] ADD 
	CONSTRAINT [DF__WorkDay__offMark__0CDAE408] DEFAULT ('N') FOR [offMark]
GO

ALTER TABLE [dbo].[XdmpList] ADD 
	CONSTRAINT [DF__XdmpList__editDa__1387E197] DEFAULT (getdate()) FOR [editDate]
GO

ALTER TABLE [dbo].[XfUpload] ADD 
	CONSTRAINT [DF__XfUpload__xedito__16644E42] DEFAULT ('mof') FOR [xeditor],
	CONSTRAINT [DF__XfUpload__xeditD__1758727B] DEFAULT (getdate()) FOR [xeditDate]
GO

ALTER TABLE [dbo].[ZcustomDbinfo] ADD 
	CONSTRAINT [DF__ZcustomDb__updat__1940BAED] DEFAULT (getdate()) FOR [updateTime]
GO

ALTER TABLE [dbo].[ZcustomDbstatus] ADD 
	CONSTRAINT [DF__ZcustomDb__updat__1B29035F] DEFAULT (getdate()) FOR [updateTime]
GO

