CREATE TABLE [dbo].[Dept] (
	[giCuItem] [int] NULL ,
	[deptID] [varchar] (10) NOT NULL ,
	[deptName] [varchar] (30) NOT NULL ,
	[AbbrName] [varchar] (16) NULL ,
	[eDeptName] [varchar] (60) NULL ,
	[eAbbrName] [varchar] (20) NULL ,
	[OrgRank] [char] (1) NULL ,
	[kind] [char] (1) NOT NULL ,
	[DeptHead] [varchar] (10) NULL ,
	[inUse] [char] (1) NOT NULL ,
	[tDataCat] [varchar] (20) NULL ,
	[servPhone] [varchar] (20) NULL ,
	[servAddr] [varchar] (80) NULL ,
	[servHour] [varchar] (80) NULL ,
	[eservAddr] [varchar] (80) NULL ,
	[eservHour] [varchar] (80) NULL ,
	[parent] [varchar] (10) NULL ,
	[nodeKind] [char] (1) NOT NULL ,
	[pLevel] [smallint] NULL ,
	[seq] [smallint] NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dept] WITH NOCHECK ADD 
	CONSTRAINT [DF_Dept_kind] DEFAULT ('1') FOR [kind],
	CONSTRAINT [DF_Dept_inUser] DEFAULT ('Y') FOR [inUse],
	CONSTRAINT [DF_Dept_nodeKind] DEFAULT ('D') FOR [nodeKind],
	CONSTRAINT [PK_Dept] PRIMARY KEY  NONCLUSTERED 
	(
		[deptID]
	)  ON [PRIMARY] 
GO

