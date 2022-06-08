if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Epaper2]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Epaper2]
GO

CREATE TABLE [dbo].[Epaper2] (
	[email] [char] (60) COLLATE Chinese_Taiwan_Stroke_CI_AS NOT NULL ,
	[createtime] [datetime] NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Epaper2] WITH NOCHECK ADD 
	CONSTRAINT [PK_Epaper2] PRIMARY KEY  CLUSTERED 
	(
		[email]
	)  ON [PRIMARY] 
GO


Insert Into ePaper2(email) values('frank@hyweb.com.tw');


Select * from ePaper2

