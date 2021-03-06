if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_userinfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_userinfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserLoginInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserLoginInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[userActionLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[userActionLog]
GO

CREATE TABLE [dbo].[UserLoginInfo] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[userid] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[logintime] [datetime] NULL ,
	[loginip] [varchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[act] [varchar] (5000) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[userActionLog] (
	[logTaskID] [int] IDENTITY (1, 1) NOT NULL ,
	[loginSID] [int] NOT NULL ,
	[xTarget] [varchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[xAction] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[recordNumber] [int] NULL ,
	[objTitle] [varchar] (800) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[actionRecord] [varchar] (4000) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[actiontime] [datetime] NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[UserLoginInfo] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserLoginInfo] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[userActionLog] WITH NOCHECK ADD 
	CONSTRAINT [DF_userActionLog_actiontime] DEFAULT (getdate()) FOR [actiontime],
	CONSTRAINT [PK_userActionLog] PRIMARY KEY  CLUSTERED 
	(
		[logTaskID]
	)  ON [PRIMARY] 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


Create procedure sp_userinfo( 
         @vuserid    char(10),
         @vstartdate      char(10),
         @venddate   char(10)   
  )
as
   declare  @icnt int,@sql varchar(5000)
   set nocount on

   create table #temp
	  (xid                        int,
                   UserName   	varchar(20),
	   UserID	                varchar(10),
	   loginip		varchar(50),
	   logintime	datetime,
                   act                         varchar(50),
                   xtarget                   varchar(50),
                   ctunit                     varchar(200),
                   objtitle                   varchar(800),
                   cnt                         int )

   if @vuserid <>''
   begin
           if @vstartdate <>'' and @venddate <>'' 
           begin
                    select @sql=' insert into #temp (xid,userid,loginip,logintime,act)
                                        select id,userid,loginip,logintime,act from userlogininfo 
                                        where userid=''' +@vuserid + ''' and logintime between ''' + @vstartdate + ''' and ''' + @venddate+ ''''
                    exec (@sql)

           end
           else
           begin    
                    select @sql=' insert into #temp (xid,userid,loginip,logintime,act)
                                        select id,userid,loginip,logintime,act from userlogininfo 
                                        where userid=''' +@vuserid + ''''
                    exec (@sql)
           end 
   end
   else
   begin
           if @vstartdate <>'' and @venddate <>'' 
           begin
                    select @sql=' insert into #temp (xid,userid,loginip,logintime,act)
                                        select id,userid,loginip,logintime,act from userlogininfo 
                                        where logintime between ''' + @vstartdate + ''' and ''' + @venddate + ''''
                    exec (@sql)
           end
           else
           begin    
                    select @sql=' insert into #temp (xid,userid,loginip,logintime,act)
                                        select id,userid,loginip,logintime,act from userlogininfo'
                    exec (@sql)
           end 
   end

    insert into #temp (xid,userid,loginip,logintime,act,xtarget,ctunit,objtitle)
    select a.xid,a.userid,a.loginip,b.actiontime,b.xaction,b.xtarget,b.recordnumber,b.objtitle
    from #temp a,userActionlog b where a.xid=b.loginSID

    update #temp 
    set username = b.username
    from #temp a, infouser b
    where a.userid=b.userid

    update #temp 
    set xtarget = '資料上稿--'
    where xtarget='0A'

    update #temp 
    set act = '新增資料'
    where act='1'

    update #temp 
    set act = '修改資料'
    where act='2'

    update #temp 
    set act = '刪除資料'
    where act='3' 

    update #temp
    set ctunit=c.catname
    from #temp a,cudtgeneric b,cattreenode c
    where a.ctunit=b.icuitem and b.ictunit=c.ctunitid

    select @icnt=count(*) from #temp
    update #temp set cnt=@icnt 

   select * from #temp  order by xid desc

   drop table #temp 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

