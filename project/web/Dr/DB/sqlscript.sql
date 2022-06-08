SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[farm2009_account]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[farm2009_account](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[login_id] [varchar](50) NOT NULL,
	[game_key] [varchar](50) NOT NULL,
	[realname] [nvarchar](200) NULL,
	[nickname] [nvarchar](200) NULL,
	[email] [nvarchar](200) NULL,
	[create_time] [datetime] NOT NULL,
	[last_modify] [datetime] NULL,
 CONSTRAINT [PK_farm2009_account] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[farm2009_questions]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[farm2009_questions](
	[id] [int] IDENTITY(151,1) NOT NULL,
	[qrank] [tinyint] NOT NULL CONSTRAINT [DF_QA_qrank]  DEFAULT ((0)),
	[qcontent] [nvarchar](150) NOT NULL CONSTRAINT [DF_QA_qcontent]  DEFAULT ('無資料'),
	[a1] [nvarchar](100) NOT NULL CONSTRAINT [DF_QA_a1]  DEFAULT ('無資料'),
	[a2] [nvarchar](100) NOT NULL CONSTRAINT [DF_QA_a2]  DEFAULT ('無資料'),
	[a3] [nvarchar](100) NOT NULL CONSTRAINT [DF_QA_a3]  DEFAULT ('無資料'),
	[correct] [tinyint] NOT NULL CONSTRAINT [DF_QA_correct]  DEFAULT ((0)),
	[disable] [tinyint] NOT NULL CONSTRAINT [DF_QA_disable]  DEFAULT ((0)),
 CONSTRAINT [PK_farm2009_questions] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[farm2009_account_log]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[farm2009_account_log](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[account_id] [int] NOT NULL,
	[action] [varchar](30) NOT NULL,
	[ip_address] [varchar](15) NOT NULL,
	[create_time] [datetime] NOT NULL,
 CONSTRAINT [PK_farm2009_account_log] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[farm2009_question_data]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[farm2009_question_data](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[account_id] [int] NOT NULL,
	[current_level] [tinyint] NOT NULL CONSTRAINT [DF_farm2009_question_data_current_level]  DEFAULT ((0)),
	[current_score] [int] NOT NULL CONSTRAINT [DF_farm2009_question_data_current_score]  DEFAULT ((0)),
	[current_timer] [bigint] NOT NULL CONSTRAINT [DF_farm2009_question_data_current_timer]  DEFAULT ((0)),
	[current_round] [int] NOT NULL CONSTRAINT [DF_farm2009_question_data_current_round]  DEFAULT ((1)),
	[create_time] [datetime] NOT NULL,
 CONSTRAINT [PK_farm2009_question_data] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'' ,@level0type=N'SCHEMA', @level0name=N'dbo', @level1type=N'TABLE', @level1name=N'farm2009_question_data'

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[farm2009_question_log]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[farm2009_question_log](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[data_id] [int] NOT NULL,
	[question_id] [int] NOT NULL,
	[question_order] [tinyint] NOT NULL CONSTRAINT [DF_farm2009_question_log_question_order]  DEFAULT ((0)),
	[is_correct] [tinyint] NOT NULL CONSTRAINT [DF_farm2009_question_log_is_correct]  DEFAULT ((0)),
	[create_time] [datetime] NOT NULL,
 CONSTRAINT [PK_farm2009_question_log] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_farm2009_account_log_farm2009_account]') AND parent_object_id = OBJECT_ID(N'[dbo].[farm2009_account_log]'))
ALTER TABLE [dbo].[farm2009_account_log]  WITH CHECK ADD  CONSTRAINT [FK_farm2009_account_log_farm2009_account] FOREIGN KEY([account_id])
REFERENCES [dbo].[farm2009_account] ([id])
ON UPDATE CASCADE
ON DELETE CASCADE
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_farm2009_question_data_farm2009_account]') AND parent_object_id = OBJECT_ID(N'[dbo].[farm2009_question_data]'))
ALTER TABLE [dbo].[farm2009_question_data]  WITH CHECK ADD  CONSTRAINT [FK_farm2009_question_data_farm2009_account] FOREIGN KEY([account_id])
REFERENCES [dbo].[farm2009_account] ([id])
ON UPDATE CASCADE
ON DELETE CASCADE
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_farm2009_question_log_farm2009_questions]') AND parent_object_id = OBJECT_ID(N'[dbo].[farm2009_question_log]'))
ALTER TABLE [dbo].[farm2009_question_log]  WITH CHECK ADD  CONSTRAINT [FK_farm2009_question_log_farm2009_questions] FOREIGN KEY([question_id])
REFERENCES [dbo].[farm2009_questions] ([id])
ON UPDATE CASCADE
ON DELETE CASCADE
