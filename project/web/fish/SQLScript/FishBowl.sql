SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fishbowl_account]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[fishbowl_account](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[login_id] [varchar](50) NOT NULL,
	[game_key] [varchar](50) NOT NULL,
	[realname] [nvarchar](200) NULL,
	[nickname] [nvarchar](200) NULL,
	[email] [nvarchar](200) NULL,
	[is_active] [tinyint] NOT NULL CONSTRAINT [DF_fishbowl_account_is_active]  DEFAULT ((1)),
	[create_time] [datetime] NOT NULL,
	[last_modify] [datetime] NULL,
 CONSTRAINT [PK_fishbowl_account] PRIMARY KEY CLUSTERED 
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
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fishbowl_pets_score]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[fishbowl_pets_score](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[fish_id] [int] NOT NULL,
	[health_score] [int] NOT NULL CONSTRAINT [DF_fishbowl_pets_score_health_score]  DEFAULT ((0)),
	[full_score] [int] NOT NULL CONSTRAINT [DF_fishbowl_pets_score_score]  DEFAULT ((0)),
	[water_score] [int] NOT NULL CONSTRAINT [DF_fishbowl_pets_score_water_score]  DEFAULT ((0)),
	[create_time] [datetime] NOT NULL,
 CONSTRAINT [PK_fishbowl_pets_score] PRIMARY KEY CLUSTERED 
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
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fishbowl_environment]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[fishbowl_environment](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[account_id] [int] NOT NULL,
	[water_status] [float] NOT NULL CONSTRAINT [DF_fishbowl_environment_water_status]  DEFAULT ((7)),
	[water_status_minus] [float] NOT NULL CONSTRAINT [DF_fishbowl_environment_water_status_minus]  DEFAULT ((0)),
	[money] [int] NOT NULL CONSTRAINT [DF_fishbowl_environment_money]  DEFAULT ((200)),
	[money_add] [int] NOT NULL CONSTRAINT [DF_fishbowl_environment_money_add]  DEFAULT ((0)),
	[use_uvlight] [datetime] NULL,
	[create_time] [datetime] NOT NULL,
	[last_modify] [datetime] NULL,
 CONSTRAINT [PK_fishbowl_bonus] PRIMARY KEY CLUSTERED 
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
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fishbowl_account_log]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[fishbowl_account_log](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[account_id] [int] NOT NULL,
	[action] [varchar](20) NOT NULL,
	[ip_address] [varchar](20) NOT NULL,
	[create_time] [datetime] NOT NULL,
 CONSTRAINT [PK_fishbowl_account_log] PRIMARY KEY CLUSTERED 
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
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fishbowl_pets]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[fishbowl_pets](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[account_id] [int] NOT NULL,
	[fish_type] [tinyint] NOT NULL CONSTRAINT [DF_fishbown_pets_fish_type]  DEFAULT ((0)),
	[health_status] [float] NOT NULL CONSTRAINT [DF_fishbown_pets_health_status]  DEFAULT ((10)),
	[health_status_minus] [float] NOT NULL CONSTRAINT [DF_fishbowl_pets_health_status_minus]  DEFAULT ((0)),
	[full_status] [float] NOT NULL CONSTRAINT [DF_fishbown_pets_full_status]  DEFAULT ((4)),
	[full_status_minus] [float] NOT NULL CONSTRAINT [DF_fishbowl_pets_full_status_minus]  DEFAULT ((0)),
	[is_active] [tinyint] NOT NULL CONSTRAINT [DF_fishbown_pets_is_active]  DEFAULT ((1)),
	[stars] [int] NOT NULL CONSTRAINT [DF_fishbowl_pets_stars]  DEFAULT ((0)),
	[create_time] [datetime] NULL,
	[last_modify] [datetime] NULL,
	[status] [tinyint] NOT NULL CONSTRAINT [DF_fishbowl_pets_status]  DEFAULT ((0)),
	[user_know] [tinyint] NOT NULL CONSTRAINT [DF_fishbowl_pets_user_know_fish_is_dead]  DEFAULT ((0)),
 CONSTRAINT [PK_fishbown_pets] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_fishbowl_pets_score_fishbowl_pets]') AND parent_object_id = OBJECT_ID(N'[dbo].[fishbowl_pets_score]'))
ALTER TABLE [dbo].[fishbowl_pets_score]  WITH CHECK ADD  CONSTRAINT [FK_fishbowl_pets_score_fishbowl_pets] FOREIGN KEY([fish_id])
REFERENCES [dbo].[fishbowl_pets] ([id])
ON UPDATE CASCADE
ON DELETE CASCADE
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_fishbowl_bonus_fishbowl_account]') AND parent_object_id = OBJECT_ID(N'[dbo].[fishbowl_environment]'))
ALTER TABLE [dbo].[fishbowl_environment]  WITH CHECK ADD  CONSTRAINT [FK_fishbowl_bonus_fishbowl_account] FOREIGN KEY([account_id])
REFERENCES [dbo].[fishbowl_account] ([id])
ON UPDATE CASCADE
ON DELETE CASCADE
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_fishbowl_account_log_fishbowl_account]') AND parent_object_id = OBJECT_ID(N'[dbo].[fishbowl_account_log]'))
ALTER TABLE [dbo].[fishbowl_account_log]  WITH CHECK ADD  CONSTRAINT [FK_fishbowl_account_log_fishbowl_account] FOREIGN KEY([account_id])
REFERENCES [dbo].[fishbowl_account] ([id])
ON UPDATE CASCADE
ON DELETE CASCADE
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_fishbown_pets_fishbowl_account]') AND parent_object_id = OBJECT_ID(N'[dbo].[fishbowl_pets]'))
ALTER TABLE [dbo].[fishbowl_pets]  WITH CHECK ADD  CONSTRAINT [FK_fishbown_pets_fishbowl_account] FOREIGN KEY([account_id])
REFERENCES [dbo].[fishbowl_account] ([id])
ON UPDATE CASCADE
ON DELETE CASCADE
