﻿using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

/// <summary>
/// SqlDbManager 的摘要描述
/// </summary>
public class SqlDbManager
{
	public static string CONNECTION_STRING = System.Configuration.ConfigurationManager.ConnectionStrings["COA_GIPConnectionString"].ConnectionString;

	public SqlDbManager()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}
}
