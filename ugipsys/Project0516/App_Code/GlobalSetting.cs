using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

/// <summary>
/// GlobalSetting 的摘要描述
/// </summary>
public class GlobalSetting
{
	public GlobalSetting()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}
    public static string ConnectionSettings()
    {
        return System.Configuration.ConfigurationManager.ConnectionStrings["ConnectionString"].ToString();

    }
}
