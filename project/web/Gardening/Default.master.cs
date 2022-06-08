using System;
using System.Web.UI;

public partial class _Default : MasterPage
{
    protected string loginUrl = WebUtility.GetAppSetting("RedirectPage");
    protected void Page_Load(object sender, EventArgs e)
    {
		 if (Session["memID"] == null || Session["memID"].ToString() == "")
        {
            this.LoginLogout.Text = "登入";
			this.lblUser.Text="";
            this.LoginLogout.NavigateUrl = "javascript:Login();";
        }
        else
        {
            this.LoginLogout.Text = "登出";			
			if(Session["memName"] != null || Session["memName"].ToString() != "")
				this.lblUser.Text = Session["memName"].ToString();				
			if(Session["memNickName"] != null || Session["memNickName"].ToString() != "")
				this.lblUser.Text += " [ " + Session["memNickName"].ToString() + " ] ";	
			if(Session["memID"] != null || Session["memID"].ToString() != "")
				this.lblUser.Text += " [ " + Session["memID"].ToString() + " ] ";					
            this.LoginLogout.NavigateUrl = @"..\logout.asp?redirecturl=/Gardening/index.aspx";
        }
    }
}
