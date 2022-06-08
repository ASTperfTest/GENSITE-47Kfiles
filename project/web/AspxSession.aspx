<%@ Page language="c#" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>未命名頁面</title>
</head>
<script runat="server">
	private void Page_Load(object sender, System.EventArgs e) 
	{
		String redirectUrl = Request["redirecturl"];
		if (redirectUrl == null)
		{
			redirectUrl = "";
		}
		else
		{
			redirectUrl = redirectUrl.Replace(";", "&");
		}
		
    if (Request["type"] == "logout")
    {
      //Session.Clear();
      Session["memID"] = "";
      Session["memName"] = "";
      Session["gstyle"] = "";
      if (redirectUrl != "")
      {        				
                if (   redirectUrl.ToLower().Contains("MemberInvitePage.aspx".ToLower())
                    || redirectUrl.ToLower().Contains("recommand_".ToLower())
                )
                {
                    Response.Redirect("/");
                }
                
      
				if (redirectUrl.Contains("SubjectList.aspx") || redirectUrl.Contains("knowledge") || redirectUrl.Contains("Pedia")  || redirectUrl.Contains("Gardening"))
				{
					redirectUrl += "?type=logout";
					Response.Redirect(redirectUrl);
				}
				else
				{
					redirectUrl += "&type=logout";
					Response.Redirect(redirectUrl);
				}				
      }
			else
      {
        Response.Redirect("/");
      }
    }
    else
    {
      for (int i = 0; i < Request.Form.Count; i++)
      {
        Session[Request.Form.GetKey(i)] = Request.Form[i].ToString();
      }
      //Session.Timeout = 300;

      KPILogin kpilogin = new KPILogin(Session["memID"].ToString(), "loginInterCount", "");
      kpilogin.HandleLogin();
    			
      if (redirectUrl != "")
      {        
				Response.Redirect(redirectUrl);
      }
			else
      {
        Response.Redirect("/");
      }
		}
	}
</script>
<body></body>
</html>