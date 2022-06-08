using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.IO;
using System.Xml;
using System.Data.SqlClient;
using System.Collections;
using System.Data;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Service;

public partial class knowledge_rediect : System.Web.UI.Page
{
    protected string url = System.Web.Configuration.WebConfigurationManager.AppSettings["myURL"];
    protected String knowledgeUrl ="";
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["memID"] == null)
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "MyScript",
                "alert('請先登入會員');setHref('" + WebUtility.GetAppSetting("RedirectPage") + "');", true);
        }
        else
        {
            Response.Redirect("../knowledge/knowledge_question.aspx?gardening=true");
        }

    }
}
