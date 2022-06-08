using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

using GSS.Vitals.COA.Data;

public partial class kmintraSSO : System.Web.UI.Page
{
    string MemberId;
    string Script = "<script>alert('連線逾時或尚未登入，請登入會員');window.close();</script>";
    string kmintra = StrFunc.GetWebSetting("KMintra", "appSettings") + "/coa"; 

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["userID"] == null || string.IsNullOrEmpty(Session["userID"].ToString()))
        {
            Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Noid", Script);
            return;
        }
        else
        {
            MemberId = Session["userID"].ToString();
        }
    }
    protected void btnSearch_Click(object sender, EventArgs e)
    {
        string newGuid = Guid.NewGuid().ToString();
        string searchWord = HttpUtility.UrlEncodeUnicode(searchword.Value.Trim());
        string sql = "select count(*) from SSO where USER_ID = '" + MemberId + "'";
        if (Convert.ToInt16(SqlHelper.ReturnScalar("KMConnstring", sql)) == 0)
        {
            sql = "INSERT INTO SSO (GUID, USER_ID, CREATION_DATETIME) VALUES (@GUID, @USER_ID, getdate())";
            SqlHelper.ExecuteNonQuery("KMConnstring", sql,
                DbProviderFactories.CreateParameter("KMConnstring", "@GUID", "@GUID", newGuid),
                DbProviderFactories.CreateParameter("KMConnstring", "@USER_ID", "@USER_ID", MemberId));
        }
        else
        {
            sql = "select GUID from SSO where USER_ID = '" + MemberId + "'";
            newGuid = SqlHelper.ReturnScalar("KMConnstring", sql).ToString();

            sql = "update SSO set CREATION_DATETIME=getdate() where USER_ID=@USER_ID";
            SqlHelper.ExecuteNonQuery("KMConnstring", sql,
                DbProviderFactories.CreateParameter("KMConnstring", "@USER_ID", "@USER_ID", MemberId));
        }
        if (searchWord != "")
        {
            Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "javascript", "<script type='text/javascript'>location.href='" + kmintra + "/login.aspx?ssoid=" + newGuid + "&searchkey=" + HttpUtility.UrlEncode(searchWord) + "';</script>");
        }
        else
        {
            Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "javascript", "<script type='text/javascript'>location.href='" + kmintra + "/login.aspx?ssoid=" + newGuid + "';</script>");
        }

    }
}
