using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Web.UI.HtmlControls;

public partial class MIP : System.Web.UI.Page
{
    string strURL = System.Web.Configuration.WebConfigurationManager.AppSettings["WWWUrl"].ToString();
    string InvitationCode = string.Empty;
    protected string joinMemberUrl = string.Empty;

    protected void Page_Load(object sender, EventArgs e)
    {
        setHTMLMeta();

        #region Fcode 執行 透過FaceBook加入會員
        if (Request.QueryString["Fcode"] != null &&
            !string.IsNullOrEmpty(Request.QueryString["Fcode"].ToString()))
        {
            InvitationCode = Request.QueryString["Fcode"].ToString();
            if (IsValidationCode(InvitationCode))
            {
                GotoJoinMember("Fcode", InvitationCode);
            }
            else
            {
                Response.Write(@"<script language='javascript' type='text/javascript'>alert('您的邀請碼已失效，您可能使用一般註冊程序來加入會員');
                                 location.href=('" + strURL + "');</script>");
            }
        }
        #endregion
        #region Ecode 執行 透過Email加入會員
        else if (Request.QueryString["Ecode"] != null &&
            !string.IsNullOrEmpty(Request.QueryString["Ecode"].ToString()))
        {
            InvitationCode = Request.QueryString["Ecode"].ToString();
            if (IsValidationCode(InvitationCode))
            {
                GotoJoinMember("Ecode", InvitationCode);
            }
            else
            {
                Response.Write(@"<script language='javascript' type='text/javascript'>alert('您的邀請碼已失效，您可能使用一般註冊程序來加入會員');
                                 location.href=('" + strURL + "');</script>");
            }
        }
        #endregion

        #region Ucode 執行 透過URL加入會員
        else if (Request.QueryString["Ucode"] != null &&
            !string.IsNullOrEmpty(Request.QueryString["Ucode"].ToString()))
        {
            InvitationCode = Request.QueryString["Ucode"].ToString();
            if (IsValidationCode(InvitationCode))
            {
                GotoJoinMember("Ucode", InvitationCode);
            }
            else
            {
                Response.Write(@"<script language='javascript' type='text/javascript'>alert('您的邀請碼已失效，您可能使用一般註冊程序來加入會員');
                                 location.href=('" + strURL + "');</script>");
            }
        }
        #endregion
        #region Scode 執行 分享至Facebook
        else if (Request.QueryString["Scode"] != null &&
            !string.IsNullOrEmpty(Request.QueryString["Scode"].ToString()))
        {
            InvitationCode = Request.QueryString["Scode"].ToString();
            if (IsValidationCode(InvitationCode))
            {
                GotoFB(InvitationCode);
            }
            else
            {
                Response.Write(@"<script language='javascript' type='text/javascript'>alert('您的邀請碼已失效');
                                 location.href=('" + strURL + "');</script>");
            }
        }
        #endregion
        else
        {
            // 不正確的進入方式
            Response.Write(@"<script language='javascript' type='text/javascript'>alert('返回入口網首頁');location.href=('" + strURL + "');</script>");
        }
    }

    private void setHTMLMeta()
    {
        // title
        //Page.Title = "農業知識入口網 －小知識串成的大力量";
        // Description
        HtmlMeta description = new HtmlMeta();
        description.Name = "description";
        description.Content = "朋友們，介紹你們一個好站，可以讓你輕鬆習得農業知識，掌握最新農業資訊，快來加入農業知識入口網吧～～";
        Page.Header.Controls.Add(description);
    }

    private bool IsValidationCode(string Code)
    {
        string strQueryScript = @"SELECT COUNT(*) FROM InviteFriends_Head WHERE InvitationCode = @InvitationCode";
        var count = SqlHelper.ReturnScalar("ODBCDSN", strQueryScript,
            DbProviderFactories.CreateParameter("ODBCDSN", "@InvitationCode", "@InvitationCode", Code));
        if ((int)count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }

    protected void GotoFB(string Code)
    {
        // http: //www.facebook.com/sharer/sharer.php?u=
        //              http: //kwpi-coa-kmweb.gss.com.tw/Member/MIP.aspx?Fcode=xxxxxx

        string sTemplate = "http://www.facebook.com/sharer/sharer.php?u={0}";
        string strPhysicalPath = System.IO.Path.GetFileName(Request.PhysicalPath);
        
        string shareURL = string.Format("{0}/Member/{1}?Fcode={2}", strURL, strPhysicalPath, Code);
        
        Response.Redirect(string.Format(sTemplate, shareURL));
    }

    protected void GotoJoinMember(string CodeType, string Code)
    {
        string sTemplate = "{0}/sp.asp?xdURL=coamember/member_Join.asp&mp=1&{1}={2}";
        joinMemberUrl = string.Format(sTemplate, strURL, CodeType, Code);
    }
}