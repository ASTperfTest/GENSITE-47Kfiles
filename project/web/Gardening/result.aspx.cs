using System;
using System.Web.UI;

public partial class result : Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        SetViewState();

        if (bool.Parse((string)ViewState["hasLogin"]))
        {
            if (bool.Parse((string)ViewState["isBeforeVote"]))
            {
                Response.Write("<script language='javascript'>alert('活動尚未開始，敬請期待！');location.href='"+WebUtility.GetAppSetting("WWWUrl")+"';</script>");
            }

            if (!(bool.Parse((string)ViewState["isBeforeVote"]) || bool.Parse((string)ViewState["isAfterVote"])))
            {
                Response.Write("<script language='javascript'>alert('投票尚未結束，敬請期待！');history.back();</script>");
            }

            if (!(bool.Parse((string)ViewState["isBeforeUpload"]) || bool.Parse((string)ViewState["isAfterUpload"])))
            {
                Response.Write("<script language='javascript'>alert('活動尚未結束，敬請期待！');history.back();</script>");
            }
        }
        else
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "MyScript",
                "alert('請先登入會員');setHref('" + WebUtility.GetAppSetting("RedirectPage") + "');", true);
        }
    }

    private void SetViewState()
    {
        if (ViewState["hasLogin"] == null)
        {
            ViewState["hasLogin"] = WebUtility.CheckLogin(Session["memID"]).ToString();
        }

        DateTime now = DateTime.Now;
        DateTime voteFromDate = DateTime.Parse(WebUtility.GetAppSetting("VoteFromDate"));
        DateTime voteToDate = DateTime.Parse(WebUtility.GetAppSetting("VoteToDate"));
        DateTime uploadFromDate = DateTime.Parse(WebUtility.GetAppSetting("UploadFromDate"));
        DateTime uploadToDate = DateTime.Parse(WebUtility.GetAppSetting("UploadToDate"));

        if (ViewState["isBeforeUpload"] == null || ViewState["isAfterUpload"] == null)
        {
            if (DateTime.Compare(now, uploadFromDate) > 0)
            {
                ViewState["isBeforeUpload"] = false.ToString();
            }
            else
            {
                ViewState["isBeforeUpload"] = true.ToString();
            }

            if (DateTime.Compare(now, uploadToDate) > 0)
            {
                ViewState["isAfterUpload"] = true.ToString();
            }
            else
            {
                ViewState["isAfterUpload"] = false.ToString();
            }
        }

        if (ViewState["isBeforeVote"] == null || ViewState["isAfterVote"] == null)
        {
            if (DateTime.Compare(now, voteFromDate) > 0)
            {
                ViewState["isBeforeVote"] = false.ToString();
            }
            else
            {
                ViewState["isBeforeVote"] = true.ToString();
            }

            if (DateTime.Compare(now, voteToDate) > 0)
            {
                ViewState["isAfterVote"] = true.ToString();
            }
            else
            {
                ViewState["isAfterVote"] = false.ToString();
            }
        }
    }
}
