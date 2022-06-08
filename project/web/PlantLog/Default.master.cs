using System;
using System.Web.UI;

public partial class _Default : MasterPage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string liStyle = "btnstyle05";

        DateTime now = DateTime.Now;
        DateTime voteFromDate = DateTime.Parse(WebUtility.GetAppSetting("VoteFromDate"));
        DateTime voteToDate = DateTime.Parse(WebUtility.GetAppSetting("VoteToDate"));
        DateTime uploadFromDate = DateTime.Parse(WebUtility.GetAppSetting("UploadFromDate"));
        DateTime uploadToDate = DateTime.Parse(WebUtility.GetAppSetting("UploadToDate"));

        string pic = "images/COA_PlantGrowth_01";
        string script = "<script type=\"text/javascript\" language=\"javascript\">\n$(document).ready(function() {\n";

        if (DateTime.Compare(now, uploadFromDate) > 0 && !(DateTime.Compare(now, uploadToDate) > 0))
        {
            script += "$(\"#coaTime\").html(\"" + ((uploadToDate - now).Days + 1) + "\");\n";
            pic += ".jpg";
        }
        else
        {
            if (DateTime.Compare(now, voteFromDate) > 0 && !(DateTime.Compare(now, voteToDate) > 0))
            {
                script += "$(\"#coaTime\").html(\"" + ((voteToDate - now).Days + 1) + "\");\n";
                pic += "-2.jpg";
                liStyle = "btnstyle03";
            }
            else
            {
                pic += "-3.jpg";
            }
        }

        script += "$(\"div#header\").css(\"background-image\", \"url('" + pic + "')\");\n";
        script += "$(\"#liVote\").addClass(\"" + liStyle + "\");\n});</script>";
        Page.ClientScript.RegisterStartupScript(typeof(Page), "coa", script);
    }
}
