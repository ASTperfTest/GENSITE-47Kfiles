using System;
using System.Web.Configuration;

public partial class FileTooLarge : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        HttpRuntimeSection runTime = (HttpRuntimeSection)WebConfigurationManager.GetSection("system.web/httpRuntime");

        double maxFileSize = Math.Round(runTime.MaxRequestLength / 1024.0, 1);
        LabelErrorMsg.Text = string.Format("上傳內容超過系統限制！請確認您的檔案在 {0:0.#} MB 以下.", maxFileSize);

    }
    protected void BackButton_Click(object sender, EventArgs e)
    {
        string url = Request.QueryString["url"];
        if (url != null)
        {
            Response.Redirect(url);
        }
       
    }
}
