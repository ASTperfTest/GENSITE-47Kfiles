using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Drawing.Imaging;

public partial class JpegImage : System.Web.UI.Page
{
    private Random random = new Random();
    protected void Page_Load(object sender, EventArgs e)
    {
        random = new Random();
        string temp = "";
        if (WebUtility.GetStringParameter("guid", string.Empty) == "")
        {
            Session["CaptchaImageText"] = GenerateRandomCode();
            temp = this.Session["CaptchaImageText"].ToString();
        }
        else
        {
            Session[WebUtility.GetStringParameter("guid", string.Empty)+"CaptchaImageText"] = GenerateRandomCode();
            temp = this.Session[WebUtility.GetStringParameter("guid", string.Empty) + "CaptchaImageText"].ToString();
        }
        // Create a CAPTCHA image using the text stored in the Session object.
        CaptchaImage ci = new CaptchaImage(temp, 200, 50, "Century Schoolbook");

        // Change the response headers to output a JPEG image.
        this.Response.Clear();
        this.Response.ContentType = "image/jpeg";

        // Write the image to the response stream in JPEG format.
        ci.Image.Save(this.Response.OutputStream, ImageFormat.Jpeg);

        // Dispose of the CAPTCHA image object.
        ci.Dispose();
    }

    private string GenerateRandomCode()
    {
        string s = "";
        for (int i = 0; i < 6; i++)
            s = String.Concat(s, this.random.Next(10).ToString());
        return s;
    }
}
