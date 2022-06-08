using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Century_BirthdayLogin : System.Web.UI.Page
{
    string MemberId, atMonth, atDay;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request.QueryString["month"] == null || string.IsNullOrEmpty(Request.QueryString["month"].ToString()) ||
            Request.QueryString["day"] == null || string.IsNullOrEmpty(Request.QueryString["day"].ToString()) ||
            Request.QueryString["memberID"] == null || string.IsNullOrEmpty(Request.QueryString["memberID"].ToString()))
        {
            //Response.Write("<script>alert('無效連結，即將導回農業大事紀。');location.href('Event_List.aspx');</script>");
            Response.Write("<script>alert('無效連結，即將導回首頁。');location.href('../../mp.asp?mp=1');</script>");
        }
        else
        {
            atMonth = Request.QueryString["month"].ToString();
            atDay = Request.QueryString["day"].ToString();
            if (Request.QueryString["memberID"] != null && !string.IsNullOrEmpty(Request.QueryString["memberID"].ToString()))
            {
                Session["IsBirth"] = "Y";
            }
            Response.Redirect("History_List.aspx?month=" + atMonth + "&day=" + atDay);
        }
    }
}