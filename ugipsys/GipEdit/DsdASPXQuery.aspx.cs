using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using GSS.Vitals.COA.Data;

public partial class DsdASPXQuery : System.Web.UI.Page
{
    List<String> msg = new List<String>();
    private string iCTUnit, refID, MemberID, fileName;
    string Script = "<script>alert('�s�u�O�ɩΩ|���n�J�A�Э��s�n�J');window.close();</script>";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["userID"] == null || string.IsNullOrEmpty(Session["userID"].ToString()) ||
            Session["ctNodeId"] == null || string.IsNullOrEmpty(Session["ctNodeId"].ToString()))
        {
            Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Noid", Script);
            return;
        }
    }

    protected void btnQuery_Click(object sender, EventArgs e)
    {
        string strURL = "DsdASPXList.aspx?ItemID=" + Session["itemID"].ToString() +
                        "&CtNodeID=" + Session["ctNodeId"].ToString() + "&from=query" + "&title=" + titleTxt.Text;
        Response.Write("<script language='javascript'>location.href('"+strURL+"');</script>");	
    }

}