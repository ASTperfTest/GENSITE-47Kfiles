using System;

public partial class BlockControl : System.Web.UI.UserControl
{
    protected void Page_Load(object sender, EventArgs e)
    {
       
    }

    protected void Block1Up_Click(object sender, EventArgs e)
    {        
        Session["Index"] = Block1Num.Text;
    }
    protected void Block1Down_Click(object sender, EventArgs e)
    {
        Session["Index"] = Block1Num.Text;
    }
}
