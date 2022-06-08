using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
public partial class GipEditML_InvalidUrlHead : System.Web.UI.Page
{

    //因為.4 主機沒對外所以無法執行URL動作，因此，將檢查程式放在.28，在.28中執行
    protected string callerURL = string.Empty;

    protected void Page_Load(object sender, EventArgs e)
    {
        callerURL = System.Web.Configuration.WebConfigurationManager.AppSettings["KMintra"];
        if (callerURL[callerURL.Length - 1] != '/' && callerURL[callerURL.Length - 1] != '\\') callerURL += "/";
        callerURL += @"coa/maUtility/CheckUrl/caller.aspx?";
        

        IRepository _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());

        var result = from p in _mGIPcoanew_repository.List<InvalidURLHeader>()
                     orderby p.runDate descending
                     select p;
        this.rptList.DataSource = result;
        this.rptList.DataBind();
    }

    protected void R1_ItemDataBound(Object Sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            int id = ((InvalidURLHeader)e.Item.DataItem).ID;
            Literal lit_Name = (Literal)e.Item.FindControl("lblProcess");
            lit_Name.Text = ProcessLiteral(id);
        }

    }
    protected void Repeater_OnItemCommand(object source, RepeaterCommandEventArgs e)
    {
        string id = e.CommandArgument.ToString();
        if (e.CommandName == "Select")
        {
            Response.Redirect("InvalidUrldetail.aspx?id=" + id );
        }
    }

    protected string ProcessLiteral(int id)
    {
        string ret = "";

        int totcount = 0;
        int nullcount = 0;
        int leftcount = 0;
        using (mGIPcoanewDataContext dc = new mGIPcoanewDataContext())
        {
            totcount = (from d in dc.InvalidURLDetail.AsEnumerable()
                        where d.ID == id
                        select d).Count();
            nullcount = (from d in dc.InvalidURLDetail.AsEnumerable()
                         where d.ID == id && d.State == null
                        select d).Count();
        }
        leftcount = totcount - nullcount;
        ret = leftcount.ToString() + "/" + totcount.ToString();
        if (leftcount < totcount)
        {
            ret = ret + "<a href='#' onclick='if(confirm(\"確定要重新執行嗎?\")){ react(" + id.ToString() + ")} else{return false;} '>(執行中，請重整頁面更新進度)</a>";
        }

        return ret;
    }
}
