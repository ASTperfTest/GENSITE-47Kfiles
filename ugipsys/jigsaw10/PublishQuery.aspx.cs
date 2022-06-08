using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class PublishQuery : System.Web.UI.Page
{
    public int pid;
    protected void Page_Load(object sender, EventArgs e)
    {
        // QueryString
        pid = int.Parse(Request.QueryString["pid"] ?? "0");
        int id = int.Parse(Request.QueryString["id"] ?? "0");
        string topCat = Request.QueryString["topCat"] ?? "";

        if (IsPostBack)
        {
            Server.Transfer("PublishAdd.aspx");
        }
        else
        {
            // LINQ to SQL (Repository)
            IRepository _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());

            var rs = _mGIPcoanew_repository.Get<CuDTGeneric>(p => p.iCUItem == pid );
            var rs1 = _mGIPcoanew_repository.Get<CuDTGeneric>(p => p.iCUItem == id);

            if (rs != null && rs1 != null)
                titles.Text = string.Format("【內容條例清單--{0} {1}】", rs.sTitle, rs1.sTitle);

            if (topCat == "A")
            {
                var result = from p in _mGIPcoanew_repository.List<CodeMain>()
                             where p.codeMetaID == "jigsaw"
                             orderby p.mSortValue
                             select new
                             {
                                 p.mCode,
                                 p.mValue,
                             };

                CtRootId.DataSource = result;
                CtRootId.DataTextField = "mValue";
                CtRootId.DataValueField = "mCode";
                CtRootId.DataBind();
            }
        }
    }
}
