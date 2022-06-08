using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.Linq;
using System.Data.Linq.Mapping;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
public partial class GipEditML_InvalidUrldetail : System.Web.UI.Page
{
    public int sid;
    protected PaginatedList<InvalidURLDetail> pl;
    protected bool IsDataRemoved;
    IRepository _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());
    DataContext db = new mGIPcoanewDataContext();

    protected void Page_Load(object sender, EventArgs e)
    {
        sid = int.Parse(Request["id"].ToString());


        //==================================================
        //檢查URL移除了沒
        var removeDate = (from rDate in _mGIPcoanew_repository.List<InvalidURLHeader>()
                          where rDate.ID == sid
                          select rDate.removeDate).FirstOrDefault();

        IsDataRemoved = removeDate != null;
        //資料必須全部檢查完成 才可以執行刪除的動作
        if (!IsDataRemoved)
        {
            string sql = @"select count(*) from InvalidURLDetail where [state] is null and id={0}";
            using (var db = new mGIPcoanewDataContext())
            {
                IsDataRemoved = db.ExecuteQuery<int>(sql, sid).FirstOrDefault() != 0;
                
            }
        }



        if (IsDataRemoved)
        {
            btclose.Visible = false;
            btremove.Visible = false;
            removeHead.Visible = false;
        }
        //==================================================
        

        var result = (from p in _mGIPcoanew_repository.List<InvalidURLDetail>()
                     join doc in _mGIPcoanew_repository.List<CuDTGeneric>() on p.ArticleId equals doc.iCUItem
                     where p.ID == sid && p.State != 'V'
                     orderby p.Result
                     select p);
        //分頁
        int page = int.Parse(Request.QueryString["page"] ?? "0");
        int pageSize = int.Parse(Request.QueryString["pagesize"] ?? "5000");
        pl = new PaginatedList<InvalidURLDetail>(result, page, pageSize, new string[] { "0", "10", "15", "30", "50", "5000" });

        var pp = from p in pl
                 select p;
        //Data Repeater DataBinding
        this.rptList.DataSource = pp;
        this.rptList.DataBind();

    }

    protected void R1_ItemDataBound(Object Sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            int id = ((InvalidURLDetail)e.Item.DataItem).ArticleId;
            string s = ((InvalidURLDetail)e.Item.DataItem).State.ToString();
            
            var cudt = (from p in _mGIPcoanew_repository.List<CuDTGeneric>()
                        join u in _mGIPcoanew_repository.List<BaseDSD>() on p.iBaseDSD equals u.iBaseDSD into up
                        join w in _mGIPcoanew_repository.List<CtUnit>() on p.iCTUnit equals w.CtUnitID into wp
                        from u in up.DefaultIfEmpty()
                        from w in wp.DefaultIfEmpty()
                        where p.iCUItem == id
                        select new
                        {
                            p.sTitle,
                            p.iCTUnit,
                            p.xURL,
                            u.sBaseDSDName,
                            w.CtUnitName
                        }).FirstOrDefault();
            if (cudt == null) return;

            var cnode = (from n in _mGIPcoanew_repository.List<CatTreeNode>()
                         join r in _mGIPcoanew_repository.List<CatTreeRoot>() on n.CtRootID equals r.CtRootID into nr
                         from r in nr.DefaultIfEmpty()
                         where n.CtUnitID == cudt.iCTUnit
                         select new
                         {
                             r.CtRootName
                         }).FirstOrDefault();
            if (cnode == null) return;

            Literal lnk = (Literal)e.Item.FindControl("lview");
            lnk.Text = "<a href='" + cudt.xURL + "' target='_blank'>View</a>";
            Label lbstatus = (Label)e.Item.FindControl("lstatus");
            if (s == "N")
            {
                lbstatus.Text = "未審核";
            }
            else if (s == "D")
            {
                lbstatus.Text = "已移除";
            }
            else if (s == null)
            {
                lbstatus.Text = "尚末排程檢查";
            }

            Label lsource = (Label)e.Item.FindControl("lsource");
            lsource.Text = cnode.CtRootName;
            Label lunit = (Label)e.Item.FindControl("lunit");
            lunit.Text = cudt.CtUnitName;
            Label ltitle = (Label)e.Item.FindControl("ltitle");
            ltitle.Text = cudt.sTitle;


            string sql = @"
            select
	            CTUnitId
            from(
            select CTUnitId from KnowledgeToSubject where status='Y'
            union
            select ctUnitId from CtUnit where (ctUnitName like '%新聞%' or ctUnitName like '%最新%' or ctUnitName like '%消息%')
            ) as T
            where ctUnitId={0}";

            try
            {
                Label linkType = (Label)e.Item.FindControl("linkType");
                if (db.ExecuteQuery<object>(sql, cudt.iCTUnit).Count() > 0)
                {
                    linkType.Text = "新聞";
                }
                else
                {
                    linkType.Text = "資料";
                }
            }
            catch { }
            

            
            
        }
    }

    protected void Remove_select(Object Sender, EventArgs e)
    {
        for (int i = 0; i < rptList.Items.Count; i++)
        {
            CheckBox cbx = (CheckBox)rptList.Items[i].FindControl("CheckSelect");
            TextBox tbx = (TextBox)rptList.Items[i].FindControl("tbxid");
            TextBox tbxacid = (TextBox)rptList.Items[i].FindControl("tbxacid");
            int rid = int.Parse(tbx.Text);
            int aid = int.Parse(tbxacid.Text);
            if (cbx != null)
            {
                if (cbx.Checked == true)
                {
                    IRepository _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());
                    var originInvalidUrldetail = _mGIPcoanew_repository.Get<InvalidURLDetail>(p => p.ID == rid && p.ArticleId == aid);
                    originInvalidUrldetail.State = 'V';
                    _mGIPcoanew_repository.Save();
                }
            }
        }

        Response.Redirect("InvalidUrldetail.aspx?id=" + sid);
    }


    protected void RemoveHead(Object Sender, EventArgs e)
    {
        using (var db = new mGIPcoanewDataContext())
        { 
            string sql = @"
            begin tran
	            delete from InvalidURLDetail where id = {0}
	            delete from InvalidURLHeader where id = {0} 
            commit    ";

            db.ExecuteCommand(sql, this.sid);
        }        
        
        Response.Redirect("InvalidUrlHead.aspx", true);
    }


    protected void Close_all(Object Sender, EventArgs e)
    {
        sid = int.Parse(Request["id"].ToString());
        string sqlstr = @"
            begin tran
                
                DECLARE @NewLineChar nvarchar(2) 
                set @NewLineChar = CHAR(13) + CHAR(10)
                declare @tmp table (unitId int)

                insert into @tmp
                select CTUnitId from KnowledgeToSubject where status='Y'
                union
                select ctUnitId from CtUnit where (ctUnitName like '%新聞%' or ctUnitName like '%最新%' or ctUnitName like '%消息%')

                --新聞資料：資料不下架但將URL清空，URL備份到Body
                update  CuDTGeneric    
	                set  CuDTGeneric.xBody = cast(isnull(CuDTGeneric.xBody,'') as nvarchar(max)) + @NewLineChar + '<!--無效聯結：' + CuDTGeneric.xURL + '-->' 
                        ,CuDTGeneric.xURL = ''
                from InvalidURLDetail
                inner join CuDTGeneric on CuDTGeneric.iCuitem = InvalidURLDetail.articleId 
                where 
                    InvalidURLDetail.[Id] = {0}
                and InvalidURLDetail.[State] = 'N'
                and CuDTGeneric.showType in (2, 3)
                and CuDTGeneric.iCTUnit in 
                (
	                select unitId from @tmp
                )

                --非新聞資料：將資料下架
                update  CuDTGeneric
	                set CuDTGeneric.fCTUPublic='N'
                        ,CuDTGeneric.xBody = cast(isnull(CuDTGeneric.xBody,'') as nvarchar(max)) + @NewLineChar + '<!--無效聯結：' + CuDTGeneric.xURL + '-->' 
                        ,CuDTGeneric.xURL = ''
                from InvalidURLDetail
                inner join CuDTGeneric on CuDTGeneric.iCuitem = InvalidURLDetail.articleId 
                where 
                    InvalidURLDetail.[Id] = {0}
                and InvalidURLDetail.[State] = 'N'
                and CuDTGeneric.showType in (2, 3)
                and CuDTGeneric.iCTUnit not in 
                (
                    select unitId from @tmp
                )

                update  InvalidURLDetail
	                set InvalidURLDetail.[State] = 'D'
                from InvalidURLDetail
                inner join CuDTGeneric on CuDTGeneric.iCuitem = InvalidURLDetail.articleId 
                where 
                    InvalidURLDetail.[Id] = {0}
                and InvalidURLDetail.[State] = 'N'
                and CuDTGeneric.showType in (2, 3)

                update InvalidURLHeader set removeDate = getdate() where [id] = {0}

            commit
            ";
        

        using (var db = new mGIPcoanewDataContext())
        {
            db.ExecuteCommand(sqlstr, sid);
        }


        Response.Redirect("InvalidUrlHead.aspx", true);
    }

    protected void ExportToExcel(object sender, EventArgs e)
    {
        sid = int.Parse(Request["id"].ToString());
        IRepository _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());

        var result = from p in _mGIPcoanew_repository.List<InvalidURLDetail>()
                     where p.ID == sid && p.State != 'V'
                     select p;
        gridview1.DataSource = result;
        gridview1.DataBind();
        System.IO.StringWriter tw = new System.IO.StringWriter();
        System.Web.UI.HtmlTextWriter hw =
           new System.Web.UI.HtmlTextWriter(tw);
        hw.WriteLine("<b>" +
           "匯出日期：" + DateTime.Now.ToShortDateString() +
           "</b>");
        hw.WriteLine("<br>&nbsp;");
        // Get the HTML for the control.
        gridview1.HeaderStyle.Font.Bold = true;
        gridview1.RenderControl(hw);

        // Write the HTML back to the browser.
        Response.ContentType = "application/vnd.ms-excel";
        this.EnableViewState = false;
        Response.Write(tw.ToString());
        Response.End();
    }
    public override void VerifyRenderingInServerForm(Control control)
    {
    }

    protected void GV_DataBound(Object Sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            int id = ((InvalidURLDetail)e.Row.DataItem).ArticleId;
            string s = ((InvalidURLDetail)e.Row.DataItem).State.ToString();
            IRepository _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());

            var cudt = (from p in _mGIPcoanew_repository.List<CuDTGeneric>()
                        join u in _mGIPcoanew_repository.List<BaseDSD>() on p.iBaseDSD equals u.iBaseDSD into up
                        join w in _mGIPcoanew_repository.List<CtUnit>() on p.iCTUnit equals w.CtUnitID into wp
                        from u in up.DefaultIfEmpty()
                        from w in wp.DefaultIfEmpty()
                        where p.iCUItem == id
                        select new
                        {
                            p.sTitle,
                            p.iCTUnit,
                            p.xURL,
                            u.sBaseDSDName,
                            w.CtUnitName
                        }).FirstOrDefault();

            var cnode = (from n in _mGIPcoanew_repository.List<CatTreeNode>()
                         join r in _mGIPcoanew_repository.List<CatTreeRoot>() on n.CtRootID equals r.CtRootID into nr
                         from r in nr.DefaultIfEmpty()
                         where n.CtUnitID == cudt.iCTUnit
                         select new
                         {
                             r.CtRootName
                         }).FirstOrDefault();

            var sheader = _mGIPcoanew_repository.Get<InvalidURLHeader>(p => p.ID == sid);

            Label lsource = (Label)e.Row.FindControl("source");
            lsource.Text = cnode.CtRootName;
            Label ltitle = (Label)e.Row.FindControl("title");
            ltitle.Text = cudt.sTitle;
            Label removeDate = (Label)e.Row.FindControl("removeDate");
            if (!string.IsNullOrEmpty(sheader.removeDate.ToString()))
            {
                removeDate.Text = DateTime.Parse(sheader.removeDate.ToString()).ToShortDateString();
            }
            else
            {
                removeDate.Text = "";
            }
            Label url = (Label)e.Row.FindControl("url");
            url.Text = cudt.xURL;
        }
    }


}
