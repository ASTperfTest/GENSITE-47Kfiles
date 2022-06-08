using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Data;
using System.Collections;
using System.Data.SqlClient;


/// <summary>
/// Created By Leo  --  2011-07-06
/// </summary>
public partial class Recommand_List : System.Web.UI.Page
{

    private PagedDataSource Pager;
    private DataTable dt;
    protected string memId = "";

    // 預設
    protected void Page_Init(object sender, EventArgs e)
    {
        
        if (Session["memID"] == null) { memId = "none"; }
        else if (string.IsNullOrEmpty(Session["memID"].ToString())) { memId = "none"; }
        else { memId = ""; }

        // 標籤
        string sqlQueryScript = @"SELECT * FROM TAGs ORDER BY createdDate";
        using (var reader = SqlHelper.ReturnReader("ODBCDSN", sqlQueryScript))
        {
            if (reader.HasRows)
            {
                int i = 0;
                LinkButton item = new LinkButton();
                Label labulStart = new Label();
                Label labulEnd = new Label();
                Label labliStart = new Label();
                Label labliEnd = new Label();
                labulStart.Text = "<ul>";
                labulEnd.Text = "</ul>";
                labliStart.Text = "<li>";
                labliEnd.Text = "</li>";
                item.ID = "lbtnAll";
                item.Text = "全部";
                item.CssClass = "LinkButton";
                item.Click += new EventHandler(TAGs_Click);

                panelTAGs.Controls.Add(labulStart);
                panelTAGs.Controls.Add(labliStart);
                panelTAGs.Controls.Add(item);
                panelTAGs.Controls.Add(labliEnd);

                while (reader.Read())
                {
                    i++;
                    Label labTAGsliStart = new Label();
                    Label labTAGsliEnd = new Label();
                    labTAGsliStart.Text = "<li>";
                    labTAGsliStart.ID = "labTAGsliStart_" + i.ToString();   // 若沒設定ID，系統會判定ID重覆而不產生(不是再說廢話嗎 XD) By Leo
                    labTAGsliEnd.Text = "</li>";
                    labTAGsliEnd.ID = "labTAGsliEnd_" + i.ToString();
                    LinkButton lbtnTAGs = new LinkButton();
                    lbtnTAGs.ID = "lbtnTAGs_" + i.ToString();
                    lbtnTAGs.Text = reader["tagName"].ToString();
                    lbtnTAGs.CssClass = "LinkButton";
                    lbtnTAGs.Click += new EventHandler(TAGs_Click);
                    panelTAGs.Controls.Add(labTAGsliStart);
                    panelTAGs.Controls.Add(lbtnTAGs);
                    panelTAGs.Controls.Add(labTAGsliEnd);
                }
                panelTAGs.Controls.Add(labulEnd);
            }
        }
    }

    protected void TAGs_Click(object sender, EventArgs e)
    {
        
        if (((LinkButton)sender).ID != "lbtnAll")
        {
            myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue),((LinkButton)sender));
        }
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            int PageNumber = 0;
            int PageSize = 10;
            Getrule();
            myDBinit(PageNumber, PageSize);
        }
        else
        {
            /* 若做資料排列，則會出現 無效的回傳或回呼引數。... 這錯誤；
             * 最主要是不要在 PostBack 裡再做一次 DataBind();
             * 所以屆時真的出錯，只好將Databind改到各個Control Events去寫。*/
            myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
        }
    }

    protected void Getrule()
    {
        string sql;
        int rkey = int.Parse(System.Web.Configuration.WebConfigurationManager.AppSettings["Recommandkey"].ToString());
        sql = @"SELECT *   FROM [CuDTGeneric] 
                where icuitem=" + rkey;
         var dr = SqlHelper.ReturnReader("ConnString", sql);
         if (dr.Read())
         {
             lbtitle.Text = dr["sTitle"].ToString();
             string ss = dr["xBody"].ToString();
             lbabstract.Text = ss.Substring(0, ss.IndexOf("。") + 1);
             Image1.ImageUrl = System.Web.Configuration.WebConfigurationManager.AppSettings["WWWUrl"] + "/public/data/" + dr["xImgFile"].ToString();
         }        

    }
    protected void myDBinit(int intPageNumber, int intPageSize)
    {
        string sqlQueryScript;
            sqlQueryScript = @"SELECT 
                      RecommandContent.* 
                    , isnull(rtrim(member.nickname), '') as nickname
                    , isnull(rtrim(member.realname), '') as realname
                FROM RecommandContent 
                left join member on RecommandContent.iEditor = member.account
                WHERE RecommandContent.Status = 'y' 
                    AND (RecommandContent.aContent LIKE '%'+ @kw +'%' 
                        or RecommandContent.title LIKE '%'+ @kw +'%'
                        or RecommandContent.URL LIKE '%'+ @kw +'%'
                        or RecommandContent.Source LIKE '%'+ @kw +'%'
                        )
                ORDER BY RecommandContent.aEditDate desc";
            dt = SqlHelper.GetDataTable("ConnString", sqlQueryScript,
               DbProviderFactories.CreateParameter("ConnString", "@kw", "@kw", txtSearch.Text));

        Pager = dt.Paging(intPageNumber, intPageSize);
        rptList.DataSource = Pager;
        rptList.DataBind();

        SetControl();
    }

    protected void myDBinit(int intPageNumber, int intPageSize, LinkButton lbtnTAG)
    {
        string sqlQueryScript;
        sqlQueryScript = @"SELECT DISTINCT a.*
                        , isnull(rtrim(member.nickname), '') as nickname
                        , isnull(rtrim(member.realname), '') as realname
                           FROM RecommandContent AS a 
                                INNER JOIN RecommandContent2TAGs AS b ON a.cid= b.cid	
                                left join member on a.iEditor = member.account
                                INNER JOIN TAGs AS c ON b.tagid = c.tagid 
                           WHERE a.Status = 'y' 
                                AND c.tagName = @TAGs 
                                AND (a.aContent LIKE '%'+ @kw +'%' 
                                        or a.title LIKE '%'+ @kw +'%'
                                        or a.URL LIKE '%'+ @kw +'%'
                                        or a.Source LIKE '%'+ @kw +'%'
                                    )
                           ORDER BY a.aEditDate desc";
            dt = SqlHelper.GetDataTable("ConnString", sqlQueryScript,
                DbProviderFactories.CreateParameter("ConnString", "@TAGs", "@TAGs", lbtnTAG.Text),
                DbProviderFactories.CreateParameter("ConnString", "@kw", "@kw", txtSearch.Text));

        Pager = dt.Paging(intPageNumber, intPageSize);
        rptList.DataSource = Pager;
        rptList.DataBind();

        SetControl();
    }

    // 設置控制項
    protected void SetControl()
    {
        // 設定PageNumber_DropDownList
        PageNumberDDL.Items.Clear();
        for (int i = 0; i < Pager.PageCount; i++)
        {
            PageNumberDDL.Items.Add(new ListItem((i + 1).ToString(), i.ToString()));
        }
        //Response.Write(Pager.CurrentPageIndex);
        PageNumberDDL.SelectedIndex = Pager.CurrentPageIndex;

        if (PageNumberDDL.Items.Count > 1)
        {
            if (Pager.IsFirstPage)
            {
                PreviousLink.Visible = false;
                NextLink.Visible = true;
            }
            else if (Pager.IsLastPage)
            {
                PreviousLink.Visible = true;
                NextLink.Visible = false;
            }
            else
            {
                PreviousLink.Visible = true;
                NextLink.Visible = true;
            }
        }
        else
        {
            PreviousLink.Visible = false;
            NextLink.Visible = false;
        }

        PageNumberText.Text = (Pager.CurrentPageIndex + 1).ToString();
        TotalPageText.Text = Pager.PageCount.ToString();
        TotalRecordText.Text = dt.Rows.Count.ToString();
    }

    protected void PreviousLink_Click(object sender, EventArgs e)
    {
        PageNumberDDL.SelectedIndex--;
        myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }

    protected void NextLink_Click(object sender, EventArgs e)
    {
        PageNumberDDL.SelectedIndex++;
        myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }

    protected string GetShowName(object nname, object rname)
    {
        string nickname = nname == null ? string.Empty : nname.ToString();
        string realname = rname == null ? string.Empty : rname.ToString();
        
        try
        {
            if (!string.IsNullOrEmpty(nickname))
            {
                return nickname;
            }
            else
            {
                if (string.IsNullOrEmpty(realname))
                {
                    return string.Empty;
                }
                else
                {
                    if (realname.Length <= 2)
                    {
                        return realname[0] + "*";
                    }
                    else
                    {
                        return realname[0] + "*" + realname[realname.Length - 1];
                    }
                }
            }
        }
        catch { return string.Empty; }
    }
}