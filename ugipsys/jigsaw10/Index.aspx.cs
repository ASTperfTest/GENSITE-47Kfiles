using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class jigsaw10_index : System.Web.UI.Page
{
    protected string titles;
    protected string status;
    protected string types;
    protected PaginatedList<CuDTGeneric> pl;

    protected void Page_Load(object sender, EventArgs e)
    {
        // LINQ to SQL (Repository)
        IRepository _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());

        #region //sql
        //SELECT [iCUItem],[fCTUPublic],[sTitle],[xImgFile],[xBody],[xImportant]
        //FROM [mGIPcoanew].[dbo].[CuDTGeneric]
        //where [iCTUnit]=2199 and [iBaseDSD]=7
        //order by Created_Date desc
        #endregion
        var result = from p in _mGIPcoanew_repository.List<CuDTGeneric>()
                     where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && (p.RSS == '0' || p.RSS == '1')
                     select p;

        //Query Filter
        titles = (Request["sTitle"] ?? "").Trim();//標題
        status = (Request["Status"] ?? "");//狀態
        types = (Request["Types"] ?? "");//農作物或魚種

        if (titles != "")
        {
            result = from p in result
                     where p.sTitle.Contains(titles)
                     select p;
        }
        if (status != "")
        {
            result = from p in result
                     where p.fCTUPublic == status[0]
                     select p;
        }
        if (types != "")
        {
            result = from p in result
                     where p.RSS == types[0]
                     select p;
        }

	//排序
        result = from p in result
                 orderby p.iCUItem
                 select p;

        //分頁
        int page = int.Parse(Request.QueryString["page"] ?? "0");
        int pageSize = int.Parse(Request.QueryString["pagesize"] ?? "10");
        pl = new PaginatedList<CuDTGeneric>(result, page, pageSize, new string[] { "0", "10", "15", "30", "50" });

        //加作物類型說明文字
        var remix = from p in pl
                    join s in _mGIPcoanew_repository.List<Crop>() on p.iCUItem equals s.iCUItem into k
                    from s in k.DefaultIfEmpty()
                    select new
                    {
                        p.iCUItem,
                        p.RSS,
                        p.sTitle,
                        p.fCTUPublic,
                        Categories = (s == null) ? "魚種" : (s.type.ToString() == "0") ? "水果" : (s.type.ToString() == "1") ? "蔬菜" : (s.type.ToString() == "2") ? "花卉" : (s.type.ToString() == "3") ? "雜糧特作" : "",
                    };
                        
        //Data Repeater DataBinding
        this.rptList.DataSource = remix;
        this.rptList.DataBind();
    }
}
