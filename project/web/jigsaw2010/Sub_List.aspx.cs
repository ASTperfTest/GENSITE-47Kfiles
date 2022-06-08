using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Sub_List : System.Web.UI.Page
{
    protected string county;
    protected string types;
    protected Pager pager;
	protected bool searchAllCounty = false;
	protected string searchKey = "";
	class ResultViewData
    {
        public int iCUItem { get; set; }
        public string sTitle { get; set; }
        public string xImgFile { get; set; }
        public string feature { get; set; }
        public string type { get; set; }
    }
    struct fuzzyCounty
    {
        public string county;
        public string fuzzy;
        public string excluded;
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            // LINQ to SQL
            var db = new mGIPcoanewDataContext();

            // QueryString
            county = (Request.QueryString["county"] ?? "").Trim();
            types = (Request.QueryString["types"] ?? "0");
            string categorytype = (Request.QueryString["types"] ?? "-1");
            searchKey = (Request.QueryString["searchkey"] ?? "").Trim();
            if (county == "all")
            {
                searchAllCounty = true;
            }

            nav.Text = string.Format("<a href=\"{0}\" title=\"{1}\">{1}</a> &gt; <a href=\"{2}\" title=\"{3}\">{3}</a> &gt; <a href=\"{4}\" title=\"{5}\">{5}</a>", "/", "首頁", "Index.aspx"
                , "農漁生產地圖"
                , "Sub_List.aspx?county=" + Server.UrlEncode(county) + "&searchkey=" + Server.UrlEncode(searchKey), (searchKey != "") 
                ? "搜尋結果" 
                : county);

            nav.Text = string.Format("<ul id='path_menu'><li><a href=\"{0}\" title=\"{1}\">{1}</a></li><li style='top:10px;'>></li><li><a href=\"{2}\" title=\"{3}\">{3}</a></li><li style='top:10px;'>></li><li><a href=\"{4}\" title=\"{5}\">{5}</a></li></ul>", "/"
                , "首頁", "Index.aspx"
                , "農漁生產地圖"
                , "Sub_List.aspx?county=" + Server.UrlEncode(county) + "&searchkey=" + Server.UrlEncode(searchKey), (searchKey != "")
                ? "搜尋結果"
                : county                
                );

            sub_title.Text = string.Format("{0}農漁生產地圖", county);

            fuzzyCounty fuzzy_county = GetFuzzyCounty(county);
            var result = new List<ResultViewData>();
            result = res(types, db, fuzzy_county);
            if (categorytype == "-1" && result.Count == 0)
            {
                result = res("1", db, fuzzy_county);
                if (categorytype == "-1" && result.Count >= 1)
                {
                    types = "1";
                }
            }
            //if (types != "0")
            //{
            //    result = res(categorytype, db, fuzzy_county);
            //    //魚種
            //    result = (from p in db.CuDTGeneric
            //              join f in db.Fish on p.iCUItem equals f.iCUItem
            //              where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.RSS == '1' && (f.origin.Contains(fuzzy_county.fuzzy) || searchAllCounty) && !(f.origin.Contains(fuzzy_county.excluded) && !f.origin.Contains(fuzzy_county.county) && !searchAllCounty) && (p.sTitle.Contains(searchKey) || f.commonName.Contains(searchKey) || (searchKey == ""))
            //              orderby p.Created_Date descending
            //              select new ResultViewData
            //              {
            //                  iCUItem = p.iCUItem,
            //                  sTitle = p.sTitle,
            //                  xImgFile = p.xImgFile,
            //                  feature = f.characteristic,
            //              }).ToList();
            //}
            //else
            //{
            //    //農作物
            //    result = (from p in db.CuDTGeneric
            //              join c in db.Crop on p.iCUItem equals c.iCUItem
            //              where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.RSS == '0' && (c.origin.Contains(fuzzy_county.fuzzy) || searchAllCounty) && !(c.origin.Contains(fuzzy_county.excluded) && !c.origin.Contains(fuzzy_county.county) && !searchAllCounty) && (p.sTitle.Contains(searchKey) || c.alias.Contains(searchKey) || (searchKey == ""))
            //              orderby p.Created_Date descending
            //              select new ResultViewData
            //              {
            //                  iCUItem = p.iCUItem,
            //                  sTitle = p.sTitle,
            //                  xImgFile = p.xImgFile,
            //                  feature = c.feature,
            //                  type = c.type.ToString()
            //              }).ToList();
            //}
            //分頁
            int page = int.Parse(Request.QueryString["page"] ?? "0");
            int pageSize = int.Parse(Request.QueryString["pagesize"] ?? "10");
            var PaginatedList = result.Skip(page * pageSize).Take(pageSize).ToArray();
            pager = new Pager(page, pageSize, result.Count(), new string[] { "10", "30", "50" });

            foreach (var p in PaginatedList)
            {
                p.xImgFile = myFunc.GetDefaultImg(p.xImgFile, types, p.type);
                p.feature = Server.HtmlEncode((p.feature.Length > 60) ? p.feature.Substring(0, 60) + " ..." : p.feature);
            }

            if (result.Count == 0)
            {
                listdata.Visible = false;
                nonedata.Visible = true;
            }

            //Data Repeater DataBinding
            this.rptList.DataSource = PaginatedList;
            this.rptList.DataBind();
        }
        catch
        {
            Response.Redirect(@"/");
        }
    }
    private fuzzyCounty GetFuzzyCounty(string county)
    {
        if (county.EndsWith("縣"))
            return new fuzzyCounty { county = county, fuzzy = county.Remove(county.Length - 1), excluded = string.Concat(county.Remove(county.Length - 1), "市") };
        else if (county.EndsWith("市"))
            return new fuzzyCounty { county = county, fuzzy = county.Remove(county.Length - 1), excluded = string.Concat(county.Remove(county.Length - 1), "縣") };
        else
            return new fuzzyCounty { county = county, fuzzy = county, excluded = "" };
    }

    private List<ResultViewData> res(string type, mGIPcoanewDataContext db, fuzzyCounty fuzzy_county)
    {
        var r = new List<ResultViewData>();
        if (type != "0")
        {
            //魚種
            r = (from p in db.CuDTGeneric
                      join f in db.Fish on p.iCUItem equals f.iCUItem
                      where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.RSS == '1' && (f.origin.Contains(fuzzy_county.fuzzy) || searchAllCounty) && !(f.origin.Contains(fuzzy_county.excluded) && !f.origin.Contains(fuzzy_county.county) && !searchAllCounty) && (p.sTitle.Contains(searchKey) || f.commonName.Contains(searchKey) || (searchKey == ""))
                      orderby p.Created_Date descending
                      select new ResultViewData
                      {
                          iCUItem = p.iCUItem,
                          sTitle = p.sTitle,
                          xImgFile = p.xImgFile,
                          feature = f.characteristic,
                      }).ToList();
            return r;
        }
        else
        {
            //農作物
            r = (from p in db.CuDTGeneric
                      join c in db.Crop on p.iCUItem equals c.iCUItem
                      where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.RSS == '0' && (c.origin.Contains(fuzzy_county.fuzzy) || searchAllCounty) && !(c.origin.Contains(fuzzy_county.excluded) && !c.origin.Contains(fuzzy_county.county) && !searchAllCounty) && (p.sTitle.Contains(searchKey) || c.alias.Contains(searchKey) || (searchKey == ""))
                      orderby p.Created_Date descending
                      select new ResultViewData
                      {
                          iCUItem = p.iCUItem,
                          sTitle = p.sTitle,
                          xImgFile = p.xImgFile,
                          feature = c.feature,
                          type = c.type.ToString()
                      }).ToList();
            return r;
        }
    }
}
