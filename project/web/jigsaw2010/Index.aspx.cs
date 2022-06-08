using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.Linq.SqlClient;

public partial class Index : System.Web.UI.Page
{
    protected bool springblock = true;
    protected bool summerblock = true;
    protected bool autumnblock = true;
    protected bool winterblock = true;
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            nav.Text = string.Format("<ul id='path_menu'><li><a href=\"{0}\" title=\"{1}\">{1}</a></li><li style='top:10px;'>></li><li><a href=\"{2}\" title=\"{3}\">{3}</a></li></ul>", "/", "首頁", "Index.aspx", "農漁生產地圖");

            //增加more顯示所有作物
            string season = WebUtility.GetStringParameter("season", "");
            bool moreVisable = true;
            if (!string.IsNullOrEmpty(season))
            {
                switch (season)
                {
                    case "spring":
                        summerblock = false;
                        autumnblock = false;
                        winterblock = false;
                        moreVisable = false;
                        break;
                    case "summer":
                        springblock = false;
                        autumnblock = false;
                        winterblock = false;
                        moreVisable = false;
                        break;
                    case "autumn":
                        springblock = false;
                        summerblock = false;
                        winterblock = false;
                        moreVisable = false;
                        break;
                    case "winter":
                        springblock = false;
                        summerblock = false;
                        autumnblock = false;
                        moreVisable = false;
                        break;
                    default:
                        break;
                }
            }
            season2010spring.Visible = springblock;
            morespring.Visible = moreVisable;
            season2010summer.Visible = summerblock;
            moresummer.Visible = moreVisable;
            season2010autumn.Visible = autumnblock;
            moreautumn.Visible = moreVisable;
            season2010winter.Visible = winterblock;
            moreWinter.Visible = moreVisable;

            // LINQ to SQL
            var db = new mGIPcoanewDataContext();
            int ismore = 0;
            if (!string.IsNullOrEmpty(season))
                ismore = -1;
            if (springblock)
            {
                //農作物(春)
                var result10 = (from p in db.CuDTGeneric
                                join s in db.Crop on p.iCUItem equals s.iCUItem
                                where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.RSS == '0' && (s.sortOrderSpring > ismore) && SqlMethods.Like(s.season, "%[3-5]月%")
                                orderby s.sortOrderSpring
                                select new
                                {
                                    p.iCUItem,
                                    p.sTitle,
                                });
                //魚種(春)
                var result11 = (from p in db.CuDTGeneric
                                join s in db.Fish on p.iCUItem equals s.iCUItem
                                where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.RSS == '1' && (s.sortOrderSpring > ismore) && SqlMethods.Like(s.season, "%[3-5]月%")
                                orderby s.sortOrderSpring
                                select new
                                {
                                    p.iCUItem,
                                    p.sTitle,
                                });
                this.rptList10.DataSource = result10;
                this.rptList10.DataBind();
                this.rptList11.DataSource = result11;
                this.rptList11.DataBind();
            }
            if (summerblock)
            {
                //農作物(夏)
                var result20 = (from p in db.CuDTGeneric
                                join s in db.Crop on p.iCUItem equals s.iCUItem
                                where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.RSS == '0' && (s.sortOrderSummer > ismore) && SqlMethods.Like(s.season, "%[6-8]月%")
                                orderby s.sortOrderSummer
                                select new
                                {
                                    p.iCUItem,
                                    p.sTitle,
                                });
                //魚種(夏)
                var result21 = (from p in db.CuDTGeneric
                                join s in db.Fish on p.iCUItem equals s.iCUItem
                                where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.RSS == '1' && (s.sortOrderSummer > ismore) && SqlMethods.Like(s.season, "%[6-8]月%")
                                orderby s.sortOrderSummer
                                select new
                                {
                                    p.iCUItem,
                                    p.sTitle,
                                });
                this.rptList20.DataSource = result20;
                this.rptList20.DataBind();
                this.rptList21.DataSource = result21;
                this.rptList21.DataBind();
            }
            if (autumnblock)
            {
                //農作物(秋)
                var result30 = (from p in db.CuDTGeneric
                                join s in db.Crop on p.iCUItem equals s.iCUItem
                                where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.RSS == '0' && (s.sortOrderAutumn > ismore) && (s.season.Contains("9月") || SqlMethods.Like(s.season, "%1[0-1]月%"))
                                orderby s.sortOrderAutumn
                                select new
                                {
                                    p.iCUItem,
                                    p.sTitle,
                                });
                //魚種(秋)
                var result31 = (from p in db.CuDTGeneric
                                join s in db.Fish on p.iCUItem equals s.iCUItem
                                where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.RSS == '1' && (s.sortOrderAutumn > ismore) && (s.season.Contains("9月") || SqlMethods.Like(s.season, "%1[0-1]月%"))
                                orderby s.sortOrderAutumn
                                select new
                                {
                                    p.iCUItem,
                                    p.sTitle,
                                });
                this.rptList30.DataSource = result30;
                this.rptList30.DataBind();
                this.rptList31.DataSource = result31;
                this.rptList31.DataBind();
            }
            if (winterblock)
            {
                //農作物(冬)
                var result40 = (from p in db.CuDTGeneric
                                join s in db.Crop on p.iCUItem equals s.iCUItem
                                where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.RSS == '0' && (s.sortOrderWinter > ismore) && (s.season.StartsWith("1月") || s.season.Contains("2月") || SqlMethods.Like(s.season, "%[^1]1月%"))
                                orderby s.sortOrderWinter
                                select new
                                {
                                    p.iCUItem,
                                    p.sTitle,
                                });
                //魚種(冬)
                var result41 = (from p in db.CuDTGeneric
                                join s in db.Fish on p.iCUItem equals s.iCUItem
                                where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.RSS == '1' && (s.sortOrderWinter > ismore) && (s.season.StartsWith("1月") || s.season.Contains("2月") || SqlMethods.Like(s.season, "%[^1]1月%"))
                                orderby s.sortOrderWinter
                                select new
                                {
                                    p.iCUItem,
                                    p.sTitle,
                                });
                this.rptList40.DataSource = result40;
                this.rptList40.DataBind();
                this.rptList41.DataSource = result41;
                this.rptList41.DataBind();
            }
            //Data Repeater DataBinding
            //this.rptList10.DataSource = result10;
            //this.rptList10.DataBind();
            //this.rptList11.DataSource = result11;
            //this.rptList11.DataBind();
            //this.rptList20.DataSource = result20;
            //this.rptList20.DataBind();
            //this.rptList21.DataSource = result21;
            //this.rptList21.DataBind();
            //this.rptList30.DataSource = result30;
            //this.rptList30.DataBind();
            //this.rptList31.DataSource = result31;
            //this.rptList31.DataBind();
            //this.rptList40.DataSource = result40;
            //this.rptList40.DataBind();
            //this.rptList41.DataSource = result41;
            //this.rptList41.DataBind();
        }
        catch
        {
            Response.Redirect(@"/");
        }
    }
}
