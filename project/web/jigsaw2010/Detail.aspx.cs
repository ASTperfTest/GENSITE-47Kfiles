using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using GSS.Vitals.COA.Data;
using System.Web.UI.WebControls;

public partial class Detail : System.Web.UI.Page
{
    protected string county;
    protected int? item;
    protected reOrder[] reOrders;
    protected string guid;
    protected struct reOrder
    {
        public string css;
        public string title;
    }
	protected string memberId = "";

    protected int attachmentCount = 0;

    protected void Page_Load(object sender, EventArgs e)
    {
		if (Session["memID"] != null)
		{
			memberId = Session["memID"].ToString() ;
			if(Request.QueryString["kpi"] ==null || Request.QueryString["kpi"] == "")
			{
				KPIBrowse kpiBrowse = new KPIBrowse(memberId, "browseInterCP", "6632");
				kpiBrowse.HandleBrowse();
				string relink = "";
				relink = "/jigsaw2010/Detail.aspx?item=" + Request.QueryString["item"].ToString() + "&kpi=0";
				Response.Redirect(relink);
				Response.End();
			}
		}
		
        if (IsPostBack)
        {
            string oldGuid = "";
            bool checkPic = true;
            if (Request.Form["MemberGuid"].ToString() != "")
            {
                oldGuid = Request.Form["MemberGuid"].ToString();
            }
            if (Session[oldGuid + "CaptchaImageText"] == null)
            {
                checkPic = false;
                ClientScript.RegisterClientScriptBlock(this.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('驗證碼失效!!');window.location.href='" + Request.RawUrl + "';</script>");
            }
            else if (this.MemberCaptChaTBox.Text != Session[oldGuid + "CaptchaImageText"].ToString())
            {
                checkPic = false;
                ClientScript.RegisterClientScriptBlock(this.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('驗證碼錯誤!!');history.back();</script>");
            }
            if (checkPic)
            {
                Session.Remove(oldGuid + "CaptchaImageText");
                Response.Redirect(string.Format("/addDiscussion.asp?xItem={0}&CheckCode={1}&txtDisCussion={2}", Request.QueryString["item"], "true", Server.HtmlEncode(Request.Form["txtDisCussion"])));
            }
        }
        else
        {
            guid = System.Guid.NewGuid().ToString();
            guid = guid.Replace("-", "");
            MemberCaptChaImage.ImageUrl = "~/CaptchaImage/JpegImage.aspx?guid=" + guid;
        }
        try
        {
            // LINQ to SQL
            var db = new mGIPcoanewDataContext();

            // QueryString
            county = (Request.QueryString["county"] ?? "");
            item = int.Parse(Request.QueryString["item"] ?? "0");



            // 取得附件的圖片  Start
            string sqlString = @"SELECT count(*)
                                 FROM CuDTGeneric AS A INNER JOIN CuDTAttach AS B ON A.iCuItem = B.xiCuItem
                                 WHERE A.icuitem = @item AND b.bList = 'y'";

            this.attachmentCount = Convert.ToInt32(SqlHelper.ReturnScalar("ODBCDSN", sqlString,
                DbProviderFactories.CreateParameter("ODBCDSN", "@item", "@item", item)));
            //----------------------------------------------------------------

            // appSettings (Web.config)
            int rndRows = int.Parse(myFunc.GetWebSetting("jigsaw2010", "appSettings") ?? "5");
            Random random = new Random();

            // CuDTGeneric
            var rsCuDTGeneric = (from p in db.CuDTGeneric
                                 where (p.iCTUnit == 2199) && (p.iBaseDSD == 7) && p.fCTUPublic == 'Y' && p.iCUItem == item
                                 select new
                                 {
                                     p.iCUItem,
                                     p.RSS,
                                     p.sTitle,
                                     p.xImgFile,
                                 }).First();
            // Crop | Fish
            if (rsCuDTGeneric.RSS == '0')
            {
                var result10 = (from p in db.Crop
                                where p.iCUItem == rsCuDTGeneric.iCUItem
                                select new
                                {
                                    rsCuDTGeneric.iCUItem,
                                    rsCuDTGeneric.sTitle,
                                    xImgFile = myFunc.GetDefaultImg(rsCuDTGeneric.xImgFile, rsCuDTGeneric.RSS.ToString(), p.type.ToString()),
                                    p.season,
                                    p.origin,
                                    p.feature,
                                    p.cuisine,
                                    p.nutrient,
                                    p.nutritionValue,
                                    p.selectionMethod,
                                    p.alias,
                                    p.description,
                                    p.variety,
                                    p.note,
                                    p.tips,
                                }).Take(1);

                this.rptList10.DataSource = result10;
                this.rptList10.DataBind();
            }
            else
            {
                var result11 = (from p in db.Fish
                                where p.iCUItem == rsCuDTGeneric.iCUItem
                                select new
                                {
                                    rsCuDTGeneric.iCUItem,
                                    rsCuDTGeneric.sTitle,
                                    xImgFile = myFunc.GetDefaultImg(rsCuDTGeneric.xImgFile, rsCuDTGeneric.RSS.ToString(), null),
                                    p.season,
                                    p.origin,
                                    p.distributionInTaiwan,
                                    p.commonName,
                                    p.family,
                                    p.habitatsType,
                                    p.characteristic,
                                    p.distribution,
                                    p.habitats,
                                    p.reference,
                                    p.scientificName,
                                    p.utility,
                                }).Take(1);

                this.rptList11.DataSource = result11;
                this.rptList11.DataBind();
            }

            //最新議題
            var result2 = (from p in db.KnowledgeJigsaw.Where(p => p.parentIcuitem == item)
                           join s in db.CuDTGeneric on p.gicuitem equals s.iCUItem
                           where s.topCat == "A" && s.fCTUPublic == 'Y'
                           select p.gicuitem).FirstOrDefault();

            var result20 = (result2 == null) ? null :
                            (from p in db.KnowledgeJigsaw.Where(p => p.parentIcuitem == result2 && p.Status == 'Y')
                             join s in db.CuDTGeneric on p.gicuitem equals s.iCUItem
                             orderby p.orderArticle descending
                             select new
                             {
                                 p.ArticleId,
                                 s.sTitle,
                                 p.path,
                             }).AsEnumerable().OrderBy(c => random.Next()).Take(rndRows);

            //最新議題-首筆
            var rs201 = result20.FirstOrDefault();
            var rs202 = (rs201 == null) ? null : db.CuDTGeneric.Where(p => p.iCUItem == rs201.ArticleId).FirstOrDefault();
            var result201 = new
                           {
                               xImgFile = (rs202 != null && !string.IsNullOrEmpty(rs202.xImgFile)) ? "<img src=\"" + rs202.xImgFile + "\" alt=\"" + rs202.sTitle + "\">" : "",
                               xBody = (rs202 == null) ? "" : (rs202.xBody != null && rs202.xBody.Length > 170) ? rs202.xBody.Substring(0, 170) : rs202.xBody,
                               sTitle = (rs201 == null) ? "" : rs201.sTitle,
                               path = (rs201 == null) ? "" : rs201.path,
                           };
            var result21 = new object[] { result201 };

            //最新議題-餘筆
            var result22 = from p in result20.Skip(1)
                           select new
                           {
                               p.sTitle,
                               p.path,
                           };

            //關聯文章
            var result3 = (from p in db.KnowledgeJigsaw.Where(p => p.parentIcuitem == item)
                           join s in db.CuDTGeneric on p.gicuitem equals s.iCUItem
                           where s.topCat == "B" && s.fCTUPublic == 'Y'
                           select p.gicuitem).FirstOrDefault();

            var result30 = (result3 == null) ? null :
                            from p in db.KnowledgeJigsaw.Where(p => p.parentIcuitem == result3 && p.Status == 'Y')
                            join s in db.CuDTGeneric on p.gicuitem equals s.iCUItem
                            select new
                            {
                                s.sTitle,
                                p.CtRootId,
                                p.path,

                            };

            //關聯文章-排序
            KnowledgeJigsaw rsOrder = (from p in db.KnowledgeJigsaw.Where(p => p.parentIcuitem == item)
                                       join s in db.CuDTGeneric on p.gicuitem equals s.iCUItem
                                       where s.topCat == "C"
                                       select p).First();
            //關聯文章-排一
            var result31 = result30.Where(p => p.CtRootId == getOrder(rsOrder, 1)).AsEnumerable().OrderBy(c => random.Next()).Take(rndRows).Select((a, i) => new { a.sTitle, a.path, index = i + 1 });
            //關聯文章-排二
            var result32 = result30.Where(p => p.CtRootId == getOrder(rsOrder, 2)).AsEnumerable().OrderBy(c => random.Next()).Take(rndRows).Select((a, i) => new { a.sTitle, a.path, index = i + 1 });
            //關聯文章-排三
            var result33 = result30.Where(p => p.CtRootId == getOrder(rsOrder, 3)).AsEnumerable().OrderBy(c => random.Next()).Take(rndRows).Select((a, i) => new { a.sTitle, a.path, index = i + 1 });
            //關聯文章-排四
            var result34 = result30.Where(p => p.CtRootId == getOrder(rsOrder, 4)).AsEnumerable().OrderBy(c => random.Next()).Take(rndRows).Select((a, i) => new { a.sTitle, a.path, index = i + 1 });

            //關聯影音
            var result4 = (from p in db.KnowledgeJigsaw.Where(p => p.parentIcuitem == item)
                           join s in db.CuDTGeneric on p.gicuitem equals s.iCUItem
                           where s.topCat == "D" && s.fCTUPublic == 'Y'
                           select p.gicuitem).FirstOrDefault();

            var result40 = (result4 == null) ? null :
                           from p in db.KnowledgeJigsaw.Where(p => p.parentIcuitem == result4 && p.Status == 'Y')
                           join s in db.CuDTGeneric on p.gicuitem equals s.iCUItem
                           select new
                           {
                               s.sTitle,
                               p.CtRootId,
                               p.path,
                           };
            var result41 = result40.AsEnumerable().OrderBy(c => random.Next()).Take(rndRows).Select((a, i) => new { a.sTitle, a.path, index = i + 1 });

            //最佳資源推薦
            var result5 = (from p in db.KnowledgeJigsaw.Where(p => p.parentIcuitem == item)
                           join s in db.CuDTGeneric on p.gicuitem equals s.iCUItem
                           where s.topCat == "E" && s.fCTUPublic == 'Y'
                           select p.gicuitem).FirstOrDefault();

            var result50 = (result5 == null) ? null :
                           from p in db.KnowledgeJigsaw.Where(p => p.parentIcuitem == result5 && p.Status == 'Y')
                           join s in db.CuDTGeneric on p.gicuitem equals s.iCUItem
                           select new
                           {
                               s.sTitle,
                               p.path,
                           };
            var result51 = result50.AsEnumerable().OrderBy(c => random.Next()).Take(rndRows).Select((a, i) => new { a.sTitle, a.path, index = i + 1 });

            //議題分享
            var result6 = (from p in db.KnowledgeJigsaw.Where(p => p.parentIcuitem == item)
                           join s in db.CuDTGeneric on p.gicuitem equals s.iCUItem
                           where s.topCat == "F" && s.fCTUPublic == 'Y'
                           select p.gicuitem).FirstOrDefault();

            var result60 = (result6 == null) ? null :
                           from p in db.KnowledgeJigsaw.Where(p => p.parentIcuitem == result6 && p.Status == 'Y')
                           join s in db.CuDTGeneric on p.gicuitem equals s.iCUItem
                           where s.xBody != null && s.fCTUPublic == 'Y'
                           orderby s.xPostDate descending
                           select new
                           {
                               s.xBody,
                               s.iEditor,
                               xpostdate = String.Format("{0:yyyy/MM/dd}", s.xPostDate),
                           };
            var result61 = result60.AsEnumerable().Select((a, i) => new { a.xBody, a.iEditor, a.xpostdate, index = i + 1 });

            if (result20.Count() > 0)
            {
                this.rptList20.DataSource = result20;//最新議題-頭
                this.rptList20.DataBind();
                this.rptList21.DataSource = result21;//最新議題-首筆
                this.rptList21.DataBind();
                this.rptList22.DataSource = result22;//最新議題-餘筆
                this.rptList22.DataBind();
                this.rptList29.DataSource = result20;//最新議題-尾
                this.rptList29.DataBind();
            }
            if (result30.Count() + result40.Count() > 0)
            {
                this.rptList30.DataSource = result30;//關聯文章-頭
                this.rptList30.DataBind();
            }
            if (result31.Count() > 0)
            {
                this.rptList31.DataSource = result31;//關聯文章-排一
                this.rptList31.DataBind();
            }
            if (result32.Count() > 0)
            {
                this.rptList32.DataSource = result32;//關聯文章-排二
                this.rptList32.DataBind();
            }
            if (result33.Count() > 0)
            {
                this.rptList33.DataSource = result33;//關聯文章-排三
                this.rptList33.DataBind();
            }
            if (result34.Count() > 0)
            {
                this.rptList34.DataSource = result34;//關聯文章-排四
                this.rptList34.DataBind();
            }
            if (result41.Count() > 0)
            {
                this.rptList41.DataSource = result41;//關聯影音
                this.rptList41.DataBind();
            }
            if (result51.Count() > 0)
            {
                this.rptList51.DataSource = result51;//最佳資源推薦
                this.rptList51.DataBind();
            }
            this.rptList61.DataSource = result61;//議題分享
            this.rptList61.DataBind();

            // UI Page Contents
            {
                nav.Text = string.Format("<a href=\"{0}\" title=\"{1}\">{1}</a> &gt; <a href=\"{2}\" title=\"{3}\">{3}</a> &gt; ", "/", "首頁", "Index.aspx", "農漁生產地圖");
                if (county != "")
                    nav.Text += string.Format("<a href=\"{0}\" title=\"{1}\">{1}</a> &gt; ", "Sub_List.aspx?county=" + Server.UrlEncode(county), county);

                nav.Text += string.Format("<span>{0}</span>", rsCuDTGeneric.sTitle);

                if (county != "")
                {
                    nav.Text = string.Format("<ul id='path_menu'><li><a href=\"{0}\" title=\"{1}\">{1}</a></li><li style='top:10px;'>></li><li><a href=\"{2}\" title=\"{3}\">{3}</a></li><li style='top:10px;'>></li><li><a href=\"{4}\" title=\"{5}\">{5}</a></li><li style='top:10px;'>></li><li><a href=\"{6}\" title=\"{7}\">{7}</a></li></ul>"
                        , "/", "首頁"
                        , "Index.aspx", "農漁生產地圖"
                        , "Sub_List.aspx?county=" + Server.UrlEncode(county), county
                        , "#", rsCuDTGeneric.sTitle
                        );
                }
                else
                {
                    nav.Text = string.Format("<ul id='path_menu'><li><a href=\"{0}\" title=\"{1}\">{1}</a></li><li style='top:10px;'>></li><li><a href=\"{2}\" title=\"{3}\">{3}</a></li><li style='top:10px;'>></li><li><a href=\"{4}\" title=\"{5}\">{5}</a></li></ul>"
                        , "/", "首頁"
                        , "Index.aspx", "農漁生產地圖"
                        , "#", rsCuDTGeneric.sTitle
                        );
                }


                setupOrders(rsOrder);
            }
        }
        catch (Exception ex)
        {
            //Response.Write(ex);
            Redirect(string.Format("發生錯誤：{0}", ex.Message));
        }
    }
    private void Redirect(string msg)
    {
        if (county != "")
            myFunc.GenErrMsg(msg, string.Format("window.location.href=\"Sub_List.aspx?county={0}\";", Server.UrlEncode(county)));
        else
            myFunc.GenErrMsg(msg, string.Format("window.location.href=\"Index.aspx\";"));
    }
    private int getOrder(KnowledgeJigsaw rs, int order)
    {
        if (rs.orderSiteUnit == order)
            return 1;
        else if (rs.orderSubject == order)
            return 2;
        else if (rs.orderKnowledgeHome == order)
            return 3;
        else if (rs.orderKnowledgeTank == order)
            return 4;
        else
            return 0;
    }
    private void setupOrders(KnowledgeJigsaw rs)
    {
        reOrders = new reOrder[4];
        for (int i = 0; i < reOrders.Length; i++)
        {
            switch (getOrder(rs, i + 1))
            {
                case 1:
                    reOrders[i] = new reOrder { css = "jigsawtype02", title = "入口網關聯文章" };
                    break;
                case 2:
                    reOrders[i] = new reOrder { css = "jigsawtype03", title = "主題館關聯文章" };
                    break;
                case 3:
                    reOrders[i] = new reOrder { css = "jigsawtype05", title = "知識家關聯文章" };
                    break;
                case 4:
                    reOrders[i] = new reOrder { css = "jigsawtype04", title = "知識庫關聯文章" };
                    break;
                default:
                    reOrders[i] = new reOrder { css = "", title = "" };
                    break;
            }
        }
    }


    protected string GetImgString(object sTitle, object xImgFile)
    {
        if (attachmentCount == 0)
        {
            return "<img alt='" + sTitle.ToString() + "' src='" + xImgFile.ToString() + "'  class='photo' />";
        }
        else
        {
            return "<iframe width='382' height='280px' style='border:0px;' src='FileShowPage.aspx?item=" + Request.QueryString["item"] + "'></iframe>";
        }
    }
}
