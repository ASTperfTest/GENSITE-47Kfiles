using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Transactions;
using System.Data.Linq;

public partial class EditFish : System.Web.UI.Page
{
    public int id;
    public string returnUrl;
    public string xImgURL = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        // LINQ to SQL (Repository)
        IRepository _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());

        // ReturnURL (Save&Return)
        returnUrl = getReturnURL();

        // appSettings (Web.config)
        string ImgSavePath = StrFunc.GetWebSetting("ImgSavePath", "appSettings");

        string strWebUrl = "";
        string strUploadPath = "/public/data/";
        string strFilePath = Server.MapPath(strUploadPath);
        string strFileName = ""; 
        
        id = int.Parse(Request["id"] ?? "0");

        if (IsPostBack)
        {
            string CropOrFish = Request.Form["CropOrFish"];
            string sTitle = Request.Form["sTitle"];
            //string xImgFile = Request.Form["xImgFile"];
            string sortOrderSpring = Request.Form["sortOrderSpring"];
            string sortOrderSummer = Request.Form["sortOrderSummer"];
            string sortOrderAutumn = Request.Form["sortOrderAutumn"];
            string sortOrderWinter = Request.Form["sortOrderWinter"];
            string fCTUPublic = Request.Form["fCTUPublic"];
            string season = Request.Form["season"];
            string origin = Request.Form["origin"];
            string family = Request.Form["family"];
            string scientificName = Request.Form["scientificName"];
            string commonName = Request.Form["commonName"];
            string distributionInTaiwan = Request.Form["distributionInTaiwan"];
            string habitats = Request.Form["habitats"];
            string reference = Request.Form["reference"];
            string characteristic = Request.Form["characteristic"];
            string habitatsType = Request.Form["habitatsType"];
            string distribution = Request.Form["distribution"];
            string utility = Request.Form["utility"];

            if (string.IsNullOrEmpty(sTitle))
            {
                StrFunc.GenErrMsg("請輸入中文名稱！", string.Format("Javascript:window.history.back();"));
            }
            if (string.IsNullOrEmpty(scientificName))
            {
                StrFunc.GenErrMsg("請輸入學名！", string.Format("Javascript:window.history.back();"));
            }
            if (string.IsNullOrEmpty(season))
            {
                StrFunc.GenErrMsg("請輸入產期！", string.Format("Javascript:window.history.back();"));
            }
            if (string.IsNullOrEmpty(origin))
            {
                StrFunc.GenErrMsg("請輸入產地！", string.Format("Javascript:window.history.back();"));
            }

            if (id != 0)
            {
                var originCuDTGeneric = _mGIPcoanew_repository.Get<CuDTGeneric>(p => p.iCUItem == id);
                var originFish = _mGIPcoanew_repository.Get<Fish>(p => p.iCUItem == id);
                if (originCuDTGeneric != null && originFish != null)
                {
                    try
                    {
                        if (Request.Form["action_equals_Del"] == "1")
                        {
                            #region //DELETE
                            string strSql = "delete from CuDTGeneric where iCUItem=" + StrFunc.NZQty(id.ToString()) + ";" +
                                            "delete from Fish where iCUItem=" + StrFunc.NZQty(id.ToString()) + ";";

                            using (DataContext db = new mGIPcoanewDataContext()) { db.ExecuteCommand(strSql); }
                            #endregion
                        }
                        else
                        {
                            #region //UPDATE
                            if (xImgFile.PostedFile != null)
                            {
                                // Get a reference to PostedFile object
                                HttpPostedFile uploadFile = xImgFile.PostedFile;
                                try
                                {
                                    if (uploadFile.ContentLength > 0)
                                    {
                                        strFileName = ImgSavePath + id + uploadFile.FileName.Substring(uploadFile.FileName.LastIndexOf('.'));
                                        uploadFile.SaveAs(strFilePath + "\\" + strFileName);
                                    }
                                }
                                finally { }
                            }
                            using (TransactionScope ts = new TransactionScope())
                            {
                                originCuDTGeneric.sTitle = sTitle;
                                originCuDTGeneric.xImgFile = (strFileName != "" ? strFileName : originCuDTGeneric.xImgFile);
                                originCuDTGeneric.fCTUPublic = fCTUPublic[0];
                                originFish.season = season;
                                originFish.origin = origin;
                                originFish.family = family;
                                originFish.scientificName = scientificName;
                                originFish.commonName = commonName;
                                originFish.distributionInTaiwan = distributionInTaiwan;
                                originFish.habitats = habitats;
                                originFish.reference = reference;
                                originFish.characteristic = characteristic;
                                originFish.habitatsType = habitatsType;
                                originFish.distribution = distribution;
                                originFish.utility = utility;
                                originFish.sortOrderSpring = int.Parse(sortOrderSpring);
                                originFish.sortOrderSummer = int.Parse(sortOrderSummer);
                                originFish.sortOrderAutumn = int.Parse(sortOrderAutumn);
                                originFish.sortOrderWinter = int.Parse(sortOrderWinter);

                                _mGIPcoanew_repository.Save();
                                ts.Complete();
                            }
                            #endregion
                        }
                        Redirect();
                    }
                    catch (Exception ex) { Response.Write(string.Format("<font size='2' color='red'><strong>發生錯誤：{0}</strong></font>", ex)); }
                }
            }
            else
            {
                try
                {
                    #region //INSERT
                    using (TransactionScope ts = new TransactionScope())
                    {
                        #region //文章主體
                        var newCuDTGeneric = new CuDTGeneric
                        {
                            iEditor = "hyweb",
                            iDept = "0",
                            showType = '1',
                            siteId = '1',
                            iBaseDSD = 7,
                            iCTUnit = 2199,
                            sTitle = sTitle,
                            xImgFile = strFileName,
                            fCTUPublic = fCTUPublic[0],
                            RSS = CropOrFish[0],
                            Created_Date = DateTime.Now,
                            dEditDate = DateTime.Now,
                            xPostDate = DateTime.Now,
                            ClickCount = 0,
                            HaveBuildIndex = 0,
                        };
                        _mGIPcoanew_repository.Create<CuDTGeneric>(newCuDTGeneric);
                        
                        if (xImgFile.PostedFile != null)
                        {
                            // Get a reference to PostedFile object
                            HttpPostedFile uploadFile = xImgFile.PostedFile;
                            try
                            {
                                if (uploadFile.ContentLength > 0)
                                {
                                    strFileName = ImgSavePath + newCuDTGeneric.iCUItem + uploadFile.FileName.Substring(uploadFile.FileName.LastIndexOf('.'));
                                    uploadFile.SaveAs(strFilePath + "\\" + strFileName);
                                    newCuDTGeneric.xImgFile = strFileName;
                                }
                            }
                            finally { }
                        }

                        var newFish = new Fish
                        {
                            iCUItem = newCuDTGeneric.iCUItem,
                            season = season,
                            origin = origin,
                            family = family,
                            scientificName = scientificName,
                            commonName = commonName,
                            distributionInTaiwan = distributionInTaiwan,
                            habitats = habitats,
                            reference = reference,
                            characteristic = characteristic,
                            habitatsType = habitatsType,
                            distribution = distribution,
                            utility = utility,
                            sortOrderSpring = int.Parse(sortOrderSpring),
                            sortOrderSummer = int.Parse(sortOrderSummer),
                            sortOrderAutumn = int.Parse(sortOrderAutumn),
                            sortOrderWinter = int.Parse(sortOrderWinter),
                        };
                        _mGIPcoanew_repository.Create<Fish>(newFish);
                        
                        var newCuDTx7 = new CuDTx7
                        {
                            giCuItem = newCuDTGeneric.iCUItem
                        };
                        _mGIPcoanew_repository.Create<CuDTx7>(newCuDTx7);
                        #endregion

                        #region //文章子體
                        string[] titles = { "最新議題", "議題關聯知識文章", "議題關聯知識文章單元順序設定", "議題關聯影音", "資源推薦的超連結", "使用者參與討論或分享心得" };
                        string[] topCats = { "A", "B", "C", "D", "E", "F" };
                        int i = 0;
                        foreach (string x in titles)
                        {
                            var newCuDTGeneric_subitem = new CuDTGeneric
                            {
                                iBaseDSD = 44,
                                iCTUnit = 2200,
                                fCTUPublic = 'Y',
                                sTitle = x,
                                iEditor = "hyweb",
                                iDept = "0",
                                topCat = topCats[i++],
                                showType = '1',
                                siteId = '1',
                                Created_Date = DateTime.Now,
                                dEditDate = DateTime.Now,
                                xPostDate = DateTime.Now,
                                ClickCount = 0,
                                HaveBuildIndex = 0,
                            };
                            _mGIPcoanew_repository.Create<CuDTGeneric>(newCuDTGeneric_subitem);

                            var newCuDTx7_subitem = new CuDTx7
                            {
                                giCuItem = newCuDTGeneric_subitem.iCUItem
                            };
                            _mGIPcoanew_repository.Create<CuDTx7>(newCuDTx7_subitem);

                            var newKnowledgeJigsaw = new KnowledgeJigsaw
                            {
                                parentIcuitem = newCuDTGeneric.iCUItem,
                                gicuitem = newCuDTGeneric_subitem.iCUItem,
                                orderSiteUnit = 1,
                                orderSubject = 2,
                                orderKnowledgeTank = 3,
                                orderKnowledgeHome = 4,
                            };
                            _mGIPcoanew_repository.Create<KnowledgeJigsaw>(newKnowledgeJigsaw);
                        }
                        #endregion

                        ts.Complete();
                    }
                    #endregion
                    Redirect();
                }
                catch (Exception ex) { Response.Write(string.Format("<font size='2' color='red'><strong>發生錯誤：{0}</strong></font>", ex)); }
            }
        }
        else
        {
            if (id != 0)
            {
                #region //Get
                var result = (from p in _mGIPcoanew_repository.List<CuDTGeneric>()
                              join s in _mGIPcoanew_repository.List<Fish>() on p.iCUItem equals s.iCUItem
                              where p.iCUItem == id
                              select new
                              {
                                  p.iCUItem,
                                  p.fCTUPublic,
                                  p.sTitle,
                                  p.xImgFile,
                                  Fish = s,

                              }).FirstOrDefault();

                if (result != null)
                {
                    CropOrFish.Disabled = true;
                    sTitle.Value = result.sTitle;
                    fCTUPublic.Value = result.fCTUPublic.ToString();
                    xImgURL = strWebUrl + strUploadPath + result.xImgFile;
                    sortOrderSpring.Value = result.Fish.sortOrderSpring.ToString();
                    sortOrderSummer.Value = result.Fish.sortOrderSummer.ToString();
                    sortOrderAutumn.Value = result.Fish.sortOrderAutumn.ToString();
                    sortOrderWinter.Value = result.Fish.sortOrderWinter.ToString();
                    season.Value = result.Fish.season;
                    origin.Value = result.Fish.origin;
                    family.Value = result.Fish.family;
                    scientificName.Value = result.Fish.scientificName;
                    commonName.Value = result.Fish.commonName;
                    distributionInTaiwan.Value = result.Fish.distributionInTaiwan;
                    habitats.Value = result.Fish.habitats;
                    reference.Value = result.Fish.reference;
                    characteristic.Value = result.Fish.characteristic;
                    habitatsType.Value = result.Fish.habitatsType;
                    distribution.Value = result.Fish.distribution;
                    utility.Value = result.Fish.utility;
                }
                #endregion
            }
        }
    }
    private string getReturnURL()
    {
        var result = Request.QueryString["returnUrl"] ?? "";
        if (result != "")
            Session.Add("returnUrl", result);
        else
            result = (Session["returnUrl"] ?? "Index.aspx").ToString();

        return result;
    }
    private void Redirect()
    {
        Response.Redirect(returnUrl);
    }
}
