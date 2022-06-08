using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Transactions;
using System.Data.Linq;

public partial class EditCrop : System.Web.UI.Page
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
            string type = Request.Form["type"];
            string description = Request.Form["description"];
            string alias = Request.Form["alias"];
            string season = Request.Form["season"];
            string origin = Request.Form["origin"];
            string feature = Request.Form["feature"];
            string nutritionValue = Request.Form["nutritionValue"];
            string selectionMethod = Request.Form["selectionMethod"];
            string nutrient = Request.Form["nutrient"];
            string cuisine = Request.Form["cuisine"];
            string variety = Request.Form["variety"];
            string note = Request.Form["note"];
            string tips = Request.Form["tips"];

            if (string.IsNullOrEmpty(sTitle))
            {
                StrFunc.GenErrMsg("請輸入作物名稱！", string.Format("Javascript:window.history.back();"));
            }
            if (string.IsNullOrEmpty(type))
            {
                StrFunc.GenErrMsg("請輸入作物分類！", string.Format("Javascript:window.history.back();"));
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
                var originCrop = _mGIPcoanew_repository.Get<Crop>(p => p.iCUItem == id);
                if (originCuDTGeneric != null && originCrop != null)
                {
                    try
                    {
                        if (Request.Form["action_equals_Del"] == "1")
                        {
                            #region //DELETE
                            string strSql = "delete from CuDTGeneric where iCUItem=" + StrFunc.NZQty(id.ToString()) + ";" +
                                            "delete from Crop where iCUItem=" + StrFunc.NZQty(id.ToString()) + ";";

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
                                originCrop.type = Byte.Parse(type);
                                originCrop.description = description;
                                originCrop.alias = alias;
                                originCrop.season = season;
                                originCrop.origin = origin;
                                originCrop.feature = feature;
                                originCrop.nutritionValue = nutritionValue;
                                originCrop.selectionMethod = selectionMethod;
                                originCrop.nutrient = nutrient;
                                originCrop.cuisine = cuisine;
                                originCrop.variety = variety;
                                originCrop.note = note;
                                originCrop.tips = tips;
                                originCrop.sortOrderSpring = int.Parse(sortOrderSpring);
                                originCrop.sortOrderSummer = int.Parse(sortOrderSummer);
                                originCrop.sortOrderAutumn = int.Parse(sortOrderAutumn);
                                originCrop.sortOrderWinter = int.Parse(sortOrderWinter);

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
                        
                        var newCrop = new Crop
                        {
                            iCUItem = newCuDTGeneric.iCUItem,
                            type = Byte.Parse(type),
                            description = description,
                            alias = alias,
                            season = season,
                            origin = origin,
                            feature = feature,
                            nutritionValue = nutritionValue,
                            selectionMethod = selectionMethod,
                            nutrient = nutrient,
                            cuisine = cuisine,
                            variety = variety,
                            note = note,
                            tips = tips,
                            sortOrderSpring = int.Parse(sortOrderSpring),
                            sortOrderSummer = int.Parse(sortOrderSummer),
                            sortOrderAutumn = int.Parse(sortOrderAutumn),
                            sortOrderWinter = int.Parse(sortOrderWinter),
                        };
                        _mGIPcoanew_repository.Create<Crop>(newCrop);

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
                catch(Exception ex) { Response.Write(string.Format("<font size='2' color='red'><strong>發生錯誤：{0}</strong></font>", ex)); }
            }
        }
        else
        {
            if (id != 0)
            {
                #region //Get
                var result = (from p in _mGIPcoanew_repository.List<CuDTGeneric>()
                              join s in _mGIPcoanew_repository.List<Crop>() on p.iCUItem equals s.iCUItem
                              where p.iCUItem == id
                              select new
                              {
                                  p.iCUItem,
                                  p.fCTUPublic,
                                  p.sTitle,
                                  p.xImgFile,
                                  Crop = s,
                              }).FirstOrDefault();

                if (result != null)
                {
                    CropOrFish.Disabled = true;
                    sTitle.Value = result.sTitle;
                    fCTUPublic.Value = result.fCTUPublic.ToString();
                    xImgURL = strWebUrl + strUploadPath + result.xImgFile;
                    sortOrderSpring.Value = result.Crop.sortOrderSpring.ToString();
                    sortOrderSummer.Value = result.Crop.sortOrderSummer.ToString();
                    sortOrderAutumn.Value = result.Crop.sortOrderAutumn.ToString();
                    sortOrderWinter.Value = result.Crop.sortOrderWinter.ToString();
                    type.Value = result.Crop.type.ToString();
                    description.Value = result.Crop.description;
                    alias.Value = result.Crop.alias;
                    season.Value = result.Crop.season;
                    origin.Value = result.Crop.origin;
                    feature.Value = result.Crop.feature;
                    nutritionValue.Value = result.Crop.nutritionValue;
                    selectionMethod.Value = result.Crop.selectionMethod;
                    nutrient.Value = result.Crop.nutrient;
                    cuisine.Value = result.Crop.cuisine;
                    variety.Value = result.Crop.variety;
                    note.Value = result.Crop.note;
                    tips.Value = result.Crop.tips;
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
