using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using GSS.Vitals.COA.Data;
using Newtonsoft.Json;

/// <summary>
/// Added By Leo    2011-06-20      jQuery SlideShow' Page.
/// </summary>
public partial class FileShowPage : System.Web.UI.Page
{
    protected string result = string.Empty;

    protected void Page_Load(object sender, EventArgs e)
    {
        bool isOK;
        int item = 0;
        List<imageItem> totalImageItem = new List<imageItem>();

        isOK = int.TryParse(Request["item"], out item);
        if (!isOK) isOK = int.TryParse(Request["item"], out item);
        if (isOK)
        {
            // 取得附件的圖片  Start
            string sqlString = @"SELECT A.xImgFile, A.sTitle, B.NFileName, B.aTitle
                                 FROM CuDTGeneric AS A INNER JOIN CuDTAttach AS B ON A.iCuItem = B.xiCuItem
                                 WHERE A.icuitem = @item AND b.bList = 'y'";

            using (var reader = SqlHelper.ReturnReader("ODBCDSN", sqlString,
                DbProviderFactories.CreateParameter("ODBCDSN", "@item", "@item", item)))
            {
                while (reader.Read())
                {
                    imageItem theitem = new imageItem();
                    theitem.image = "/public/Data/jigsaw/" + Request.QueryString["item"] + "/" + reader["NFileName"].ToString();
                    theitem.title = reader["aTitle"].ToString();
                    theitem.url = "";
                    totalImageItem.Add(theitem);
                }
                // 取得附件的圖片  End

                // 取得原始的圖片  Start
                /*
                string sqlString2 = @"SELECT     sTitle, xImgFile
                                      FROM       CuDTGeneric
                                      WHERE     (iCUItem = @iCUItem)";

                using (var reader1 = SqlHelper.ReturnReader("ODBCDSN", sqlString2,
                    DbProviderFactories.CreateParameter("ODBCDSN", "@iCUItem", "@iCUItem", item)))
                {
                    while (reader1.Read())
                    {
                        imageItem OriginalImage = new imageItem();
                        OriginalImage.image = "/public/Data/" + reader1["xImgFile"].ToString();
                        OriginalImage.title = reader1["sTitle"].ToString();
                        OriginalImage.url = "";
                        totalImageItem.Add(OriginalImage);
                    }
                }

                if (totalImageItem.Count == 0)
                {
                    this.Visible = false;
                    totalImageItem.Add(new imageItem());
                }
                */ 
                // 取得原始的圖片  End
                result = Newtonsoft.Json.JsonConvert.SerializeObject(totalImageItem);
            }
        }
        //List<imageItem> totalImageItem = new List<imageItem>();
        //for (int intX = 1; intX < 4; intX++)
        //{
        //    imageItem xitem = new imageItem();
        //    xitem.image = "cat0" + intX + ".jpg";
        //    xitem.title = "城市的野貓0" + intX;
        //    xitem.url = "";
        //    totalImageItem.Add(xitem);
        //}
        //result = Newtonsoft.Json.JsonConvert.SerializeObject(totalImageItem);
    }

    private class imageItem
    {
        public string image { get; set; }
        public string title { get; set; }
        public string url { get; set; }
    }
}