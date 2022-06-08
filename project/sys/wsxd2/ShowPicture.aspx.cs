using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;
using System.Web.Configuration;
using AjaxControlToolkit;

/// <summary>
/// 圖片瀏覽
/// </summary>
public partial class ShowPicture : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            SetPictureDropDownListDataSource();
            GetPictureData(Request.QueryString["fileName"]);
            //ImageShow.Attributes.Add("onload", "if (this.height>500) this.height=500;");
        }
    }

    private string ConvertToDBString(string orignalString)
    {
        return "'" + orignalString + "'";
    }
    protected void PreImageButton_Click(object sender, ImageClickEventArgs e)
    {
        int dlCount = ImageDataList.Items.Count;
		int rownumber = int.Parse(((HiddenField)FindControl("HiddenFieldRowNumber")).Value);
        if (dlCount > 2)
        {
            TitleShow.Text = ((Label)ImageDataList.Items[rownumber-1].FindControl("LabelTitle")).Text;
            string fileName = ((HiddenField)ImageDataList.Items[rownumber-1].FindControl("HiddenFieldFileName")).Value;
            GetPictureData(fileName);
        }
    }

    private void MarkPicture()
    {
        for (int index = 0; index < ImageDataList.Items.Count; index++)
        {
            ImageButton imgBtn = ((ImageButton)ImageDataList.Items[index].FindControl("ImageButton"));
            if (ImageShow.ImageUrl == imgBtn.ImageUrl)
            {
				HiddenFieldRowNumber.Value = index.ToString();
                imgBtn.BorderColor = System.Drawing.Color.Red;
                imgBtn.BorderWidth = 3;
                imgBtn.BorderStyle = BorderStyle.Dashed;
                break;
            }
        }
    }

    protected void NextImageButton_Click(object sender, ImageClickEventArgs e)
    {
        int dlCount = ImageDataList.Items.Count;
		int rownumber = int.Parse(((HiddenField)FindControl("HiddenFieldRowNumber")).Value);
        if (dlCount > 2)
        {
            TitleShow.Text = ((Label)ImageDataList.Items[rownumber + 1].FindControl("LabelTitle")).Text;
            string fileName = ((HiddenField)ImageDataList.Items[rownumber + 1].FindControl("HiddenFieldFileName")).Value;
            GetPictureData(fileName);

        }
    }

    protected void ImageDataList_SelectedIndexChanged(object sender, EventArgs e)
    {
        ImageButton theImage =
            (ImageButton)ImageDataList.Items[ImageDataList.SelectedIndex].FindControl("ImageButton");
        ImageShow.ImageUrl = theImage.ImageUrl.ToString();
        int indexBegin = theImage.ImageUrl.ToString().LastIndexOf("/");
        string fileName = theImage.ImageUrl.ToString().Substring(indexBegin);

        TitleShow.Text = ((Label)ImageDataList.Items[ImageDataList.SelectedIndex].FindControl("LabelTitle")).Text;
        GetPictureData(fileName.Replace("/", ""));
    }
    /// <summary>
    /// 調整圖片比例大小
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void ImageDataList_ItemDataBound(object sender, DataListItemEventArgs e)
    {
        ImageButton img = (ImageButton)e.Item.FindControl("ImageButton");
        img.Attributes.Add("onload", "if (this.width>80){ this.width=80;} ");
    }

    /// <summary>
    /// Gets the picture data.
    /// </summary>
    /// <param name="inputFileName">Name of the input file.</param>
    private void GetPictureData(string inputFileName)
    {
        string fileName = inputFileName;
        string pathType = Request.QueryString["PathType"].ToUpper();
        string sqlParameter = string.Empty;

        int totalCount = 0;
        int picRowNum = 0;

        string imgPath = string.Empty;
        string picCountSql = string.Empty;
        string picRowNumsSql = string.Empty;
        string select5PicsSql = string.Empty;
        DataTable dt = new DataTable();
        DataTable picDataTable = new DataTable();

        string connString =
                WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;

        if (pathType == "A")
        {
            imgPath = "../public/Attachment/";
            sqlParameter = Request.QueryString["CuItem"];
            picCountSql = @"select count(1) as total from CuDtAttach where bList='Y' and xiCUItem=@sqlParameter
and ( NFileName like  '%.jpg' or NFileName like '%.jpge' or NFileName like '%.gif' or NFileName like '%.png' or NFileName like '%.bmp')";
            picRowNumsSql = GetPictureRowNumsSqlByUItem();
            select5PicsSql = Get5PictureSelectSqlByUitem(imgPath);

        }
        else if (pathType == "D")
        {
            imgPath = "../public/Data/";
            sqlParameter = Request.QueryString["CtNodeID"];
            picCountSql = GetPictureCountSqlByNoteId();
            picRowNumsSql = GetPictureRowNumsSqlByNoteId();
            select5PicsSql = Get5PictureSelectSqlByNoteId(imgPath);
        }

        using (SqlConnection cn = new SqlConnection())
        {
            cn.ConnectionString = connString;
            cn.Open();

            using (SqlCommand picCountCmd = new SqlCommand(picCountSql, cn))
            {
                picCountCmd.Parameters.AddWithValue("@sqlParameter", sqlParameter);
                totalCount = (int)picCountCmd.ExecuteScalar();
            }

            using (SqlCommand picRowNumsCmd = new SqlCommand(picRowNumsSql, cn))
            {
                picRowNumsCmd.Parameters.AddWithValue("@sqlParameter", sqlParameter);
                picRowNumsCmd.Parameters.AddWithValue("@fileName", fileName);

                new SqlDataAdapter(picRowNumsCmd).Fill(dt);
            }

            if (dt.Rows.Count != 0)
            {
                picRowNum = Convert.ToInt32(dt.Rows[0]["rowNum"]);
                TitleShow.Text = dt.Rows[0]["aTitle"].ToString();
                ddlPicture.SelectedIndex = picRowNum - 1;

                if (totalCount > 5 && picRowNum > 2 && picRowNum <= (totalCount - 3))
                    select5PicsSql = select5PicsSql + " where tt.rowNum between " + (picRowNum - 2) + " and " + (picRowNum + 2);
                else if (picRowNum > (totalCount - 3))
                    select5PicsSql = select5PicsSql + " where tt.rowNum between " + (totalCount - 4) + " and " + (totalCount);
                else if (totalCount > 5 && picRowNum < 3)
                    select5PicsSql = select5PicsSql + " where tt.rowNum between 1 and 5 ";

                using (SqlCommand pic5Cmd = new SqlCommand(select5PicsSql, cn))
                {
                    pic5Cmd.Parameters.AddWithValue("@sqlParameter", sqlParameter);
                    new SqlDataAdapter(pic5Cmd).Fill(picDataTable);
                }
            }
        }

        ImageShow.ImageUrl = imgPath + fileName;

        ImageDataList.DataSource = picDataTable;
        ImageDataList.DataBind();

        LabelPicTotal.Text = totalCount.ToString();

        //處理按鈕的顯示
        if (totalCount > 5)
        {
            if (picRowNum < totalCount )
                NextImageButton.Visible = true;
            else
                NextImageButton.Visible = false;

            if (picRowNum <= 1)
                PreImageButton.Visible = false;
            else
                PreImageButton.Visible = true;
        }

        MarkPicture();
    }

    #region 字串給值
    private string Get5PictureSelectSqlByNoteId(string imgPath)
    {
        return @"with tt as 
	        (
		        SELECT row_number() over (ORDER BY htx.iCUItem ) as rowNum,
                        xImgFile as NFileName,sTitle as aTitle
		        FROM CuDTGeneric AS htx  
		        JOIN CtUnit AS u ON u.CtUnitID=htx.iCtUnit 
		        JOIN CatTreeNode AS n ON n.CtUnitID=htx.iCtUnit 
		        LEFT JOIN dept AS xr1 ON xr1.deptID=htx.iDept 
		        WHERE htx.fCTUPublic='Y'  
		        AND (htx.avEnd is NULL OR htx.avEnd >=GetDate()) 
		        AND (htx.avBegin is NULL OR htx.avBegin <=GetDate()) 
		        AND n.CtNodeID IN (@sqlParameter) 
                AND ( xImgFile like  '%.jpg' or xImgFile like '%.jpge' or xImgFile like '%.gif' or xImgFile like '%.png' or xImgFile like '%.bmp')
	        )
	        select * ," + ConvertToDBString(imgPath) + @"+ ISNUll(tt.NFileName,'') as imgUrl 
	        from tt";
    }

    private static string GetPictureRowNumsSqlByNoteId()
    {
        return @"with tt as 
                (
			        SELECT row_number() over (ORDER BY htx.iCUItem ) as rowNum,xImgFile as NFileName,sTitle as aTitle
			        FROM CuDTGeneric AS htx  
			        JOIN CtUnit AS u ON u.CtUnitID=htx.iCtUnit 
			        JOIN CatTreeNode AS n ON n.CtUnitID=htx.iCtUnit 
			        LEFT JOIN dept AS xr1 ON xr1.deptID=htx.iDept 
			        WHERE htx.fCTUPublic='Y'  
			        AND (htx.avEnd is NULL OR htx.avEnd >=GetDate()) 
			        AND (htx.avBegin is NULL OR htx.avBegin <=GetDate()) 
			        AND n.CtNodeID IN (@sqlParameter) 
                    AND ( xImgFile like  '%.jpg' or xImgFile like '%.jpge' or xImgFile like '%.gif' or xImgFile like '%.png' or xImgFile like '%.bmp')
                )
			    select rowNum,aTitle 
			    from tt
			    where NFileName = @fileName";
    }

    private static string GetPictureCountSqlByNoteId()
    {
        return @"SELECT count(1) as total
				FROM CuDTGeneric AS htx  
				JOIN CtUnit AS u ON u.CtUnitID=htx.iCtUnit 
				JOIN CatTreeNode AS n ON n.CtUnitID=htx.iCtUnit 
				LEFT JOIN dept AS xr1 ON xr1.deptID=htx.iDept 
				WHERE htx.fCTUPublic='Y'  
				AND (htx.avEnd is NULL OR htx.avEnd >=GetDate()) 
				AND (htx.avBegin is NULL OR htx.avBegin <=GetDate()) 
				AND n.CtNodeID IN (@sqlParameter)
                AND ( xImgFile like  '%.jpg' or xImgFile like '%.jpge' or xImgFile like '%.gif' or xImgFile like '%.png' or xImgFile like '%.bmp')";
    }

    private string Get5PictureSelectSqlByUitem(string imgPath)
    {
        return @"with tt as
			(
				SELECT 
				row_number() over (ORDER BY dhtx.listSeq ) as rowNum,dhtx.NFileName,dhtx.aTitle
				FROM CuDtAttach AS dhtx 
				WHERE bList='Y' 
				AND dhtx.xiCUItem=@sqlParameter
                AND ( NFileName like  '%.jpg' or NFileName like '%.jpge' or NFileName like '%.gif' or NFileName like '%.png' or NFileName like '%.bmp')

			)
			select * ," + ConvertToDBString(imgPath) + @"+ ISNUll(tt.NFileName,'') as imgUrl from
			tt";
    }

    private static string GetPictureRowNumsSqlByUItem()
    {
        return @"with tt as
			    (
				    SELECT 
				    row_number() over (ORDER BY dhtx.listSeq ) as rowNum,dhtx.NFileName,dhtx.aTitle
				    FROM CuDtAttach AS dhtx 
				    WHERE bList='Y' 
				    AND dhtx.xiCUItem=@sqlParameter
                    AND ( NFileName like  '%.jpg' or NFileName like '%.jpge' or NFileName like '%.gif' or NFileName like '%.png' or NFileName like '%.bmp')
			    )
			    select rowNum,aTitle from
			    tt
			    where tt.NFileName = @fileName";
    }
    #endregion      

    private void SetPictureDropDownListDataSource()
    {
        string xiCUItemParemeter = Request.QueryString["CuItem"];
        string xCtNodeIDParemeter = Request.QueryString["CtNodeID"];
        string pathType = Request.QueryString["PathType"].ToUpper();
        string connString =
               WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;
        DataTable dt = new DataTable();

        using (SqlConnection cn = new SqlConnection())
        {
            cn.ConnectionString = connString;
            cn.Open();
            string selectCount = "";

            if (pathType == "A")
            {
                selectCount = @"select row_number() over (ORDER BY listSeq ) as rowNum,NFileName from CuDtAttach where bList='Y' and xiCUItem=@xiCUItemParemeter
                and ( NFileName like  '%.jpg' or NFileName like '%.jpge' or NFileName like '%.gif' or NFileName like '%.png' or NFileName like '%.bmp')";
               
                using (SqlCommand selectCountCmd = new SqlCommand(selectCount, cn))
                {
                    selectCountCmd.Parameters.AddWithValue("@xiCUItemParemeter", xiCUItemParemeter);
                    new SqlDataAdapter(selectCountCmd).Fill(dt);
                }
            }
            else if (pathType == "D")
            {
                selectCount = @"SELECT row_number() over (ORDER BY htx.iCUItem ) as rowNum,xImgFile as NFileName 
				FROM CuDTGeneric AS htx  
				JOIN CtUnit AS u ON u.CtUnitID=htx.iCtUnit 
				JOIN CatTreeNode AS n ON n.CtUnitID=htx.iCtUnit 
				LEFT JOIN dept AS xr1 ON xr1.deptID=htx.iDept 
				WHERE htx.fCTUPublic='Y'  
				AND (htx.avEnd is NULL OR htx.avEnd >=GetDate()) 
				AND (htx.avBegin is NULL OR htx.avBegin <=GetDate()) 
				AND n.CtNodeID IN (@xCtNodeIDParemeter)
                AND ( xImgFile like  '%.jpg' or xImgFile like '%.jpge' or xImgFile like '%.gif' or xImgFile like '%.png' or xImgFile like '%.bmp')";

                using (SqlCommand selectCountCmd = new SqlCommand(selectCount, cn))
                {
                    selectCountCmd.Parameters.AddWithValue("@xCtNodeIDParemeter", xCtNodeIDParemeter);
                    new SqlDataAdapter(selectCountCmd).Fill(dt);
                }
            }
            
            ddlPicture.DataSource = dt;
            ddlPicture.DataValueField = "NFileName";
            ddlPicture.DataTextField = "rowNum";
            ddlPicture.DataBind();
        }
    }

    protected void DDLPicture_SelectedIndexChanged(object sender, EventArgs e)
    {
        string fileName = ddlPicture.SelectedValue;
        GetPictureData(fileName);
    }
 
}
