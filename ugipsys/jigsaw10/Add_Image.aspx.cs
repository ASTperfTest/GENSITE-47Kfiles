using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing;
using System.Drawing.Imaging;
using GSS.Vitals.COA.Data;
using System.Data;

/// <summary>
/// Added By Leo    2011-06-21      上傳照片及縮圖
/// </summary>
public partial class Add_Image : System.Web.UI.Page
{
    List<String> msg = new List<String>();
    String now = DateTime.Now.ToString("yyyy/MM/dd");

    public void Page_Load(object sender, EventArgs e)
    {
        // 解決 錯誤訊息：「無效的回傳或回呼引數。 主要是GridView 在 Page_Load 只要再一次DataBind()即發生錯誤
        if (!IsPostBack)
        {
            myDBinit();
        }
    }

    // 重新DataBind..
    private void myDBinit()
    {
        String SQLSelectScript = @"SELECT  a.Stitle, a.xImgFile, b.aTitle, b.NFileName, b.xiCuItem, b.ixCuAttach 
                                   FROM CuDTGeneric as a inner join CuDTAttach as b ON a.icuitem = b.xicuitem 
                                   WHERE a.icuitem = @item AND b.bList = 'y'";
        gvView.DataSource = SqlHelper.GetDataTable("ConnString", SQLSelectScript, 
            DbProviderFactories.CreateParameter("ConnString", "@item", "@item", Request.QueryString["item"]));
        
        // 實作分頁的方式
        //DbProviderFactories.CreateParameter("ConnString", "@item", "@item", Request.QueryString["item"])).Paging(1,2);
        String[] gvKey = { "ixCuAttach" };
        gvView.DataKeyNames = gvKey;
        gvView.DataBind();
    }
    // 上傳確認
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        check(fileupload1, txtName);
        if (msg.Count == 0)
        {
            try
            {
                // 指定路徑 ServerMapPath
                String activeDir = Server.MapPath("../project/project/sys/public/Data/jigsaw/");
                String fileName = fileupload1.FileName;
                String tempFileName = String.Empty;
                // 建立編號資料夾
                String SavePath = System.IO.Path.Combine(activeDir, Request["item"]);
                System.IO.Directory.CreateDirectory(SavePath);
                String pathtoCheck = SavePath + "\\" + fileName;

                // 簡易式檢查檔案是否重複，若重複改名
                if (System.IO.File.Exists(pathtoCheck))
                {
                    int intCount = 2;
                    while (System.IO.File.Exists(pathtoCheck))
                    {
                        tempFileName = "(" + intCount.ToString() + ")" + fileName;
                        pathtoCheck = SavePath + "\\" + tempFileName;
                        intCount++;
                    }
                    fileName = tempFileName;
                }
                // 儲存原始檔
                fileupload1.SaveAs(SavePath + "\\" + fileName);


                String SQLInsertCommand = @"INSERT INTO CuDTAttach
                                      (xiCuItem, aTitle, NFileName, aEditDate, bList)
                                      VALUES (@xiCuItem, @aTitle, @NFileName, @aEditDate, 'y')";
                SqlHelper.ExecuteNonQuery("ConnString", SQLInsertCommand,
                    DbProviderFactories.CreateParameter("ConnString", "@xiCuItem", "@xiCuItem", Request["item"]),
                    DbProviderFactories.CreateParameter("ConnString", "@aTitle", "@aTitle", txtName.Text),
                    DbProviderFactories.CreateParameter("ConnString", "@NFileName", "@NFileName", fileupload1.FileName),
                    DbProviderFactories.CreateParameter("ConnString", "@aEditDate", "@aEditDate", now));

                myDBinit();
                txtName.Text = String.Empty;
                //Response.Write(SavePath + "\\" + fileupload1.FileName);
            }
            catch (Exception)
            {
                Response.Write("error");
                throw;
            }
        }
        else
        {
            String errMsg = String.Empty;
            for (int intX = 0; intX < msg.Count; intX++)
            {
                errMsg += msg[intX] + ",";
            }
            Response.Write("<script language='javascript'>alert('" + errMsg + "並未輸入。')</script>");
        }
    }

    // 檢查 FileUpload & TextBox
    protected List<String> check(FileUpload file, TextBox txt)
    {
        if (file.FileName != String.Empty && txt.Text != String.Empty)
            return null;
        else
        {
            if (file.FileName == String.Empty)
                msg.Add("檔案");
            if (txt.Text == String.Empty)
                msg.Add("附件名稱");
        }
        return msg;
    }

    // 刪除所選
    protected void btnDelete_Click(object sender, EventArgs e)
    {
        String SQLDeletScript = "DELETE FROM CuDTAttach WHERE ixCuAttach = @ixCuAttach";
        //List<String> SQLDeleteParametersName = new List<String>();

        for (int intX = 0; intX < this.gvView.Rows.Count; intX++)
        {
            if (((CheckBox)gvView.Rows[intX].FindControl("CheckBox1")).Checked)
            {
                SqlHelper.ExecuteNonQuery("ConnString", SQLDeletScript,
                    DbProviderFactories.CreateParameter("ConnString", "@ixCuAttach", "@ixCuAttach", gvView.DataKeys[intX].Value.ToString()));
                
                // 輸入 DataKey.Value
                //Response.Write(gvView.DataKeys[intX].Value.ToString() + "<br />");
            }
        }
        myDBinit();
    }
}
