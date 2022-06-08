using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Collections;
using System.Transactions;
using GSS.Vitals.COA.Data;

public partial class Recommand_Add : System.Web.UI.Page
{
    string MemberId;
    string Script = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='/';</script>";
    protected string errString = string.Empty;
    protected string saveOK = string.Empty;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["memID"] == null || string.IsNullOrEmpty(Session["memID"].ToString()))
        {
            Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Noid", Script);
            return;
        }
        else
        {
            MemberId = Session["memID"].ToString();
            if (!IsPostBack)
            {
                // 標籤
                string sqlQueryScript = @"SELECT * FROM TAGs ORDER BY createdDate";
                using (var dt = SqlHelper.GetDataTable("ODBCDSN", sqlQueryScript))
                {
                    listTAGs.DataSource = dt;
                    listTAGs.DataTextField = "tagName";
                    listTAGs.DataValueField = "tagName";
                    listTAGs.DataBind();
                }
            }
        }

        string sql;
        int rkey = int.Parse(System.Web.Configuration.WebConfigurationManager.AppSettings["Recommandkey"].ToString());
        sql = @"SELECT *   FROM [CuDTGeneric] 
                where icuitem=" + rkey;
        var dr = SqlHelper.ReturnReader("ConnString", sql);
        if (dr.Read())
        {
            lbtitle.Text = dr["sTitle"].ToString();
            lbxbody.Text = dr["xBody"].ToString();
            Image1.ImageUrl = System.Web.Configuration.WebConfigurationManager.AppSettings["WWWUrl"] + "/public/data/" + dr["xImgFile"].ToString();
        }        
    }


    bool IsDataOK()
    {

        bool isok = !string.IsNullOrEmpty(txtUrl.Text)
            && !string.IsNullOrEmpty(txtTitle.Text)
            && !string.IsNullOrEmpty(txtSource.Text)
            && !string.IsNullOrEmpty(txtContent.Text);
        
        if (!isok)
        {
            errString = "資訊填寫不完整，請重新輸入。";
        }
        else
        {

            foreach (ListItem item in listTAGs.Items)
            {
                if (item.Selected)
                {
                    errString = "";
                    isok = true;
                    break;
                }
                errString = "請選擇分類。";
                isok = false;
            }



            if (txtTitle.Text.Length > 50)
            {
                isok = false;
                errString = "文章標題，不可超過50個字元";
                return isok;
            }
            if (txtUrl.Text.Length > 200)
            {
                isok = false;
                errString = "推薦好文網址，不可超過200個字元";
                return isok;
            }

            if (txtSource.Text.Length > 50)
            {
                isok = false;
                errString = "資料出處，不可超過50個字元";
                return isok;
            }
            if (txtContent.Text.Length > 300 || txtContent.Text.Length < 20)
            {
                isok = false;
                errString = "推薦原因，必須介於20至300個字元";
                return isok;
            }

            // 檢查推薦好文的網址
            string sql = @"SELECT URL
                             FROM RecommandContent 
                             WHERE URL = '" + txtUrl.Text.Trim()+ "'";
            try
            {
                if (SqlHelper.ReturnScalar("ConnString", sql) != null)
                {
                    isok = false;
                    errString = "此網址已被推薦";
                    return isok;
                }
            }
            catch(Exception) { }

        }
        return isok;
    }

    protected void btnConfirm_Click(object sender, EventArgs e)
    {
        if (IsDataOK())
        {
            if (!txtUrl.Text.ToLower().StartsWith("http://"))
            {
                txtUrl.Text = "http://" + txtUrl.Text;
            }

            // 使用TransactionScope確保交易成功
            using (TransactionScope scope = new TransactionScope())
            {
                try
                {   
                    string sqlInsertLisScript = @"
                       set nocount on 
                       INSERT INTO RecommandContent (URL, Title, aEditDate, aContent, iEditor, created_Date, Source) 
                       VALUES (@URL, @Title, getdate(), @aContent, @iEditor, getdate(), @Source)
                       select SCOPE_IDENTITY()";

                    var cId = SqlHelper.ReturnScalar("ODBCDSN", sqlInsertLisScript,
                         DbProviderFactories.CreateParameter("ODBCDSN", "@URL", "@URL", txtUrl.Text),
                         DbProviderFactories.CreateParameter("ODBCDSN", "@Title", "@Title", txtTitle.Text),
                         DbProviderFactories.CreateParameter("ODBCDSN", "@aContent", "@aContent", txtContent.Text),
                         DbProviderFactories.CreateParameter("ODBCDSN", "@iEditor", "@iEditor", MemberId),
                         DbProviderFactories.CreateParameter("ODBCDSN", "@Source", "@Source", txtSource.Text));


                    // 判斷文章是否貼上標籤
                    ArrayList arrTAGs = new ArrayList();
                    foreach (ListItem item in listTAGs.Items)
                    {
                        if (item.Selected)
                        {
                            arrTAGs.Add(item.Text);
                        }
                    }
                    if (arrTAGs.Count > 0)
                    {

                        // 1.取得此篇文章的tagID
                        string strWhereAgrs = string.Empty;
                        for (int i = 0; i < arrTAGs.Count; i++)
                        {
                            if (i == 0)
                            {
                                strWhereAgrs = " tagName = '" + arrTAGs[i] + "'";
                            }
                            else
                            {
                                strWhereAgrs += " OR tagName = '" + arrTAGs[i] + "'";
                            }
                        }
                        string sqlQuerytagIDScript = @"SELECT tagID FROM TAGs WHERE tagName
                                                           IN  (SELECT tagName FROM TAGs WHERE " + strWhereAgrs + ")";

                        //Response.Write(sqlQuerytagIDScript);
                        using (var tagReader = SqlHelper.ReturnReader("ODBCDSN", sqlQuerytagIDScript))
                        {
                            while (tagReader.Read())
                            {
                                // 3.利用迴圈寫入Recommand2TAGs
                                string sqlInsertList2TagsScript = @"INSERT INTO RecommandContent2TAGs (cID, tagID) VALUES (@cID, @tagID);";
                                SqlHelper.ExecuteNonQuery("ODBCDSN", sqlInsertList2TagsScript,
                                    DbProviderFactories.CreateParameter("ODBCDSN", "@cID", "@cID", cId),
                                    DbProviderFactories.CreateParameter("ODBCDSN", "@tagID", "@tagID", tagReader["tagID"]));
                            }
                        }

                    }
                    scope.Complete();
                    saveOK = "Y";
                }
                catch (Exception)
                {
                    //throw;
                }
                
            }
        }
    }
}