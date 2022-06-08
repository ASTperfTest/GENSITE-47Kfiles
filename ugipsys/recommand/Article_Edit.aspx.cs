using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Transactions;
using System.Collections;

public partial class Article_Edit : System.Web.UI.Page
{
    protected string errString = string.Empty;
    protected string saveOK = string.Empty;

    protected void Page_Init(object sender, EventArgs e)
    {
        // 標籤
        string sqlTAGsQueryScript = @"SELECT * FROM TAGs ORDER BY createdDate";
        using (var dt = SqlHelper.GetDataTable("ConnString", sqlTAGsQueryScript))
        {
            listTAGs.DataSource = dt;
            listTAGs.DataTextField = "tagName";
            listTAGs.DataValueField = "tagName";
            listTAGs.DataBind();
        }
    }

    protected void Page_load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            myDBinit();
        }
    }

    // 檢查資料填寫狀況
    private bool isDataOK()
    {
        bool isOK = !string.IsNullOrEmpty(txtTitle.Text)
            && !string.IsNullOrEmpty(txtURL.Text)
            && !string.IsNullOrEmpty(txtContent.Text)
            && !string.IsNullOrEmpty(txtSource.Text);
        if (!isOK)
        {
            errString = "資訊填寫不完整，請重新輸入。";
        }
        else
        {
            if (txtTitle.Text.Length > 50)
            {
                isOK = false;
                errString = "標題不可超過50個字元";
                return isOK;
            }
            if (txtURL.Text.Length > 200)
            {
                isOK = false;
                errString = "URL不可超過200個字元";
                return isOK;
            }
            if (txtContent.Text.Length > 300)
            {
                isOK = false;
                errString = "文章內容不可超過300個字元";
                return isOK;
            }
            if (txtSource.Text.Length > 50)
            {
                isOK = false;
                errString = "資料來源不可超過50個字元";
                return isOK;
            }
        }
        return isOK;
    }

    // 儲存送出
    protected void btnOK_Click(object sender, EventArgs e)
    {
        if (isDataOK())
        {
            using (TransactionScope Scope = new TransactionScope())
            {
                try
                {
                    string sqlUpdateScript = @"UPDATE       RecommandContent 
                                       SET          Title = @Title, URL = @URL, aContent = @aContent,
                                                    aEditDate = @EditDate, Source = @Source
                                       WHERE cID = @cID ";

                    SqlHelper.ExecuteNonQuery("ConnString", sqlUpdateScript,
                        DbProviderFactories.CreateParameter("ConnString", "@Title", "@Title", txtTitle.Text),
                        DbProviderFactories.CreateParameter("ConnString", "@URL", "@URL", txtURL.Text),
                        DbProviderFactories.CreateParameter("ConnString", "@aContent", "@aContent", txtContent.Text),
                        DbProviderFactories.CreateParameter("ConnString", "@EditDate", "@EditDate", DateTime.Now.ToString("yyyy/MM/dd")),
                        DbProviderFactories.CreateParameter("ConnString", "@Source", "@Source", txtSource.Text),
                        DbProviderFactories.CreateParameter("ConnString", "@cID", "@cID", Request.QueryString["cID"]));

                    // 標籤雲的修改
                    // 1.先刪除cID所關聯Tags
                    string sqlDeleteScript = "DELETE   FROM    RecommandContent2TAGs   WHERE   cID = @cID";
                    SqlHelper.ExecuteNonQuery("ConnString", sqlDeleteScript,
                          DbProviderFactories.CreateParameter("ConnString", "@cID", "@cID", Request.QueryString["cID"]));
                    // 2.判斷是否有設定TAGs
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
                        // 3.取得新設定的tagID
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
                        //Response.Write(strWhereAgrs);
                        using (var tagReader = SqlHelper.ReturnReader("ConnString", sqlQuerytagIDScript))
                        {
                            while (tagReader.Read())
                            {
                                // 4.利用迴圈寫入Recommand2TAGs
                                string sqlInsertList2TagsScript = @"INSERT INTO RecommandContent2TAGs (cID, tagID) VALUES (@cID, @tagID);";
                                SqlHelper.ExecuteNonQuery("ConnString", sqlInsertList2TagsScript,
                                    DbProviderFactories.CreateParameter("ConnString", "@cID", "@cID", Request.QueryString["cID"]),
                                    DbProviderFactories.CreateParameter("ConnString", "@tagID", "@tagID", tagReader["tagID"]));
                            }
                        }
                    }
                    myDBinit();
                    Scope.Complete();
                    //Response.Redirect("Index.aspx");
                    saveOK = "Y";
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
    }

    protected void myDBinit()
    {
        string sqlQueryScript = @"SELECT * FROM RecommandContent WHERE cID = @cID";
        using (var reader = SqlHelper.ReturnReader("ConnString", sqlQueryScript,
            DbProviderFactories.CreateParameter("ConnString", "@cID", "@cID", Request.QueryString["cID"])))
        {
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    txtTitle.Text = reader["Title"].ToString();
                    txtURL.Text = reader["URL"].ToString();
                    txtContent.Text = reader["aContent"].ToString();
                    txtSource.Text = reader["Source"].ToString();
                    // 讀取標籤
                    string sqlQueryTagNameScript = @"SELECT     A.tagName 
                                                     FROM       TAGs AS A INNER JOIN 
                                                                    RecommandContent2TAGs AS B 
                                                     ON A.tagID = B.tagID WHERE B.cID = @cID";
                    using (var TAGs_reader = SqlHelper.ReturnReader("ConnString", sqlQueryTagNameScript,
                        DbProviderFactories.CreateParameter("ConnString", "@cID", "@cID", Request.QueryString["cID"])))
                    {

                        if (TAGs_reader.HasRows)
                        {
                            while (TAGs_reader.Read())
                            {
                                foreach (ListItem item in listTAGs.Items)
                                {
                                    // DB取出的值可能會被補上空白，所以使用Trim()去頭尾的空白字元
                                    if (item.Text == TAGs_reader["tagName"].ToString().Trim())
                                    {
                                        item.Selected = true;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

}
