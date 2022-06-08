using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Configuration;

public partial class kmactivity_kmwebpuzzle_index : System.Web.UI.Page
{
    protected string token = "";
    string meid = "";
    protected string hosturl = "";
    protected string rr = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        hosturl = Server.UrlEncode(Request.Url.Host);
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        if (Session["memID"] == null || Session["memID"].ToString() == "")
        {
            Response.Write("<script>alert('連線逾時或尚未登入，請登入會員!');window.location.href=\"01.aspx?a=puzzle\";</script>");
            Response.End();
        }
        meid = Session["memID"].ToString().Trim();
        GetRandomString();

        string sql = " select * from ACTIVITY where getdate() between start_time and end_time ";
        var dt = SqlHelper.GetDataTable("PuzzleConnString", sql);
        if (dt.Rows.Count == 0)
        {
                Response.Write("<script>alert('目前不是活動期間，請於活動期間再登入遊戲!');window.location.href=\"01.aspx?a=puzzle\";</script>");
                Response.End();
        }
        sql = @"
        BEGIN TRAN
        if exists( select * from SSo where login_id = @login_id )
        begin
        update SSo set LastActiveTime = getdate() where login_id =  @login_id
        update ACCOUNT set REALNAME = m.REALNAME,NICKNAME=isnull(m.NICKNAME,''),email = m.email,MODIFY_DATE=getdate()
	        from (select * from mGIPcoanew..member  where account = @login_id ) m
	        where ACCOUNT.login_id =  @login_id 
        select Token from SSo where login_id =  @login_id
        end
        else
        begin
        insert into SSo (Token,LOGIN_ID,LastActiveTime)
        select right(sys.fn_VarBinToHexStr(hashbytes('MD5', cast(newid() as varchar(50)))), 32), @login_id,getdate()
	        if not exists( select * from ACCOUNT where login_id =  @login_id)
	        begin
	        insert into ACCOUNT (LOGIN_ID,REALNAME,NICKNAME,EMAIL,Energy,CREATE_DATE,MODIFY_DATE)
	        select account,REALNAME,isnull(NICKNAME,''),EMAIL,0,getdate(),getdate() from mGIPcoanew..member where account =  @login_id
	        end
	        else
	        begin
	        update ACCOUNT set REALNAME = m.REALNAME,NICKNAME=isnull(m.NICKNAME,''),email = m.email,MODIFY_DATE=getdate()
	        from (select * from mGIPcoanew..member  where account =  @login_id ) m
	        where ACCOUNT.login_id =  @login_id
	        end
        select Token from SSo where login_id =  @login_id
        end
        commit
        ";
        var obj = SqlHelper.ReturnScalar("PuzzleConnString", sql,
                DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", meid));
        if (obj != null && !string.IsNullOrEmpty(obj.ToString()))
        {
            token = obj.ToString();
        }
        else
        {
            Response.Write("<script>alert('資料發生錯誤，請按確定導回首頁!');window.location.href=\"01.aspx?a=puzzle\";</script>");
            Response.End();
        }
    }
    protected void Unnamed1_Click(object sender, EventArgs e)
    {
        string sql = "update account set Energy = Energy+50 where LOGIN_ID = @login_id ";
        SqlHelper.ExecuteNonQuery("PuzzleConnString", sql,
               DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", meid));
    }
    static Random rnd;
    private void GetRandomString()
    {
        rnd = new Random();
        rr = rnd.Next(100, 999).ToString();
    }
}