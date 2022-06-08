using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;

public partial class bgMusic_bgMusicEdit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request.QueryString["type"] == "Add")
        {
            Lbltitle.Text = "新增背景音樂";
        }
        else
        {
            Lbltitle.Text = "編輯背景音樂";
            Btsave.Text = "確定";
            SqlConnection conn1 = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
            SqlCommand sqlcmd = new SqlCommand();
            conn1.Open();
            sqlcmd.Connection = conn1;
            sqlcmd.CommandText = "Select * from BackgroundMusic where bgMusicID=@bgMusicID";
            sqlcmd.Parameters.Add("@bgMusicID", System.Data.SqlDbType.VarChar);
            sqlcmd.Parameters["@bgMusicID"].Value = Request.QueryString["bgMusicID"].ToString();
            SqlDataReader dr= sqlcmd.ExecuteReader();
            if (dr.Read())
            {
                this.txttitle.Text = dr["Title"].ToString();
                lbfilename.Text = dr["FileName"].ToString();
            }
            dr.Close();
            conn1.Close();
        }
    }
    protected void Btsave_Click(object sender, EventArgs e)
    {
        if (!checkdata())
        {
            return;
        }
        SqlConnection conn1 = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
        SqlCommand sqlcmd = new SqlCommand();
        string fileN = string.Empty;
        string file_name = string.Empty;
        string path = string.Empty;
        string sql = string.Empty;

        try
        {
            conn1.Open();
            sqlcmd.Connection = conn1;
            sqlcmd.Parameters.Clear();
            sqlcmd.Parameters.Add("@bgMusicID", System.Data.SqlDbType.VarChar);
            sqlcmd.Parameters.Add("@Title", System.Data.SqlDbType.NVarChar);
            if (Request.QueryString["type"] == "Add")
            {
                fileN = FUMusic.FileName;
                string subfilename = System.IO.Path.GetExtension(fileN);
                Guid g;
                g = Guid.NewGuid();
                file_name = g.ToString() + subfilename;
                sql = "Insert into BackgroundMusic(bgMusicID,Title) values(@bgMusicID,@Title)";
                sqlcmd.CommandText = sql;
                sqlcmd.Parameters["@bgMusicID"].Value = file_name;
                sqlcmd.Parameters["@Title"].Value = txttitle.Text;
                sqlcmd.ExecuteNonQuery();
                RegisterStartupScript("myAlert", "<script>RetrunMessage('新增成功')</script>");

            }
            else
            {            
                sql = "Update BackgroundMusic set Title=@Title where bgMusicID=@bgMusicID";
                sqlcmd.CommandText = sql;
                sqlcmd.Parameters["@bgMusicID"].Value = Request.QueryString["bgMusicID"].ToString();
                sqlcmd.Parameters["@Title"].Value = txttitle.Text;
                sqlcmd.ExecuteNonQuery();
                RegisterStartupScript("myAlert", "<script>alert('更新成功')</script>");
                file_name = Request.QueryString["bgMusicID"].ToString();

            }
            if (this.FUMusic.HasFile)
            {
                string MusicServer = Server.MapPath("~");
                int ll = MusicServer.IndexOf("ugipsys");
                MusicServer = MusicServer.Substring(0, ll);
                path = MusicServer + "project\\web\\subject\\midi\\";
                fileN = FUMusic.FileName;
                sqlcmd.Parameters.Add("@FileName", System.Data.SqlDbType.VarChar);

                FUMusic.SaveAs(path + file_name);
                sql = "Update BackgroundMusic set FileName=@FileName where bgMusicID=@bgMusicID";
                sqlcmd.CommandText = sql;
                sqlcmd.Parameters["@bgMusicID"].Value = file_name;
                sqlcmd.Parameters["@FileName"].Value = fileN;
                sqlcmd.ExecuteNonQuery();
            }
        }
        catch (Exception ex)
        {
            Response.Write(ex.Message);
        }
        finally
        {
            sqlcmd.Dispose();
            conn1.Close();
        }        
    }
    protected void Btcancel_Click(object sender, EventArgs e)
    {
        Server.Transfer("bgMusicList.aspx");
    }
    protected bool checkdata()
    {
        if (this.txttitle.Text == string.Empty)
        {
            RegisterStartupScript("myAlert", "<script>alert('曲目標題不可為空白')</script>");
            return false;
        }
        if (Request.QueryString["type"] == "Add")
        {
            if (!FUMusic.HasFile)
            {
                RegisterStartupScript("myAlert", "<script>alert('無上傳曲目')</script>");
                return false;
            }
        }
        
        return true;
    }

}
