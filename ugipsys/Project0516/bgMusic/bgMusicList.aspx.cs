using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data.SqlClient;
using System.Configuration;
using System.Text;

public partial class bgMusic_bgMusicList : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            DSBgMusic.ConnectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
            DSBgMusic.SelectCommand = "Select * from BackgroundMusic ";
            rptList.DataSource = DSBgMusic;
            rptList.DataBind();
        }
        
    }
    protected void Repeater_OnItemCommand(object source, RepeaterCommandEventArgs e)
    {
        string bgid = e.CommandArgument.ToString();
        if (e.CommandName == "Edit")
        {
            Server.Transfer("bgMusicEdit.aspx?bgMusicID=" + bgid + "&type=Update");
        }
        if (e.CommandName == "Delete")
        {            
            try
            {
                DSBgMusic.ConnectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
                DSBgMusic.DeleteCommand = "Delete from  BackgroundMusic where bgMusicID='" + bgid + "'";
                DSBgMusic.Delete();
                DSBgMusic.SelectCommand = "Select * from BackgroundMusic ";
                rptList.DataSource = DSBgMusic;
                rptList.DataBind();
                SqlConnection conn1 = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
                SqlCommand sqlcmd = new SqlCommand();
                conn1.Open();
                sqlcmd.Connection = conn1;
                sqlcmd.CommandText = "Update CuDTx7 set bgMusic='Random' where  bgMusic='" + bgid + "'";
                sqlcmd.ExecuteNonQuery();
                conn1.Close();
                string MusicServer = Server.MapPath("~");
                int ll = MusicServer.IndexOf("ugipsys");
                MusicServer = MusicServer.Substring(0, ll);
                string path = MusicServer + "project\\web\\subject\\midi\\";
                System.IO.File.Delete(path + bgid);
            }
            catch(Exception ex)
            {
                Response.Write(ex.Message);
            }
        }
    }
}
