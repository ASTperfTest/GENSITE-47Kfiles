using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;
using System.IO;

/// <summary>
/// Setting 的摘要描述
/// </summary>
public class Setting
{
    public string ctunitid = "";
	public Setting()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}

    public void deletetitle(string strSQL,string strSQL2,string ID,string path1,string xmlpath)
    {
        string getunitid = "select * from cattreenode where ctnodekind='U' and ctrootid=@Rootid";
       
        string getfile = "select * from nodeinfo where ctrootid =@Rootid";
        string getbanner = "select * from nodebanner where ctrootid=@Rootid";
        string delnode = "delete from cattreenode where ctrootid =@Rootid";
        string checkxml = "select pvxdmp from cattreeroot where ctrootid = @Rootid";


        SqlConnection conn = new SqlConnection(ConnectionSettings());
        conn.Open();
       
        SqlCommand cmdgetfile = new SqlCommand(getfile, conn); //抓檔案刪除
        cmdgetfile.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(ID);
        
        SqlDataReader reader = cmdgetfile.ExecuteReader();
        while (reader.Read())
        {
            if (!Convert.IsDBNull(reader["pic"]))
            {
                string path = reader["pic"].ToString();
                System.IO.File.Delete(path1+path);
            }
            if (!Convert.IsDBNull(reader["pic_title"]))
            {
                string path = reader["pic_title"].ToString();
                System.IO.File.Delete(path1 + path);
            }

        }
        reader.Close(); 
        SqlCommand delbanner = new SqlCommand(getbanner, conn);
        delbanner.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(ID);
        SqlDataReader reader2 = delbanner.ExecuteReader(); //抓banner刪除
        while (reader2.Read())
        {
            string path = reader2["pic"].ToString();
            System.IO.File.Delete(path1 + path);

        }
        reader2.Close();
        

        SqlCommand getunit = new SqlCommand(getunitid, conn);//刪除主題單元,代碼,CtunitX??.xml
        getunit.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(ID);
        SqlDataReader unit = getunit.ExecuteReader();
       
        while (unit.Read())
        {
            
            ctunitid += (ctunitid=="")? unit["CtUnitid"].ToString() :","+unit["CtUnitid"].ToString();

           
        }
        unit.Close();
        deleteunitcode(ctunitid,xmlpath);


        SqlCommand delxml = new SqlCommand(checkxml,conn);
        delxml.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(ID);
        string xml = delxml.ExecuteScalar().ToString(); //xdmp刪除
        if (File.Exists(xmlpath+ "xdmp" + xml + ".xml"))
        {
            File.Delete(xmlpath+ "xdmp" + xml + ".xml");
        }
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        SqlCommand cmd2 = new SqlCommand(strSQL2, conn);
        SqlCommand cmd3 = new SqlCommand(delnode, conn);
        cmd3.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(ID);

        cmd.ExecuteNonQuery();
        cmd2.ExecuteNonQuery();
        cmd3.ExecuteNonQuery();
        conn.Close();
        
    }


    public void deleteunitcode(string unitids,string xmlpath)
    {
        SqlConnection conn = new SqlConnection(ConnectionSettings());
        conn.Open();
        string[] unitid = unitids.Split(',');

        for (int i = 0; i < unitid.Length; i++)
        {
            File.Delete(xmlpath + "CtUnitX" + unitid[i].ToString() + ".xml");
            string deletecode = "delete from codemain where codemetaid = @Metaid";
            SqlCommand delcode = new SqlCommand(deletecode, conn);
            delcode.Parameters.Add("@Metaid", SqlDbType.NVarChar).Value = "CustomWebCat_" + unitid[i].ToString();

            string deletecodedef = "delete from codemetadef where codeid = @Metaid";
            SqlCommand delcodedef = new SqlCommand(deletecodedef, conn);
            delcodedef.Parameters.Add("@Metaid", SqlDbType.NVarChar).Value = "CustomWebCat_" + unitid[i].ToString();



            if (unitids != "")
            {
                string delunit = "delete from ctunit where ctunitid = @Unitid";
                SqlCommand deleteunit = new SqlCommand(delunit, conn);
                deleteunit.Parameters.Add("@Unitid", SqlDbType.Int).Value = Convert.ToInt32(unitid[i].ToString());
                deleteunit.ExecuteNonQuery();
            }

            delcode.ExecuteNonQuery();
            delcodedef.ExecuteNonQuery();
         
        }

        conn.Close();
    }

    public string ConnectionSettings()
    {
        return System.Configuration.ConfigurationManager.ConnectionStrings["ConnectionString"].ToString();

    }

    public string Xmlpath()
    {
        return "../../GIPDSD/";
    
    }

    public string Filepath()
    {

        return "./Public/";
        

    }
    public string ImgFilepath()
    {

        return "./Public/";


    }


    public string EditFilepath()
    {

        return "./Public/";
    
    }


    public string Stylepath()
    {
        return "./images/coastyle/";

    }

    public string EditStylepath()
    {
        return "../images/coastyle/";

    }


}
