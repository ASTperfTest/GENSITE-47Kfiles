using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Data;
using System.Collections;
using System.Data.SqlClient;
public partial class recommand_Recommandinfodetail : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
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
}
