using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data.SqlClient;
using System.Configuration;
using System.Text;

public partial class AutoComplete : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        GetAutoCompleteData();
    }

    private void GetAutoCompleteData()
    {
        string prefixText = Context.Request.QueryString[0];

        StringBuilder sqlString = new StringBuilder();
        using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString))
        {
            SqlCommand cmd = conn.CreateCommand();
            SqlDataReader dr = null;

            sqlString.Append("SELECT ");
            sqlString.Append("	CatTreeRoot.CtRootName ");
            sqlString.Append("FROM NodeInfo  ");
            sqlString.Append("INNER JOIN CatTreeRoot ON CatTreeRoot.CtRootID = NodeInfo.CtrootID  ");
            sqlString.Append("INNER JOIN [Type] ON NodeInfo.type1 = Type.classname  ");
            //sqlString.Append("WHERE (CatTreeRoot.inUse = 'Y')  ");
            sqlString.Append("where (old_subject = 'N' or old_subject is null ) ");
            sqlString.Append("and (CatTreeRoot.CtRootName like '%" + prefixText + "%') ");
            sqlString.Append("order by  ");
            sqlString.Append("	(Case When NodeInfo.order_num is null Then 1 Else 0 End) ");
            sqlString.Append("	, order_num DESC ");
            sqlString.Append("	, CatTreeRoot.CtrootID  ");

            try
            {
                conn.Open();
                cmd.CommandText = sqlString.ToString();
                sqlString.Remove(0, sqlString.Length);
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string item = dr["CtRootName"].ToString() + Environment.NewLine;
                    Response.Write(item);
                }
            }
            catch { }
        }
    }
}
