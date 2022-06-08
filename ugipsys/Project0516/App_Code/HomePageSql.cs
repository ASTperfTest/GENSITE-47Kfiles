using System;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;


/// <summary>
/// HomePageSql 的摘要描述
/// </summary>
public class HomePageSql
{
    
    public HomePageSql()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}
    public DataTable NodeItem(int TreeRootId)
    {       
        DataTable dt = new DataTable();
        SqlConnection conn = new SqlConnection(GlobalSetting.ConnectionSettings());
        StringBuilder sql = new StringBuilder();
        sql.Append("select Node.* from CatTreeNode as Node  inner join CatTreeRoot as Root on Node.ctRootId = Root.ctRootId");
        sql.Append(" where Root.ctRootId = @TreeRootId And Node.ctNodeKind ='U'  order by Node.dataParent");
       
        SqlDataAdapter da = new SqlDataAdapter(sql.ToString(),conn);
        da.SelectCommand.Parameters.Add("@TreeRootId", SqlDbType.Int).Value = TreeRootId;        
        try
        {            
            da.Fill(dt);
            DataRow dr = dt.NewRow();
            dr["catName"] = "不顯示";
            dr["ctNodeId"] = (int)0;
            dt.Rows.InsertAt(dr, 0);
            


        }
        catch (Exception e)
        {
            
        }
        finally
        {
            conn.Close();
            da.Dispose();
            conn.Dispose();
        }
        return dt;
        
    }
    
    public void createXdmpList(int xdmpId,object purpose, string editor, object deptId)
    {        
        SqlConnection conn = new SqlConnection(GlobalSetting.ConnectionSettings());
        SqlCommand cmd = new SqlCommand();
        SqlCommand cmd2 = new SqlCommand();
        conn.Open();
        SqlTransaction tran = conn.BeginTransaction();
        cmd.Transaction = tran;
        cmd2.Transaction = tran;
        try
        {
            cmd2.CommandText = "select ctRootName from CatTreeRoot where ctRootId = @xdmpId";
            cmd2.Parameters.AddWithValue("@xdmpId",xdmpId);
            cmd2.Connection = conn;
            string xdmpName = cmd2.ExecuteScalar().ToString();


            cmd.CommandText = "Insert into XdmpList(xdmpId, xdmpName, purpose, editDate, editor, deptId) values(@xdmpId, @xdmpName, @purpose, @editDate, @editor, @deptID)";
            cmd.Connection = conn;
            cmd.Parameters.AddWithValue("@xdmpId",xdmpId);
            //cmd.Parameters.Add("@xdmpId", SqlDbType.Int).Value = xdmpId;
            cmd.Parameters.Add("@xdmpName",SqlDbType.NVarChar).Value= xdmpName;
            cmd.Parameters.AddWithValue("@purpose",purpose);
           
           
            cmd.Parameters.Add("@editDate", SqlDbType.DateTime).Value = DateTime.Now;
            cmd.Parameters.Add("@editor", SqlDbType.NVarChar).Value = editor;
            cmd.Parameters.Add("@deptId", SqlDbType.NVarChar).Value=deptId;
            
            cmd.ExecuteNonQuery();
            tran.Commit();

        }
        catch
        {
            tran.Rollback();

        }
        finally
        {
            conn.Close();
            conn.Dispose();
            cmd.Dispose();
            cmd2.Dispose();
            tran.Dispose();
        }
    }
    public string LayoutPicName(string xdmpId)
    {
        SqlConnection conn = new SqlConnection(GlobalSetting.ConnectionSettings());
        SqlCommand cmd = new SqlCommand();
        string picName = "";
        conn.Open();
        try
        {
            cmd.CommandText = "select list from nodeinfo where ctrootid=@xdmpId";
            cmd.Parameters.Add("@xdmpId", SqlDbType.Int).Value = xdmpId;
            cmd.Connection = conn;
            picName = cmd.ExecuteScalar().ToString();

        }
        catch
        {

        }
        finally
        {
            conn.Close();
            conn.Dispose();
            cmd.Dispose();
        }
        return picName ;
        
    }
    
}
