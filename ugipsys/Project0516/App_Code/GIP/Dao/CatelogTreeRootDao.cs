using System;
using System.Data;
using System.Data.SqlClient;

/// <summary>
/// CatelogTreeRootDao 的摘要描述
/// </summary>
public class CatelogTreeRootDao
{
	private static CatelogTreeRootDao _instance;

	public CatelogTreeRootDao()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}

	public static CatelogTreeRootDao getInstance()
	{
		if (_instance == null)
			_instance = new CatelogTreeRootDao();
		return _instance;
	}

	public Object findById(Object id)
	{
		using (SqlConnection conn = (SqlConnection)SqlDbHelper.getInstance().getConnection())
		{
			conn.Open();
			using (SqlCommand command = new SqlCommand())
			{
				command.CommandText = "SELECT * FROM CatTreeRoot WHERE 1 = 1 AND ctRootId = @ctRootId";
				command.Connection = conn;
				command.Parameters.AddWithValue("@ctRootId", Convert.ToInt32(id));
				IDataReader reader = command.ExecuteReader();
				if (reader.Read())
				{
					return populateFromReader(reader);
				}
			}
		}

		return null;
	}

	public void insert(Object vo)
	{
	}

	public void update(Object vo)
	{
	}

	public void delete(Object vo)
	{
	}

	public CatelogTreeRoot populateFromReader(IDataReader reader)
	{
		CatelogTreeRoot root = new CatelogTreeRoot();

		root.Id = Convert.ToInt16(reader["ctRootId"]);
		root.Name = Convert.ToString(reader["ctRootName"]);
		root.InUse = Convert.ToString(reader["inUse"]).Equals("Y");
		root.ModifyUser = Convert.ToString(reader["editor"]);
		root.ModifyDate = Convert.ToDateTime(reader["editDate"]);
		
		return root;
	}
}
