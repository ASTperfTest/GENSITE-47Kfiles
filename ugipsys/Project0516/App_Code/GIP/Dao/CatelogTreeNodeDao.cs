using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections;

/// <summary>
/// CatelogTreeNodeDao 的摘要描述
/// </summary>
public class CatelogTreeNodeDao
{
	private static CatelogTreeNodeDao _instance;

	public CatelogTreeNodeDao() {}

	public static CatelogTreeNodeDao getInstance()
	{
		if (_instance == null)
			_instance = new CatelogTreeNodeDao();
		return _instance;
	}

	public Object findById(Object id)
	{
		using (SqlConnection conn = (SqlConnection)SqlDbHelper.getInstance().getConnection())
		{
			conn.Open();
			using (SqlCommand command = new SqlCommand())
			{
				command.CommandText = "SELECT * FROM CatTreeNode WHERE 1 = 1 AND ctNodeId = @ctNodeId";
				command.Connection = conn;
				command.Parameters.AddWithValue("@ctNodeId", Convert.ToInt32(id));
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
		CatelogTreeNode obj = (CatelogTreeNode)vo;

		using (SqlConnection conn = (SqlConnection)SqlDbHelper.getInstance().getConnection())
		{
			conn.Open();
			using (SqlCommand command = new SqlCommand())
			{
				command.CommandText = "INSERT INTO CatTreeNode (ctRootId, ctNodeKind, catName, ctNameLogo, catShowOrder, dataLevel, dataParent, childCount, dataCount, ctUnitId, inUse, editDate, editUserId,  xslList, xslData, ctNodeNpkind, rss,catNameMemo) VALUES (@ctRootId, @ctNodeKind, @catName, @ctNameLogo, @catShowOrder, @dataLevel, @dataParent, @childCount, @dataCount, @ctUnitId, @inUse, @editDate, @editUserId,  @xslList, @xslData, @ctNodeNpkind, @rss,@catNameMemo)";
				command.CommandText += "\n" + "SELECT @Id = SCOPE_IDENTITY()";
				command.Connection = conn;

				command.Parameters.Add("@Id", SqlDbType.Int).Direction = ParameterDirection.Output;
				command.Parameters.Add("@ctRootId", SqlDbType.Int).Value = obj.RootId;
				command.Parameters.Add("@ctNodeKind", SqlDbType.NVarChar).Value = obj.Kind;
				command.Parameters.Add("@catName", SqlDbType.NVarChar).Value = obj.Name;
				command.Parameters.Add("@ctNameLogo", SqlDbType.NVarChar).Value = DBNull.Value;
				command.Parameters.Add("@catShowOrder", SqlDbType.Int).Value = obj.Order;
				command.Parameters.Add("@dataLevel", SqlDbType.Int).Value = obj.Level;
				command.Parameters.Add("@dataParent", SqlDbType.Int).Value = SqlDbHelper.getDBNull(obj.ParentId);
				command.Parameters.Add("@childCount", SqlDbType.Int).Value = obj.ChildCount;
				command.Parameters.Add("@dataCount", SqlDbType.Int).Value = 0;
				command.Parameters.Add("@ctUnitId", SqlDbType.Int).Value = SqlDbHelper.getDBNull(obj.UnitId);
				command.Parameters.Add("@inUse", SqlDbType.NVarChar).Value = obj.InUse ? "Y" : "N";
				command.Parameters.Add("@editDate", SqlDbType.DateTime).Value = obj.ModifyDate;
				command.Parameters.Add("@editUserId", SqlDbType.NVarChar).Value = obj.ModifyUser;
				//command.Parameters.Add("@dcondition", SqlDbType.NVarChar).Value = "";
				command.Parameters.Add("@xslList", SqlDbType.NVarChar).Value = SqlDbHelper.getDBNull(obj.XslList);
				command.Parameters.Add("@xslData", SqlDbType.NVarChar).Value = SqlDbHelper.getDBNull(obj.XslData);
				command.Parameters.Add("@ctNodeNpkind", SqlDbType.NVarChar).Value = DBNull.Value;
				command.Parameters.Add("@rss", SqlDbType.NVarChar).Value = DBNull.Value;
				command.Parameters.Add("@catNameMemo", SqlDbType.NVarChar).Value = obj.CatNameMemo;

				int result = command.ExecuteNonQuery();

				obj.Id = Convert.ToInt32(command.Parameters["@Id"].Value);
			}
		}
	}

	public void update(Object vo)
	{
		CatelogTreeNode obj = (CatelogTreeNode)vo;

		using (SqlConnection conn = (SqlConnection)SqlDbHelper.getInstance().getConnection())
		{
			conn.Open();
			using (SqlCommand command = new SqlCommand())
			{
				string sql = "UPDATE CatTreeNode SET ctRootId = @ctRootId, ctNodeKind = @ctNodeKind, catName = @catName, ctNameLogo = @ctNameLogo, catShowOrder = @catShowOrder, dataLevel = @dataLevel, dataParent = @dataParent, childCount = @childCount, dataCount = @dataCount, ctUnitId = @ctUnitId, inUse = @inUse, editDate = @editDate, editUserId = @editUserId, dcondition = @dcondition, xslList = @xslList, xslData = @xslData, ctNodeNpkind = @ctNodeNpkind, rss = @rss,catNameMemo=@catNameMemo WHERE ctNodeId = @ctNodeId";
				
				command.CommandText = sql;
				command.Connection = conn;
				command.Parameters.AddWithValue("@ctRootId", obj.RootId);
				command.Parameters.AddWithValue("@ctNodeKind", obj.Kind);
				command.Parameters.AddWithValue("@ctNodeId", obj.Id);
				command.Parameters.AddWithValue("@catName", obj.Name);
				command.Parameters.AddWithValue("@ctNameLogo", DBNull.Value);
				command.Parameters.AddWithValue("@catShowOrder", obj.Order);
				command.Parameters.AddWithValue("@dataLevel", obj.Level);
				command.Parameters.AddWithValue("@dataParent", SqlDbHelper.getDBNull1(obj.ParentId));
				command.Parameters.AddWithValue("@childCount", obj.ChildCount);
				command.Parameters.AddWithValue("@dataCount", 0);
				command.Parameters.AddWithValue("@ctUnitId", SqlDbHelper.getDBNull(obj.UnitId));
				command.Parameters.Add("@inUse", SqlDbType.VarChar).Value = obj.InUse ? "Y" : "N";
				command.Parameters.AddWithValue("@editDate", obj.ModifyDate);
				command.Parameters.AddWithValue("@editUserId", obj.ModifyUser);
				command.Parameters.AddWithValue("@dcondition", SqlDbHelper.getDBNull(obj.Condition));
				command.Parameters.AddWithValue("@xslList", SqlDbHelper.getDBNull(obj.XslList));
				command.Parameters.AddWithValue("@xslData", SqlDbHelper.getDBNull(obj.XslData));
				command.Parameters.AddWithValue("@ctNodeNpkind", DBNull.Value);
				command.Parameters.AddWithValue("@rss", DBNull.Value);
				command.Parameters.Add("@catNameMemo", SqlDbType.NVarChar).Value = SqlDbHelper.getDBNull(obj.CatNameMemo);

				int result = command.ExecuteNonQuery();
			}
		}
	}

	public void delete(Object vo)
	{
		CatelogTreeNode obj = (CatelogTreeNode)vo;

		using (SqlConnection conn = (SqlConnection)SqlDbHelper.getInstance().getConnection())
		{
			conn.Open();
			using (SqlCommand command = new SqlCommand())
			{
				string sql = "DELETE FROM CatTreeNode WHERE ctNodeId = @ctNodeId";

				command.CommandText = sql;
				command.Connection = conn;
				command.Parameters.AddWithValue("@ctNodeId", obj.Id);

				int result = command.ExecuteNonQuery();
			}
		}
	}

	public IList findByFilter(IDictionary filter)
	{
		IList results = new ArrayList();

		using (SqlConnection conn = (SqlConnection)SqlDbHelper.getInstance().getConnection())
		{
			conn.Open();
			using (SqlCommand command = new SqlCommand())
			{
				string sql = "SELECT * FROM CatTreeNode WHERE 1 = 1";
				foreach (DictionaryEntry entry in filter)
				{
					if(entry.Value != null)
					{
						sql += " AND " + entry.Key + " = @" + entry.Key;
						command.Parameters.AddWithValue("@" + entry.Key, entry.Value);
					}
					else
					{
						sql += " AND (" + entry.Key + " = 0 OR " + entry.Key + " IS NULL) ";
					}
				}

				sql += " ORDER BY catShowOrder ASC";
				
				command.CommandText = sql;
				command.Connection = conn;

				IDataReader reader = command.ExecuteReader();
				while (reader.Read())
				{
					results.Add(populateFromReader(reader));
				}
			}
		}

		return results;
	}

	public CatelogTreeNode populateFromReader(IDataReader reader)
	{
		CatelogTreeNode node = new CatelogTreeNode();
		node.Id = Convert.ToInt16(reader["ctNodeId"]);
		node.RootId = Convert.ToInt16(reader["ctRootId"]);
		node.Kind = (string)reader["ctNodeKind"];
		node.Name = (string)reader["catName"];
		node.Order = Convert.ToInt16(reader["catShowOrder"]);
		node.Level = Convert.ToInt16(reader["dataLevel"]);

        if (Convert.IsDBNull(reader["dataParent"]))
            node.ParentId = null;
        else if (Convert.ToInt32(reader["dataParent"]) == 0)
            node.ParentId = null;
        else
		node.ParentId = SqlDbHelper.getNullableInt32(reader["dataParent"]);

		node.ChildCount = Convert.ToInt16(reader["childCount"]);
		node.UnitId = SqlDbHelper.getNullableInt32(reader["ctUnitId"]);
		node.InUse = (string)reader["inUse"] == "Y" ? true : false;
		node.CreateUser = (string)reader["editUserId"];
		node.CreateDate = (DateTime)reader["editDate"];
		node.ModifyUser = (string)reader["editUserId"];
		node.ModifyDate = (DateTime)reader["editDate"];
		node.Condition = SqlDbHelper.getNullableString(reader["dcondition"]);
		node.XslList = SqlDbHelper.getNullableString(reader["xslList"]);
		node.XslData = SqlDbHelper.getNullableString(reader["xslData"]);
		node.CatNameMemo = SqlDbHelper.getNullableString(reader["CatNameMemo"]);
		return node;
	}

}
