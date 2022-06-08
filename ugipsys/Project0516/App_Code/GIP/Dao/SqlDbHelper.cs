using System;
using System.Data;
using System.Collections;

/// <summary>
/// SqlDbHelper 的摘要描述
/// </summary>
public class SqlDbHelper
{
	private static SqlDbHelper _instance;

	public SqlDbHelper()
	{
	}

	public static SqlDbHelper getInstance()
	{
		if (_instance == null)
			_instance = new SqlDbHelper();
		return _instance;
	}

	public IDbConnection getConnection()
	{
		return new System.Data.SqlClient.SqlConnection(SqlDbManager.CONNECTION_STRING);
	}

	public string getSelectCommandText(string tableName, string[] conditionNames)
	{
		string commandText = "SELECT * FROM {0} WHERE {1}";

		System.Text.StringBuilder conditionNameBuilder = new System.Text.StringBuilder();

		for (int i = 0; i < conditionNames.Length; i++)
		{
			if (i != 0)
			{
				conditionNameBuilder.Append(" AND ");
			}
			conditionNameBuilder.Append(String.Format("{0} = @{1}", conditionNames[i], conditionNames[i]));
		}

		return String.Format(commandText, tableName, conditionNameBuilder.ToString());
	}

	public string getInsertCommandText(Table table)
	{
		return getInsertCommandText(table.Name, table.getColumnNames());
	}

	public string getInsertCommandText(string tableName, string[] columnNames)
	{
		string commandText = "INSERT INTO {0} ({1}) VALUES ({2})";

		System.Text.StringBuilder columnNameBuilder = new System.Text.StringBuilder();
		System.Text.StringBuilder columnValueBuilder = new System.Text.StringBuilder();

		for (int i = 0; i < columnNames.Length; i++)
		{
			if (i != 0)
			{
				columnNameBuilder.Append(", ");
				columnValueBuilder.Append(", ");
			}
			columnNameBuilder.Append(columnNames[i]);
			columnValueBuilder.Append("@" + columnNames[i]);
		}

		return String.Format(commandText, tableName, columnNameBuilder.ToString(), columnValueBuilder.ToString());
	}

	public string getUpdateCommandText(Table table)
	{
		return getUpdateCommandText(table.Name, table.getColumnNames(), table.getKeyNames());
	}

	public string getUpdateCommandText(string tableName, string[] columnNames, string[] keyNames)
	{
		string commandText = "UPDATE {0} SET {1} WHERE {2}";

		System.Text.StringBuilder columnsBuilder = new System.Text.StringBuilder();
		System.Text.StringBuilder keysBuilder = new System.Text.StringBuilder();

		for (int i = 0; i < columnNames.Length; i++)
		{
			if (i != 0)
			{
				columnsBuilder.Append(", ");
			}
			columnsBuilder.Append(String.Format("{0} = @{1}", columnNames[i], columnNames[i]));
		}

		for (int i = 0; i < keyNames.Length; i++)
		{
			if (i != 0)
			{
				keysBuilder.Append(", ");
			}
			keysBuilder.Append(String.Format("{0} = @{1}", keyNames[i], keyNames[i]));
		}

		return String.Format(commandText, tableName, columnsBuilder.ToString(), keysBuilder.ToString());
	}

	public string getDeleteCommandText(Table table)
	{
		return getDeleteCommandText(table.Name, table.getKeyNames());
	}

	public string getDeleteCommandText(string tableName, string[] keyNames)
	{
		string commandText = "DELETE FROM {0} WHERE {1}";

		System.Text.StringBuilder keysBuilder = new System.Text.StringBuilder();

		for (int i = 0; i < keyNames.Length; i++)
		{
			if (i != 0)
			{
				keysBuilder.Append(" AND ");
			}
			keysBuilder.Append(String.Format("{0} = @{1}", keyNames[i], keyNames[i]));
		}

		return String.Format(commandText, tableName, keysBuilder.ToString());
	}

	public static Nullable<int> getNullableInt32(Object obj)
	{
		if (obj == DBNull.Value)
			return null;
		else
			return Convert.ToInt32(obj);
	}

	public static string getNullableString(Object obj)
	{
		if (obj == DBNull.Value)
			return null;
		else
			return Convert.ToString(obj);
	}

	public static Object getDBNull(Object obj)
	{
		if (obj == null)
			return DBNull.Value;
		else
			return obj;
	}
public static Object getDBNull1(Object obj)
	{
		if (obj == null)
			return 0;
		else
			return obj;
	}
}
