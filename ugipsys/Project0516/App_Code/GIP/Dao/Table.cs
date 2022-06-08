using System;
using System.Data;
using System.Collections.Generic;

/// <summary>
/// Table 的摘要描述
/// </summary>
public class Table
{
	private string _name;

	public string Name
	{
		get { return _name; }
		set { _name = value; }
	}

	private List<Column> _keys;

	public List<Column> Keys
	{
		get { return _keys; }
		set { _keys = value; }
	}

	private List<Column> _columns;

	public List<Column> Columns
	{
		get { return _columns; }
		set { _columns = value; }
	}

	public Table()
	{
		_columns = new List<Column>();
	}

	public string[] getKeyNames()
	{
		string[] keyNames = new string[this.Keys.Count];

		for (int i = 0; i < this.Keys.Count; i++)
		{
			keyNames[i] = this.Keys[i].Name;
		}

		return keyNames;
	}

	public string[] getColumnNames()
	{
		string[] columnNames = new string[this.Columns.Count];
		
		for (int i = 0; i < this.Columns.Count; i++)
		{
			columnNames[i] = this.Columns[i].Name;
		}

		return columnNames;
	}
}
