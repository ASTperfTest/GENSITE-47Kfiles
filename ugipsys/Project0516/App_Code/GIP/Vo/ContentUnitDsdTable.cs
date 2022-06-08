using System;
using System.Collections.Generic;

/// <summary>
/// ContentUnitDsdTable 的摘要描述
/// </summary>
public class ContentUnitDsdTable
{
	private string _name;	// 資料表名稱

	public string Name
	{
		get { return _name; }
		set { _name = value; }
	}

	private string _desc;	// 說明

	public string Desc
	{
		get { return _desc; }
		set { _desc = value; }
	}

	private IDictionary<string, ContentUnitDsdField> _fields = new Dictionary<string, ContentUnitDsdField>();	// 欄位資料

	public IDictionary<string, ContentUnitDsdField> Fields
	{
		get { return _fields; }
	}

	public ContentUnitDsdTable()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}

	public override string ToString()
	{
		System.Text.StringBuilder builder = new System.Text.StringBuilder();

		builder.Append("====================================================\n");
		builder.Append("Object: " + base.ToString() + "\n");
		builder.Append("----------------------------------------------------\n");

		System.Reflection.PropertyInfo[] properties = this.GetType().GetProperties();

		foreach (System.Reflection.PropertyInfo property in properties)
		{
			builder.Append(property.Name + " = " + property.GetValue(this, null) + "\n");
		}

		foreach (KeyValuePair<string, ContentUnitDsdField> entry in Fields)
		{
			builder.Append(entry.Value);
		}

		builder.Append("====================================================\n");

		return builder.ToString();
	}
}
