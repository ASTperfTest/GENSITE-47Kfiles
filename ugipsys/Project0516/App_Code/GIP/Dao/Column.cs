using System;
using System.Data;

/// <summary>
/// Column 的摘要描述
/// </summary>
public class Column
{
	private string _name;

	public string Name
	{
		get { return _name; }
		set { _name = value; }
	}

	private SqlDbType _type;

	public SqlDbType Type
	{
		get { return _type; }
		set { _type = value; }
	}

	private int _size;

	public int Size
	{
		get { return _size; }
		set { _size = value; }
	}

	private Object _value;

	public Object Value
	{
		get { return _value; }
		set { _value = value; }
	}

	public Column()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}

	public Column(string name, SqlDbType type)
	{
		this.Name = name;
		this.Type = type;
	}

	public Column(string name, SqlDbType type, int size) : this(name, type)
	{
		this.Size = size;
	}
}
