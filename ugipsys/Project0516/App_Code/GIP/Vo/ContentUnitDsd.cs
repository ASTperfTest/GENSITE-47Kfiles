using System;
using System.Data;
using System.Collections.Generic;

/// <summary>
/// ContentUnitDSD 的摘要描述
/// </summary>
public class ContentUnitDsd
{
	private int _unitId;						// 單位編號

	public int UnitId
	{
		get { return _unitId; }
		set { _unitId = value; }
	}

	private int _baseDad;						// 基礎資料結構編號

	public int BaseDad
	{
		get { return _baseDad; }
		set { _baseDad = value; }
	}

	private string _showClientSqlOrderBy;		// 前台條列排序欄位

	public string ShowClientSqlOrderBy
	{
		get { return _showClientSqlOrderBy; }
		set { _showClientSqlOrderBy = value; }
	}

	private string _showClientSqlOrderByType;	// 升/降羃

	public string ShowClientSqlOrderByType
	{
		get { return _showClientSqlOrderByType; }
		set { _showClientSqlOrderByType = value; }
	}

	private string _formClientCat;				// 分類欄位

	public string FormClientCat
	{
		get { return _formClientCat; }
		set { _formClientCat = value; }
	}

	private string _formClientCatRefLookup;		// 參照代碼

	public string FormClientCatRefLookup
	{
		get { return _formClientCatRefLookup; }
		set { _formClientCatRefLookup = value; }
	}

	private string _showClientStyle;			// 條列呈現

	public string ShowClientStyle
	{
		get { return _showClientStyle; }
		set { _showClientStyle = value; }
	}

	private string _formClientStyle;			// 內容呈現

	public string FormClientStyle
	{
		get { return _formClientStyle; }
		set { _formClientStyle = value; }
	}

    private string _clientDefault;              // 預設

    public string ClientDefault
    {
        get { return _clientDefault; }
        set { _clientDefault = value; }    
    }

    private string _showTypeStr;

    public string showTypeStr
    {
        get { return _showTypeStr; }
        set { _showTypeStr = value; }
    }

    

	private IDictionary<string, ContentUnitDsdTable> _tables = new Dictionary<string, ContentUnitDsdTable>();

	public IDictionary<string, ContentUnitDsdTable> Tables
	{
		get { return _tables; }
	}

	public ContentUnitDsd()
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

		foreach (KeyValuePair<string, ContentUnitDsdTable> entry in Tables)
		{
			builder.Append(entry.Value);
		}

		builder.Append("====================================================\n");

		return builder.ToString();
	}
}
