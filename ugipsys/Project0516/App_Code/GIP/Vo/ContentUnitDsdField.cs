using System;

/// <summary>
/// ContentUnitDsdItem 的摘要描述
/// </summary>
public class ContentUnitDsdField
{
	private string _name;				// 欄位代碼

	public string Name
	{
		get { return _name; }
		set { _name = value; }
	}

	private string _desc;				// 欄位說明

	public string Desc
	{
		get { return _desc; }
		set { _desc = value; }
	}

	private int _seq;					// 顯示順序

	public int Seq
	{
		get { return _seq; }
		set { _seq = value; }
	}

	private string _oriLabel;			// 欄位原有名稱

	public string OriLabel
	{
		get { return _oriLabel; }
		set { _oriLabel = value; }
	}

	private string _label;				// 欄位顯示名稱

	public string Label
	{
		get { return _label; }
		set { _label = value; }
	}

	private string _dataType;			// 欄位型態

	public string DataType
	{
		get { return _dataType; }
		set { _dataType = value; }
	}

	private int _dataLength;			// 欄位長度

	public int DataLength
	{
		get { return _dataLength; }
		set { _dataLength = value; }
	}

	private int _inputLength;		// 輸入長度

	public int InputLength
	{
		get { return _inputLength; }
		set { _inputLength = value; }
	}

	private bool _canNull;				// 是否允許NULL

	public bool CanNull
	{
		get { return _canNull; }
		set { _canNull = value; }
	}

	private bool _isPrimaryKey;			// 是否為主鍵

	public bool IsPrimaryKey
	{
		get { return _isPrimaryKey; }
		set { _isPrimaryKey = value; }
	}

	private bool _isIdentity;			// 是否為識別

	public bool IsIdentity
	{
		get { return _isIdentity; }
		set { _isIdentity = value; }
	}

	private string _inputType;			// 輸入形式

	public string InputType
	{
		get { return _inputType; }
		set { _inputType = value; }
	}

	private int _rows;					// 行數

	public int Rows
	{
		get { return _rows; }
		set { _rows = value; }
	}

	private int _cols;					// 列數

	public int Cols
	{
		get { return _cols; }
		set { _cols = value; }
	}
	
	private bool _isShowListClient;		// 前台條列是否顯示?

	public bool IsShowListClient
	{
		get { return _isShowListClient; }
		set { _isShowListClient = value; }
	}

	private bool _isFormListClient;		// 前台內容是否顯示?

	public bool IsFormListClient
	{
		get { return _isFormListClient; }
		set { _isFormListClient = value; }
	}

	private bool _isQueryListClient;	// 前台查詢是否顯示?

	public bool IsQueryListClient
	{
		get { return _isQueryListClient; }
		set { _isQueryListClient = value; }
	}

	private bool _isShowList;			// 後台條列是否顯示?

	public bool IsShowList
	{
		get { return _isShowList; }
		set { _isShowList = value; }
	}

	private bool _isFormList;			// 後台新增/編修是否顯示?

	public bool IsFormList
	{
		get { return _isFormList; }
		set { _isFormList = value; }
	}

	private bool _isQueryList;			// 後台查詢是否顯示?

	public bool IsQueryList
	{
		get { return _isQueryList; }
		set { _isQueryList = value; }
	}

	private string _paramKind;			// 參數型態

	public string ParamKind
	{
		get { return _paramKind; }
		set { _paramKind = value; }
	}

	private string _refLookup;			// 參考

	public string RefLookup
	{
		get { return _refLookup; }
		set { _refLookup = value; }
	}

    private string _showTypeStr;

    public string showTypeStr
    {
        get { return _showTypeStr; }
        set { _showTypeStr = value; }
    }
    //private System.Collections.DictionaryEntry _showTypeStr = new System.Collections.DictionaryEntry();

    //public System.Collections.DictionaryEntry showTypeStr
    //{
    //    get { return _showTypeStr; }
    //    set { _showTypeStr = value; }
    //}


	private System.Collections.DictionaryEntry _clientDefault = new System.Collections.DictionaryEntry();		// 頁面預設值

	public System.Collections.DictionaryEntry ClientDefault
	{
		get { return _clientDefault; }
		set { _clientDefault = value; }
	}

	public ContentUnitDsdField()
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

		builder.Append("====================================================\n");

		return builder.ToString();
	}
}
