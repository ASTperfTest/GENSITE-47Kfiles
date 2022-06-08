using System;
using System.Data;
using System.Collections.Generic;

/// <summary>
/// ContentUnit 的摘要描述
/// </summary>
public class ContentUnit
{
	public const string MODULE_CONTENT = "1";
	public const string MODULE_LIST = "2";
	public const string MODULE_OUTSIDE = "4";
	public const string MODULE_URL = "U";

	private int _id;

	public int Id
	{
		get { return _id; }
		set { _id = value; }
	}

	private string _name;

	public string Name
	{
		get { return _name; }
		set { _name = value; }
	}

	private string _kind;

	public string Kind
	{
		get { return _kind; }
		set { _kind = value; }
	}

	private string _redirectUrl;

	public string RedirectUrl
	{
		get { return _redirectUrl; }
		set { _redirectUrl = value; }
	}

	private bool _newWindow = false;

	public bool NewWindow
	{
		get { return _newWindow; }
		set { _newWindow = value; }
	}

	private int _baseDsd;

	public int BaseDsd
	{
		get { return _baseDsd; }
		set { _baseDsd = value; }
	}

	private bool _unitOnly = true;

	public bool UnitOnly
	{
		get { return _unitOnly; }
		set { _unitOnly = value; }
	}

	private bool _inUse;

	public bool InUse
	{
		get { return _inUse; }
		set { _inUse = value; }
	}

	private bool _checkYN = false;

	public bool CheckYN
	{
		get { return _checkYN; }
		set { _checkYN = value; }
	}

	private string _deptId;

	public string DeptId
	{
		get { return _deptId; }
		set { _deptId = value; }
	}

	private IList<string> _listFields = new List<string>();

	public IList<string> ListFields
	{
		get { return _listFields; }
	}

	private IList<string> _contentFields = new List<string>();

	public IList<string> ContentFields
	{
		get { return _contentFields; }
	}

	private IList<string> _dataCategories = new List<string>();

	public IList<string> DataCategories
	{
		get { return _dataCategories; }
		set { _dataCategories = value; }
	}
	
	private string _catNameMemo;

	public string CatNameMemo
	{
		get { return _catNameMemo; }
		set { _catNameMemo = value; }
	}


	public ContentUnit()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}
}
