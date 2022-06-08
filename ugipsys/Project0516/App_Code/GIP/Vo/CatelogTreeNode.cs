using System;
using System.Collections.Generic;

/// <summary>
/// CatelogTreeNode 的摘要描述
/// </summary>
public class CatelogTreeNode
{
	public static string CATELOG = "C";
	public static string UNIT = "U";

	private int _id;

	public int Id
	{
		get { return _id; }
		set { _id = value; }
	}

	private int _rootId;

	public int RootId
	{
		get { return _rootId; }
		set { _rootId = value; }
	}

	private string _kind;

	public string Kind
	{
		get { return _kind; }
		set { _kind = value; }
	}

	private string _name;

	public string Name
	{
		get { return _name; }
		set { _name = value; }
	}

	private int _order;

	public int Order
	{
		get { return _order; }
		set { _order = value; }
	}

	private int _level;

	public int Level
	{
		get { return _level; }
		set { _level = value; }
	}

	private Nullable<int> _parentId;

	public Nullable<int> ParentId
	{
		get { return _parentId; }
		set { _parentId = value; }
	}

	private int _childCount;

	public int ChildCount
	{
		get { return _childCount; }
		set { _childCount = value; }
	}

	private Nullable<int> _unitId;

	public Nullable<int> UnitId
	{
		get { return _unitId; }
		set { _unitId = value; }
	}

	private bool _inUse;

	public bool InUse
	{
		get { return _inUse; }
		set { _inUse = value; }
	}

	private string _xslList;

	public string XslList
	{
		get { return _xslList; }
		set { _xslList = value; }
	}

	private string _xslData;

	public string XslData
	{
		get { return _xslData; }
		set { _xslData = value; }
	}

	private string _condition;

	public string Condition
	{
		get { return _condition; }
		set { _condition = value; }
	}

	private string _createUser;

	public string CreateUser
	{
		get { return _createUser; }
		set { _createUser = value; }
	}

	private DateTime _createDate;

	public DateTime CreateDate
	{
		get { return _createDate; }
		set { _createDate = value; }
	}

	private string _modifyUser;

	public string ModifyUser
	{
		get { return _modifyUser; }
		set { _modifyUser = value; }
	}

	private DateTime _modifyDate;

	public DateTime ModifyDate
	{
		get { return _modifyDate; }
		set { _modifyDate = value; }
	}

	private CatelogTreeRoot _root;

	public CatelogTreeRoot Root
	{
		get { return _root; }
		set { _root = value; }
	}

	public CatelogTreeNode _parent;

	public CatelogTreeNode Parent
	{
		get { return _parent; }
		set
		{
			_parent = value;
			if (_parent != null)
				_parentId = _parent.Id;
		}
	}

	private IList<CatelogTreeNode> _childs;

	public IList<CatelogTreeNode> Childs
	{
		get { return _childs; }
		set { _childs = value; }
	}
	
	private string _catNameMemo;

	public string CatNameMemo
	{
		get { return _catNameMemo; }
		set { _catNameMemo = value; }
	}

	public CatelogTreeNode()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}
}
