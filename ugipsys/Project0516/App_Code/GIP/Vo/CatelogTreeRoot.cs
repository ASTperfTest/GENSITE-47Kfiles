using System;
using System.Collections.Generic;

/// <summary>
/// CatelogTreeRoot 的摘要描述
/// </summary>
public class CatelogTreeRoot
{
	private int _id;

	public int Id
	{
		get { return _id; }
		set { _id = value; }
	}

	private String _name;

	public String Name
	{
		get { return _name; }
		set { _name = value; }
	}

	private Boolean _inUse;

	public Boolean InUse
	{
		get { return _inUse; }
		set { _inUse = value; }
	}

	private String _modifyUser;

	public String ModifyUser
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

	private List<CatelogTreeNode> _nodes = null;

	public List<CatelogTreeNode> Nodes
	{
		get { return _nodes; }
	}

	public CatelogTreeRoot()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}
}
