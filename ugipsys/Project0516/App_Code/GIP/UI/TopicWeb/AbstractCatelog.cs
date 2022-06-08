using System;

/// <summary>
/// AbstractCatelog 的摘要描述
/// </summary>
namespace Hyweb.M00.COA.GIP.TopicWeb
{
	[Serializable]
	public class AbstractCatelog
	{
		public const string FOLDER = "C";
		public const string UNIT = "U";

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

		private Nullable<int> _parentId;

		public Nullable<int> ParentId
		{
			get { return _parentId; }
			set { _parentId = value; }
		}

		private int _order;

		public int Order
		{
			get { return _order; }
			set { _order = value; }
		}

		private bool _inUse;

		public bool InUse
		{
			get { return _inUse; }
			set { _inUse = value; }
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

		public AbstractCatelog()
		{
			//
			// TODO: 在此加入建構函式的程式碼
			//
		}
	}
}