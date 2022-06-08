using System;
using System.Collections;

/// <summary>
/// CatelogNode 的摘要描述
/// </summary>
namespace Hyweb.M00.COA.GIP.TopicWeb
{
	[Serializable]
	public class CatelogNode : AbstractCatelog
	{
		private int _unitId;				// 單元編號

		public int UnitId
		{
			get { return _unitId; }
			set { _unitId = value; }
		}

		private string _unitKind;			// 功能類型

		public string UnitKind
		{
			get { return _unitKind; }
			set { _unitKind = value; }
		}

		private string _listStyle;			// 列表版型

		public string ListStyle
		{
			get { return _listStyle; }
			set { _listStyle = value; }
		}

		private IList _listFields = new ArrayList();			// 列表顯示欄位

		public IList ListFields
		{
			get { return _listFields; }
			set { _listFields = value; }
		}

		private string _contentStyle;		// 內容版型

		public string ContentStyle
		{
			get { return _contentStyle; }
			set { _contentStyle = value; }
		}

		private IList _contentFields = new ArrayList();		// 內容顯示欄位

		public IList ContentFields
		{
			get { return _contentFields; }
			set { _contentFields = value; }
		}

		private string _dataCategoryType;	// 資料大類形態

		public string DataCategoryType
		{
			get { return _dataCategoryType; }
			set { _dataCategoryType = value; }
		}

		private IList _dataCategories;		// 資料大類

		public IList DataCategories
		{
			get { return _dataCategories; }
			set { _dataCategories = value; }
		}
		
		private string _catNameMemo;			// memo

		public string CatNameMemo
		{
			get { return _catNameMemo; }
			set { _catNameMemo = value; }
		}

		public CatelogNode()
		{
			//
			// TODO: 在此加入建構函式的程式碼
			//
		}
	}
}