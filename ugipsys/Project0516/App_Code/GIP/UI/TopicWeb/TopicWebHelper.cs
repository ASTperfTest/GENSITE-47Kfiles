using System;
using System.Data;
using System.Collections;
using System.Collections.Generic;

/// <summary>
/// TopicWebHelper 的摘要描述
/// </summary>
namespace Hyweb.M00.COA.GIP.TopicWeb
{
	public class TopicWebHelper
	{
		private static TopicWebHelper _instance;

		public TopicWebHelper()
		{
			//
			// TODO: 在此加入建構函式的程式碼
			//
		}

		public static TopicWebHelper getInstance()
		{
			if (_instance == null)
				_instance = new TopicWebHelper();
			return _instance;
		}

		// 取得根目錄
		public CatelogTreeRoot getRoot(int rootId)
		{
			return (CatelogTreeRoot)CatelogTreeRootDao.getInstance().findById(rootId);
		}

		// 取得根目錄的子目錄
		public IList getChildNodes(CatelogTreeRoot root)
		{
			IDictionary filter = new Hashtable();
			filter.Add("ctRootId", root.Id);
			filter.Add("dataParent", null);

			return CatelogTreeNodeDao.getInstance().findByFilter(filter);
		}

		// 新增目錄
		public CatelogTreeNode addCatelogFolder(string name, bool isOpen, int rootId, int parentId, string userId,string memo)
		{
			CatelogTreeRoot root = getRoot(rootId);
			CatelogTreeNode parent = (CatelogTreeNode)CatelogTreeNodeDao.getInstance().findById(parentId);

			CatelogTreeNode node = new CatelogTreeNode();
			node.Name = name;
			node.Kind = CatelogTreeNode.CATELOG;
			node.InUse = isOpen;
			node.Level = 1;
			node.Order = 99;
			node.RootId = root.Id;
			node.CatNameMemo = memo;

			if (parent != null)
			{
				node.ParentId = parent.Id;
			}
			else
			{
				node.ParentId = 0;
			}
			
			node.CreateUser = userId;
			node.CreateDate = DateTime.Now;
			node.ModifyUser = userId;
			node.ModifyDate = DateTime.Now;

			CatelogTreeNodeDao.getInstance().insert(node);

			return node;
		}

		// 取得目錄
		public CatelogTreeNode getCatelogFolder(int catelogId)
		{
			CatelogTreeNode node = (CatelogTreeNode)CatelogTreeNodeDao.getInstance().findById(catelogId);
			node.Root = (CatelogTreeRoot)CatelogTreeRootDao.getInstance().findById(node.RootId);
			node.Parent = (CatelogTreeNode)CatelogTreeNodeDao.getInstance().findById(node.ParentId);

			return node;
		}

		// 更新目錄
		public void updateCatelogFolder(CatelogTreeNode node,string username)
		{
			node.ModifyUser = username;
			node.ModifyDate = DateTime.Now;
			CatelogTreeNodeDao.getInstance().update(node);
		}

		// 刪除目錄
		public void deleteCatelogFolder(int nodeId)
		{
			CatelogTreeNode node = (CatelogTreeNode)CatelogTreeNodeDao.getInstance().findById(nodeId);

			// 檢查是否有子結點
			IList nodes = this.getNodesByParent(node);
			if (nodes.Count > 0)
				throw new Exception("此目錄含有子節點，無法刪除。");

			CatelogTreeNodeDao.getInstance().delete(node);
		}

		// 取得目錄下的節點
		public IList getNodesByParent(CatelogTreeNode node)
		{
			IDictionary filter = new Hashtable();
			filter.Add("dataParent", node.Id);

			return CatelogTreeNodeDao.getInstance().findByFilter(filter);
		}

		// 往上移動節點
		public void moveUpNode(int nodeId)
		{
			CatelogTreeNode node = (CatelogTreeNode)CatelogTreeNodeDao.getInstance().findById(nodeId);

			IList childs;

			if (node.ParentId.HasValue)
			{
				CatelogTreeNode parent = (CatelogTreeNode)CatelogTreeNodeDao.getInstance().findById(node.ParentId);

				IDictionary filter = new Hashtable();
				filter.Add("dataParent", parent.Id);

				childs = CatelogTreeNodeDao.getInstance().findByFilter(filter);
			}
			else
			{
				CatelogTreeRoot root = (CatelogTreeRoot)CatelogTreeRootDao.getInstance().findById(node.RootId);

				IDictionary filter = new Hashtable();
				filter.Add("ctRootId", root.Id);
				filter.Add("dataParent", 0);

				childs = CatelogTreeNodeDao.getInstance().findByFilter(filter);
			}

			if (childs != null)
			{
				for (int i = childs.Count - 1; i > 0; i--)
				{
					CatelogTreeNode nodeBefore = (CatelogTreeNode)childs[i-1];
					CatelogTreeNode nodeCurrent = (CatelogTreeNode)childs[i];

					if (nodeCurrent.Id == node.Id)
					{
						nodeBefore.Order = i;
						nodeCurrent.Order = i - 1;

						CatelogTreeNodeDao.getInstance().update(nodeBefore);
						CatelogTreeNodeDao.getInstance().update(nodeCurrent);
					}
				}
			}

		}

		// 往下移動節點
		public void moveDownNode(int nodeId)
		{
			CatelogTreeNode node = (CatelogTreeNode)CatelogTreeNodeDao.getInstance().findById(nodeId);

//if(node != null)
//	throw new Exception("nodeId = " + nodeId + " and node.ParentId = " + node.ParentId + ". ");

			IList childs;

			if (node.ParentId.HasValue)
			{
				CatelogTreeNode parent = (CatelogTreeNode)CatelogTreeNodeDao.getInstance().findById(node.ParentId);

				IDictionary filter = new Hashtable();
				filter.Add("dataParent", parent.Id);

				childs = CatelogTreeNodeDao.getInstance().findByFilter(filter);
			}
			else
			{
				CatelogTreeRoot root = (CatelogTreeRoot)CatelogTreeRootDao.getInstance().findById(node.RootId);

				IDictionary filter = new Hashtable();
				filter.Add("ctRootId", root.Id);
				filter.Add("dataParent", 0);

				childs = CatelogTreeNodeDao.getInstance().findByFilter(filter);
			}

			if (childs != null)
			{
				for (int i = 0; i < childs.Count - 1; i++)
				{
					CatelogTreeNode nodeCurrent = (CatelogTreeNode)childs[i];
					CatelogTreeNode nodeAfter = (CatelogTreeNode)childs[i+1];

					if (nodeCurrent.Id == node.Id)
					{
						nodeCurrent.Order = i + 1;
						nodeAfter.Order = i;

						CatelogTreeNodeDao.getInstance().update(nodeCurrent);
						CatelogTreeNodeDao.getInstance().update(nodeAfter);
					}
				}
			}
		}

		// 取得可選取的欄位列表
		public IList getAvailiableFields()
		{
			IList avaliableFields = new ArrayList();

			avaliableFields.Add(new DictionaryEntry("ximgFile", "圖片"));
			avaliableFields.Add(new DictionaryEntry("xpostDate", "張貼日期"));

			return avaliableFields;
		}

		// 取得資料大類設定
		public IList getDataCategoryTypes()
		{
			IList list = new ArrayList();

			list.Add(new DictionaryEntry("CUSTOM", "自訂分類"));

			// 取得代碼表中開頭為CustomWebDefCat的代碼
			CodeMainDataSetTableAdapters.CodeMetaDefTableAdapter codeMetaAdapter = new CodeMainDataSetTableAdapters.CodeMetaDefTableAdapter();
			CodeMainDataSet.CodeMetaDefDataTable codeMetaTable = codeMetaAdapter.FindByIdStartWith("CustomWebDefCat");

			foreach (CodeMainDataSet.CodeMetaDefRow row in codeMetaTable.Rows)
			{
				list.Add(new DictionaryEntry(row.codeId, row.codeName));
			}

			return list;
		}

		// 取得預設資料大類
		public IList getDefaultDataCategories(string type)
		{
			IList list = new ArrayList();

			CodeMainDataSetTableAdapters.CodeMainTableAdapter codeMainAdapter = new CodeMainDataSetTableAdapters.CodeMainTableAdapter();
			CodeMainDataSet.CodeMainDataTable codeMainTable = codeMainAdapter.FindByMeta(type);
			codeMainTable.DefaultView.Sort = "msortValue";

			foreach (CodeMainDataSet.CodeMainRow row in codeMainTable.Rows)
			{
				list.Add(new DictionaryEntry(row.mcode, row.mvalue));
			}

			return list;
		}

		// 取得個人化資料大類
		public IList getCustomDataCategories(int unitId)
		{
			IList list = new ArrayList();

			CodeMainDataSetTableAdapters.CodeMainTableAdapter codeMainAdapter = new CodeMainDataSetTableAdapters.CodeMainTableAdapter();
			CodeMainDataSet.CodeMainDataTable codeMainTable = codeMainAdapter.FindByMeta("CustomWebCat_" + unitId.ToString());
			codeMainTable.DefaultView.Sort = "msortValue";

			foreach (CodeMainDataSet.CodeMainRow row in codeMainTable.Rows)
			{
				list.Add(new DictionaryEntry(row.mcode, row.mvalue));
			}

			return list;
		}

		// 新增項目節點
		public CatelogTreeNode addCatelogNode(IDictionary arguments)
		{
			int rootId = (int)arguments["rootId"];
			int parentId = (int)arguments["parentId"];
			string userId = (string)arguments["userId"];

			string nodeName = (string)arguments["nodeName"];
			bool isOpen = (bool)arguments["isOpen"];
			string nodeType = (string)arguments["nodeType"];

			string listStyle = (string)arguments["listStyle"];
			IList listFields = (IList)arguments["listFields"];

			string contentStyle = (string)arguments["contentStyle"];
			IList contentFields = (IList)arguments["contentFields"];

			string dataCategoryType = (string)arguments["dataCategoryType"];
			IList dataCategories = (IList)arguments["dataCategories"];

            bool inUse = (bool)arguments["inUse"];
            string Genname = (string)arguments["Genname"];
            int datalevel = (int)arguments["datalevel"];
			
			string CatNameMemo = (string)arguments["catNameMemo"];

			// 新增單元範本
			ContentUnit unit = new ContentUnit();
		        if(parentId == 0)
			{
				unit.Name=nodeName;
				}
				else
			{
			unit.Name = Genname+nodeName;
			}
			unit.Kind = (nodeType == "CP" ? ContentUnit.MODULE_CONTENT : ContentUnit.MODULE_LIST);
			unit.BaseDsd = 7;
            unit.InUse = inUse;

			ContentUnitDao.getInstance().insert(unit);
			
			// 新增單元範本XML
			IList avaliableFields = getAvailiableFields();

			ContentUnitDsdDao.XmlPath = System.Web.HttpContext.Current.Server.MapPath(System.Configuration.ConfigurationManager.AppSettings["GipDsdDictionary"]);
			ContentUnitDsd dsd = ContentUnitDsdDao.getInstance().findById(0);
			dsd.UnitId = unit.Id;
			dsd.BaseDad = unit.BaseDsd;

			foreach (DictionaryEntry entry in avaliableFields)
			{
				dsd.Tables["CuDtGeneric"].Fields[(string)entry.Key].IsShowListClient = listFields.Contains((string)entry.Key);
				dsd.Tables["CuDtGeneric"].Fields[(string)entry.Key].IsFormListClient = contentFields.Contains((string)entry.Key);
			}

			if (dataCategoryType == "CUSTOM")
			{
				// 個人化資料大類代碼

                dsd.FormClientCat = "topCat";//"CustomWebCat_" + unit.Id;
				dsd.Tables["CuDtGeneric"].Fields["topCat"].RefLookup = "CustomWebCat_" + unit.Id;

				CodeMainDataSetTableAdapters.CodeMetaDefTableAdapter metaAdapter = new CodeMainDataSetTableAdapters.CodeMetaDefTableAdapter();
				CodeMainDataSetTableAdapters.CodeMainTableAdapter mainAdapter = new CodeMainDataSetTableAdapters.CodeMainTableAdapter();

				// 先刪除個人化資料大類
				metaAdapter.Delete("CustomWebCat_" + unit.Id);

				// 新增個人化資料大類META
				metaAdapter.Insert("CustomWebCat_" + unit.Id, Genname+"個人化資料大類", "CodeMain", "codeMetaId", "CustomWebCat_" + unit.Id, "mcode", "mvalue", "msortValue", "代碼值", "顯示值", "排序值", "主題網", null, "Y", null);

				// 新增個人化資料大類MAIN
				for (int i = 0; i < dataCategories.Count; i++)
				{
					mainAdapter.Insert("CustomWebCat_" + unit.Id, (string)dataCategories[i], (string)dataCategories[i], i.ToString());
				}
			}
			else
			{
				// 預設資料大類代碼

                 if (parentId == 0)
                {
                    dsd.FormClientCat = "";//dataCategoryType;
                }
                else
                {
                    dsd.FormClientCat = "topCat";//dataCategoryType;
                }
				dsd.Tables["CuDtGeneric"].Fields["topCat"].RefLookup = dataCategoryType;
			}

			ContentUnitDsdDao.getInstance().insert(dsd);

			// 新增節點資料
			CatelogTreeNode node = new CatelogTreeNode();
			node.Name = nodeName;
			node.Kind = CatelogTreeNode.UNIT;
			node.InUse = isOpen;
            node.Level = datalevel;
			node.Order = 99;
			node.RootId = rootId;
			node.UnitId = unit.Id;

			if (parentId != 0)
			{
				node.ParentId = parentId;
			}
			else
			{
				node.ParentId = 0;
			}

			node.XslList = listStyle;
			node.XslData = contentStyle;
			node.CreateUser = userId;
			node.CreateDate = DateTime.Now;
			node.ModifyUser = userId;
			node.ModifyDate = DateTime.Now;
			node.CatNameMemo = CatNameMemo;

			CatelogTreeNodeDao.getInstance().insert(node);

			// 新增上稿權限
			CtUnitDataSetTableAdapters.CtUserSetTableAdapter userSetAdapter = new CtUnitDataSetTableAdapters.CtUserSetTableAdapter();
			userSetAdapter.Insert(userId, node.Id, 1);

			return node;
		}

		// 更新項目節點
		public CatelogTreeNode updateCatelogNode(IDictionary arguments)
		{
			int nodeId = (int)arguments["nodeId"];

			string userId = (string)arguments["userId"];

			string nodeName = (string)arguments["nodeName"];
			string catNameMemo = (string)arguments["catNameMemo"];
			bool isOpen = (bool)arguments["isOpen"];
			string nodeType = (string)arguments["nodeType"];

			string listStyle = (string)arguments["listStyle"];
			IList listFields = (IList)arguments["listFields"];

			string contentStyle = (string)arguments["contentStyle"];
			IList contentFields = (IList)arguments["contentFields"];

			string dataCategoryType = (string)arguments["dataCategoryType"];
			IList dataCategories = (IList)arguments["dataCategories"];

			CatelogTreeNode node = (CatelogTreeNode)CatelogTreeNodeDao.getInstance().findById(nodeId);
			ContentUnit unit = (ContentUnit)ContentUnitDao.getInstance().get(node.UnitId.Value);
			ContentUnitDsdDao.XmlPath = System.Web.HttpContext.Current.Server.MapPath(System.Configuration.ConfigurationManager.AppSettings["GipDsdDictionary"]);
			ContentUnitDsd dsd = ContentUnitDsdDao.getInstance().findById(unit.Id);

			// 更新單元範本
			unit.Name = nodeName;
			unit.CatNameMemo = catNameMemo;
			unit.Kind = (nodeType == "CP" ? ContentUnit.MODULE_CONTENT : ContentUnit.MODULE_LIST);

			ContentUnitDao.getInstance().update(unit);

			// 更新單元範本XML
			IList avaliableFields = getAvailiableFields();

			dsd.UnitId = unit.Id;
			dsd.BaseDad = unit.BaseDsd;

			foreach (DictionaryEntry entry in avaliableFields)
			{
				dsd.Tables["CuDtGeneric"].Fields[(string)entry.Key].IsShowListClient = listFields.Contains((string)entry.Key);
				dsd.Tables["CuDtGeneric"].Fields[(string)entry.Key].IsFormListClient = contentFields.Contains((string)entry.Key);
			}

			if (dataCategoryType == "CUSTOM")
			{
				// 個人化資料大類代碼

				dsd.FormClientCat = "CustomWebCat_" + unit.Id;
				dsd.Tables["CuDtGeneric"].Fields["topCat"].RefLookup = "CustomWebCat_" + unit.Id;

				CodeMainDataSetTableAdapters.CodeMetaDefTableAdapter metaAdapter = new CodeMainDataSetTableAdapters.CodeMetaDefTableAdapter();
				CodeMainDataSetTableAdapters.CodeMainTableAdapter mainAdapter = new CodeMainDataSetTableAdapters.CodeMainTableAdapter();

				// 先刪除個人化資料大類
				metaAdapter.Delete("CustomWebCat_" + unit.Id);
				mainAdapter.DeleteByMeta("CustomWebCat_" + unit.Id);

				// 新增個人化資料大類META
				metaAdapter.Insert("CustomWebCat_" + unit.Id, "個人化資料大類", "CodeMain", "codeMetaId", "CustomWebCat_" + unit.Id, "mcode", "mvalue", "msortValue", "代碼值", "顯示值", "排序值", "主題網", null, "Y", null);

				// 新增個人化資料大類MAIN.
                //2011/05/24 sam_chem 修正存檔時應該要分key跟value 
                int categorySoryOrder = 0;
                foreach (DictionaryEntry categoryItem in dataCategories)
                {
                    mainAdapter.Insert("CustomWebCat_" + unit.Id, (string)categoryItem.Key, (string)categoryItem.Value, categorySoryOrder.ToString());
                    categorySoryOrder++;
                }
                //for (int i = 0; i < dataCategories.Count; i++)
                //{
                //    mainAdapter.Insert("CustomWebCat_" + unit.Id, (string)dataCategories[i], (string)dataCategories[i], i.ToString());
                //}
			}
			else
			{
				// 預設資料大類代碼

				dsd.FormClientCat = dataCategoryType;
				dsd.Tables["CuDtGeneric"].Fields["topCat"].RefLookup = dataCategoryType;
			}

			ContentUnitDsdDao.getInstance().update(dsd);

			// 更新節點資料
			node.Name = nodeName;
			node.CatNameMemo = catNameMemo;
			node.InUse = isOpen;
			node.XslList = listStyle;
			node.XslData = contentStyle;
			node.ModifyUser = userId;
			node.ModifyDate = DateTime.Now;

			CatelogTreeNodeDao.getInstance().update(node);

			return node;
		}

		// 取得項目節點
		public CatelogNode getCatelogNode(int nodeId)
		{
			CatelogTreeNode node = (CatelogTreeNode)CatelogTreeNodeDao.getInstance().findById(nodeId);
			ContentUnit unit = (ContentUnit)ContentUnitDao.getInstance().get(node.UnitId);
			ContentUnitDsdDao.XmlPath = System.Web.HttpContext.Current.Server.MapPath(System.Configuration.ConfigurationManager.AppSettings["GipDsdDictionary"]);
			ContentUnitDsd dsd = ContentUnitDsdDao.getInstance().findById(node.UnitId.Value);

			CatelogNode catelogNode = new CatelogNode();
			catelogNode.Id = node.Id;
			catelogNode.Name = node.Name;
			catelogNode.Kind = CatelogNode.UNIT;
			catelogNode.InUse = node.InUse;
			catelogNode.UnitId = node.UnitId.Value;
			if (unit.Kind == ContentUnit.MODULE_CONTENT)
				catelogNode.UnitKind = "CP";
			else if (unit.Kind == ContentUnit.MODULE_LIST)
				catelogNode.UnitKind = "LP";
			else
				catelogNode.UnitKind = "";
			if(dsd.FormClientCat != String.Empty)
				catelogNode.DataCategoryType = dsd.FormClientCat.StartsWith("CustomWebCat_") ? "CUSTOM" : dsd.FormClientCat;
			catelogNode.ListStyle = node.XslList;
			catelogNode.ContentStyle = node.XslData;

			foreach (KeyValuePair<string, ContentUnitDsdField> entry in dsd.Tables["CuDtGeneric"].Fields)
			{
				if (entry.Value.IsShowListClient)
					catelogNode.ListFields.Add(entry.Key);
				if(entry.Value.IsFormListClient)
					catelogNode.ContentFields.Add(entry.Key);
			}
			
			catelogNode.CatNameMemo = node.CatNameMemo;

			return catelogNode;
		}

		// 刪除項目節點
		public void deleteCatelogNode(int nodeId)
		{
			CatelogTreeNode node = (CatelogTreeNode)CatelogTreeNodeDao.getInstance().findById(nodeId);

			// 檢查是否有資料
			ContentDataSetTableAdapters.CuDtGenericTableAdapter genericAdapter = new ContentDataSetTableAdapters.CuDtGenericTableAdapter();
			int count = (int)genericAdapter.CountContentByUnit(node.UnitId.Value);
			if (count > 0)
				throw new Exception("該單元下有資料，無法刪除。");

			ContentUnit unit = (ContentUnit)ContentUnitDao.getInstance().get(node.UnitId);
			ContentUnitDsdDao.XmlPath = System.Web.HttpContext.Current.Server.MapPath(System.Configuration.ConfigurationManager.AppSettings["GipDsdDictionary"]);
			ContentUnitDsd dsd = ContentUnitDsdDao.getInstance().findById(node.UnitId.Value);

			//ContentUnitDsdDao.getInstance().delete(dsd);
			ContentUnitDao.getInstance().delete(unit);
			CatelogTreeNodeDao.getInstance().delete(node);

			CtUnitDataSetTableAdapters.CtUserSetTableAdapter userSetAdapter = new CtUnitDataSetTableAdapters.CtUserSetTableAdapter();
			userSetAdapter.Delete(node.CreateUser, node.Id);
		}

		/*
		 * 以下暫時無用
		 */

		public CatelogTreeNode getCatelogTreeNodeRootByUser(string userId)
		{
			IDictionary filter = new Hashtable();
			filter.Add("ctRootId", 4);
			filter.Add("ctNodeKind", AbstractCatelog.FOLDER);
			filter.Add("dataParent", null);
			filter.Add("editUserId", userId);

			IList nodes = CatelogTreeNodeDao.getInstance().findByFilter(filter);

			if (nodes.Count > 0)
			{
				return (CatelogTreeNode)nodes[0];
			}

			return null;
		}

		public AbstractCatelog transformToAbstractCatelog(CatelogTreeNode node)
		{
			AbstractCatelog catelog = null;

			if (node.Kind.Equals(AbstractCatelog.FOLDER))
			{
				catelog = new CatelogFolder();
				catelog.Id = node.Id;
				catelog.Kind = node.Kind;
				catelog.Name = node.Name;
				catelog.Order = node.Order;
				catelog.ParentId = node.ParentId;
				catelog.CreateUser = node.CreateUser;
				catelog.CreateDate = node.CreateDate;
				catelog.ModifyUser = node.ModifyUser;
				catelog.ModifyDate = node.ModifyDate;
			}
			else if (node.Kind.Equals(AbstractCatelog.UNIT))
			{
				catelog = new CatelogNode();
				catelog.Id = node.Id;
				catelog.Kind = node.Kind;
				catelog.Name = node.Name;
				catelog.Order = node.Order;
				catelog.ParentId = node.ParentId;
				catelog.CreateUser = node.CreateUser;
				catelog.CreateDate = node.CreateDate;
				catelog.ModifyUser = node.ModifyUser;
				catelog.ModifyDate = node.ModifyDate;
			}

			return catelog;
		}

	}
}