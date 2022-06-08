using System;
using System.Data;

/// <summary>
/// ContentUnitDao 的摘要描述
/// </summary>
public class ContentUnitDao
{
	private static ContentUnitDao _instance;

	private ContentUnitDao()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}

	public static ContentUnitDao getInstance()
	{
		if (_instance == null)
			_instance = new ContentUnitDao();
		return _instance;
	}

	public Object get(Object id)
	{
		CtUnitDataSetTableAdapters.CtUnitTableAdapter adapter = new CtUnitDataSetTableAdapters.CtUnitTableAdapter();
		CtUnitDataSet.CtUnitDataTable table = adapter.GetDataById(Convert.ToInt32(id));
		return populateFromDataRow(table.Rows[0]);
	}

	public void insert(Object vo)
	{
		ContentUnit obj = (ContentUnit)vo;
		CtUnitDataSetTableAdapters.CtUnitTableAdapter adapter = new CtUnitDataSetTableAdapters.CtUnitTableAdapter();
		int ctUnitId;
		adapter.Insert(obj.Name, null, null, null, null, null, null, obj.Kind, obj.RedirectUrl, (obj.NewWindow ? "Y" : "N"), obj.BaseDsd, (obj.UnitOnly ? "Y" : "N"), (obj.InUse ? "Y" : "N"), null, (obj.CheckYN ? "Y" : "N"), obj.DeptId, null, null, out ctUnitId);
		obj.Id = ctUnitId;
	}

	public void update(Object vo)
	{
		ContentUnit obj = (ContentUnit)vo;
		CtUnitDataSetTableAdapters.CtUnitTableAdapter adapter = new CtUnitDataSetTableAdapters.CtUnitTableAdapter();
		adapter.Update(obj.Name, null, null, null, null, null, null, obj.Kind, obj.RedirectUrl, (obj.NewWindow ? "Y" : "N"), obj.BaseDsd, (obj.UnitOnly ? "Y" : "N"), (obj.InUse ? "Y" : "N"), null, (obj.CheckYN ? "Y" : "N"), obj.DeptId, null, null, obj.Id);
	}

	public void delete(Object vo)
	{
		ContentUnit obj = (ContentUnit)vo;
		CtUnitDataSetTableAdapters.CtUnitTableAdapter adapter = new CtUnitDataSetTableAdapters.CtUnitTableAdapter();
		adapter.Delete(obj.Id);
	}

	public ContentUnit populateFromDataRow(System.Data.DataRow row)
	{
		CtUnitDataSet.CtUnitRow concreteRow = (CtUnitDataSet.CtUnitRow)row;

		ContentUnit obj = new ContentUnit();
		obj.Id = concreteRow.ctUnitId;
		obj.Name = concreteRow.ctUnitName;
		obj.Kind = concreteRow.ctUnitKind;
		obj.RedirectUrl = concreteRow.redirectUrl;
		obj.NewWindow = concreteRow.newWindow == "Y" ? true : false;
		obj.BaseDsd = concreteRow.ibaseDsd;
		obj.InUse = concreteRow.inUse == "Y" ? true : false;
		obj.DeptId = concreteRow.deptId;

		return obj;
	}

}
