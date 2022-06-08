using System;
using System.Data;

/// <summary>
/// CreateTable 的摘要描述
/// </summary>
public class CreateTable
{
	public CreateTable()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}
    
    public DataTable createMyDataSet()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("MenuTree", typeof(string));
        dt.Columns.Add("MpStyle", typeof(string));
        return dt;
    }
    public DataTable createDataSet()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("DataLable", typeof(string));
        dt.Columns.Add("DataRemark", typeof(string));
        dt.Columns.Add("DataNode", typeof(string));
        dt.Columns.Add("SqlCondition", typeof(string));
        dt.Columns.Add("SqlOrderBy", typeof(string));
        dt.Columns.Add("SqlTop", typeof(string));
        dt.Columns.Add("WithContent", typeof(string));
        dt.Columns.Add("ContentData", typeof(string));
        dt.Columns.Add("ContentLength", typeof(string));
        dt.Columns.Add("IsTitle", typeof(bool));
        dt.Columns.Add("IsPic", typeof(bool));
        dt.Columns.Add("IsPostDate", typeof(bool));
        dt.Columns.Add("IsExcerpt", typeof(bool));
        dt.Columns.Add("Type1", typeof(bool));
        dt.Columns.Add("Type2", typeof(bool));
        dt.Columns.Add("Type3", typeof(bool));
        return dt;
    }   
  
}
