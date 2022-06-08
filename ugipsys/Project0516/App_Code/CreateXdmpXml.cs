using System;
using System.Data;
using System.Xml;
using System.Web.UI.WebControls;

/// <summary>
/// CreateXdmpXml 的摘要描述
/// </summary>
public class CreateXdmpXml
{
	public CreateXdmpXml()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}
    //建置ｘｍｌ檔
    public void createXmlFile(DataSet ds, string FilePath )
    {        
                
        XmlWriterSettings settings = new XmlWriterSettings();
        settings.Indent = true;
        settings.IndentChars = ("    ");
        using (XmlWriter writer = XmlWriter.Create(FilePath, settings))
        {
            // Write XML data.
            writer.WriteStartElement("MpDataSet");
            writer.WriteElementString("MenuTree", ds.Tables[0].Rows[0]["MenuTree"].ToString());
            writer.WriteElementString("MpStyle", ds.Tables[0].Rows[0]["MpStyle"].ToString());

            foreach (DataRow dr in ds.Tables[1].Rows)
            {
                writer.WriteStartElement("DataSet");
                //原有gip的
                writer.WriteElementString("DataLable", dr["DataLable"].ToString());
                writer.WriteElementString("DataRemark", dr["DataRemark"].ToString());
                writer.WriteElementString("DataNode", dr["DataNode"].ToString());
                writer.WriteElementString("SqlCondition", dr["SqlCondition"].ToString());
                writer.WriteElementString("SqlOrderBy", dr["SqlOrderBy"].ToString());
                writer.WriteElementString("SqlTop", dr["SqlTop"].ToString());
                //WithContent??為何？
                writer.WriteElementString("WithContent", dr["WithContent"].ToString());
                writer.WriteElementString("ContentData", dr["ContentData"].ToString());
                writer.WriteElementString("ContentLength", dr["ContentLength"].ToString());
                //新增的欄位
                writer.WriteElementString("IsTitle", boolYN((bool)dr["IsTitle"]));
                writer.WriteElementString("IsPic", boolYN((bool)dr["IsPic"]));
                writer.WriteElementString("IsPostDate", boolYN((bool)dr["IsPostDate"]));
                writer.WriteElementString("IsExcerpt", boolYN((bool)dr["IsExcerpt"]));
                writer.WriteElementString("ShowStyle", style((bool)dr["Type1"], (bool)dr["Type2"], (bool)dr["Type3"]));
                writer.WriteElementString("mpShow", "Y");
                writer.WriteEndElement();

            }
            writer.WriteEndElement();
            writer.Flush();           
        }

    }

    public DataSet loadXmlFile(string FilePath)
    {        
        DataSet ds = new DataSet();
        ds.ReadXml(FilePath.ToString());        
        return ds;
    }
    

    public string boolYN(bool CheckedValue)
    {
        return (CheckedValue) ? "Y" : "N"; 
    }

    public string style(bool radiobutton1, bool radiobutton2, bool radiobutton3)
    {
        string message = "";
        if (radiobutton3)
        {
            message = "Style3";
        }
        else if (radiobutton2)
        {
            message = "Style2";
        }
        else
        {
            message = "Style1";
        }
        return message;
    }
}
