using System;
using System.Xml;
using System.Collections.Generic;

/// <summary>
/// ContentUnitDsdDao 的摘要描述
/// </summary>
public class ContentUnitDsdDao
{
	private static ContentUnitDsdDao _instance;
	public static string XmlPath;

	protected ContentUnitDsdDao()
	{
	}

	public static ContentUnitDsdDao getInstance()
	{
		if (_instance == null)
			_instance = new ContentUnitDsdDao();
		return _instance;
	}

	public ContentUnitDsd findById(int id)
	{
		XmlDocument doc = new XmlDocument();

		System.IO.Stream xmlFileStream;
		if (System.IO.File.Exists(XmlPath + "/CtUnitX" + id + ".xml"))
			xmlFileStream = System.IO.File.OpenRead(XmlPath + "/CtUnitX" + id + ".xml");
		else
			xmlFileStream = System.IO.File.OpenRead(XmlPath + "/CuDTxSUB.xml");

		doc.Load(xmlFileStream);

		xmlFileStream.Close();

		XmlNode root = doc.DocumentElement;

		Console.Write(root.OuterXml);

		ContentUnitDsd dsd = convertNodeToDsd(root);

		dsd.UnitId = id;

		return dsd;
	}

	public void insert(ContentUnitDsd vo)
	{
		XmlDocument doc = new XmlDocument();
		XmlDeclaration declaration = doc.CreateXmlDeclaration("1.0", "utf-8", null);
		doc.AppendChild(declaration);
		XmlNode node = doc.ImportNode(convertDsdToNode(vo), true);
		doc.AppendChild(node);

		string filePath = XmlPath + "/CtUnitX" + vo.UnitId + ".xml";
		if (System.IO.File.Exists(filePath))
		{
			throw new Exception("檔案已經存在");
		}

		doc.Save(filePath);
	}

	public void update(ContentUnitDsd vo)
	{
		XmlDocument doc = new XmlDocument();
		XmlDeclaration declaration = doc.CreateXmlDeclaration("1.0", "utf-8", null);
		doc.AppendChild(declaration);
		XmlNode node = doc.ImportNode(convertDsdToNode(vo), true);
		doc.AppendChild(node);
		doc.Save(XmlPath + "/CtUnitX" + vo.UnitId + ".xml");
	}

	public void delete(ContentUnitDsd vo)
	{
		string filePath = XmlPath + "/CtUnitX" + vo.UnitId + ".xml";
		System.IO.File.Delete(filePath);
	}

	private ContentUnitDsd convertNodeToDsd(XmlNode node)
	{
		ContentUnitDsd dsd = new ContentUnitDsd();

		dsd.ShowClientStyle = node.SelectSingleNode("showClientStyle").InnerText;
		dsd.FormClientStyle = node.SelectSingleNode("formClientStyle").InnerText;
		dsd.ShowClientSqlOrderBy = node.SelectSingleNode("showClientSqlOrderBy").InnerText;
        dsd.FormClientCat = node.SelectSingleNode("formClientCat").InnerText;
 //       dsd.showTypeStr = node.SelectSingleNode("showTypeStr").InnerText;
 //       xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
		XmlNodeList tableNodes = node.SelectNodes("dsTable");

		foreach (XmlNode tableNode in tableNodes)
		{
			ContentUnitDsdTable table = convertNodeToTable(tableNode);
			dsd.Tables.Add(table.Name, table);
		}

		return dsd;
	}

	private ContentUnitDsdTable convertNodeToTable(XmlNode node)
	{
		ContentUnitDsdTable table = new ContentUnitDsdTable();

		table.Name = node.SelectSingleNode("tableName").InnerText;
		table.Desc = node.SelectSingleNode("tableDesc").InnerText;

		XmlNodeList fieldNodes = node.SelectNodes("fieldList/field");

		foreach (XmlNode fieldNode in fieldNodes)
		{
			ContentUnitDsdField field = convertNodeToField(fieldNode);
			table.Fields.Add(field.Name, field);
		}

		return table;
	}

	private ContentUnitDsdField convertNodeToField(XmlNode node)
	{
		ContentUnitDsdField field = new ContentUnitDsdField();

		field.Seq = Convert.ToInt32(node.SelectSingleNode("fieldSeq").InnerText);
		field.Name = node.SelectSingleNode("fieldName").InnerText;
		field.OriLabel = node.SelectSingleNode("fieldLabelInit").InnerText;
		field.Label = node.SelectSingleNode("fieldLabel").InnerText;
		field.DataType = node.SelectSingleNode("dataType").InnerText;
		field.DataLength = Convert.ToInt32(node.SelectSingleNode("dataLen").InnerText);
		field.InputLength = Convert.ToInt32(node.SelectSingleNode("inputLen").InnerText);
		field.CanNull = node.SelectSingleNode("canNull").InnerText == "Y" ? true : false;
		field.IsPrimaryKey = node.SelectSingleNode("isPrimaryKey").InnerText == "Y" ? true : false;
		field.IsIdentity = node.SelectSingleNode("identity").InnerText == "Y" ? true : false;
		field.InputType = node.SelectSingleNode("inputType").InnerText;
		field.Desc = node.SelectSingleNode("fieldDesc").InnerText;
		field.RefLookup = node.SelectSingleNode("refLookup").InnerText;
		int rows;
		Int32.TryParse(node.SelectSingleNode("rows").InnerText, out rows);
		field.Rows = rows;
		int cols;
		Int32.TryParse(node.SelectSingleNode("cols").InnerText, out cols);
		field.Cols = cols;
		field.IsShowListClient = node.SelectSingleNode("showListClient").InnerText != String.Empty ? true : false;
		field.IsFormListClient = node.SelectSingleNode("formListClient").InnerText != String.Empty ? true : false;
		field.IsQueryListClient = node.SelectSingleNode("queryListClient").InnerText != String.Empty ? true : false;
		field.IsShowList = node.SelectSingleNode("showList").InnerText != String.Empty ? true : false;
		field.IsFormList = node.SelectSingleNode("formList").InnerText != String.Empty ? true : false;
		field.IsQueryList = node.SelectSingleNode("queryList").InnerText != String.Empty ? true : false;
		field.ParamKind = node.SelectSingleNode("paramKind").InnerText;
        if (node.SelectSingleNode("showTypeStr") != null)
        {
            field.showTypeStr = node.SelectSingleNode("showTypeStr").InnerText;
        }
        //XmlNode nodeshowTyptStr = node.SelectSingleNode("showTypeStr");
        //if (nodeshowTyptStr != null)
        //    field.showTypeStr = new System.Collections.DictionaryEntry(node.SelectSingleNode("showTypeStr").InnerText);
		XmlNode nodeClientDefault = node.SelectSingleNode("clientDefault");
		if(nodeClientDefault != null)
			field.ClientDefault = new System.Collections.DictionaryEntry(node.SelectSingleNode("clientDefault/type").InnerText, node.SelectSingleNode("clientDefault/set").InnerText);

		return field;
	}

	private XmlNode convertDsdToNode(ContentUnitDsd dsd)
	{
		XmlDocument doc = new XmlDocument();

		XmlElement elmDataSchemaDef = doc.CreateElement("DataSchemaDef");
		
		XmlElement elmShowClientStyle = doc.CreateElement("showClientStyle");
		elmShowClientStyle.InnerText = dsd.ShowClientStyle;
		elmDataSchemaDef.AppendChild(elmShowClientStyle);

		XmlElement elmFormClientStyle = doc.CreateElement("formClientStyle");
		elmFormClientStyle.InnerText = dsd.FormClientStyle;
		elmDataSchemaDef.AppendChild(elmFormClientStyle);

		XmlElement elmShowClientSqlOrderBy = doc.CreateElement("showClientSqlOrderBy");
		elmShowClientSqlOrderBy.InnerText = dsd.ShowClientSqlOrderBy;
		elmDataSchemaDef.AppendChild(elmShowClientSqlOrderBy);

		XmlElement elmFormClientCat = doc.CreateElement("formClientCat");
		elmFormClientCat.InnerText = dsd.FormClientCat;
		elmDataSchemaDef.AppendChild(elmFormClientCat);

		foreach (KeyValuePair<string, ContentUnitDsdTable> entry in dsd.Tables)
		{
			XmlNode node = doc.ImportNode(convertTableToNode(entry.Value), true);
			elmDataSchemaDef.AppendChild(node);
		}

		return elmDataSchemaDef;
	}

	private XmlNode convertTableToNode(ContentUnitDsdTable table)
	{
		XmlDocument doc = new XmlDocument();

		XmlElement elmTable = doc.CreateElement("dsTable");

		XmlElement elmTableName = doc.CreateElement("tableName");
		elmTableName.InnerText = table.Name;
		elmTable.AppendChild(elmTableName);

		XmlElement elmTableDesc = doc.CreateElement("tableDesc");
		elmTableDesc.InnerText = table.Desc;
		elmTable.AppendChild(elmTableDesc);

		XmlElement elmFieldList = doc.CreateElement("fieldList");

		foreach (KeyValuePair<string, ContentUnitDsdField> entry in table.Fields)
		{
			XmlNode node = doc.ImportNode(convertFieldToNode(entry.Value), true);
			elmFieldList.AppendChild(node);
		}

		elmTable.AppendChild(elmFieldList);

		return elmTable;
	}

	private XmlNode convertFieldToNode(ContentUnitDsdField field)
	{
		XmlDocument doc = new XmlDocument();

		XmlElement elmField = doc.CreateElement("field");

		XmlElement elmFieldSeq = doc.CreateElement("fieldSeq");
		elmFieldSeq.InnerText = Convert.ToString(field.Seq);
		elmField.AppendChild(elmFieldSeq);

		XmlElement elmFieldName = doc.CreateElement("fieldName");
		elmFieldName.InnerText = field.Name;
		elmField.AppendChild(elmFieldName);

		XmlElement elmFieldLabelInit = doc.CreateElement("fieldLabelInit");
		elmFieldLabelInit.InnerText = field.OriLabel;
		elmField.AppendChild(elmFieldLabelInit);

		XmlElement elmFieldLabel = doc.CreateElement("fieldLabel");
		elmFieldLabel.InnerText = field.Label;
		elmField.AppendChild(elmFieldLabel);

		XmlElement elmDataType = doc.CreateElement("dataType");
		elmDataType.InnerText = field.DataType;
		elmField.AppendChild(elmDataType);

		XmlElement elmDataLen = doc.CreateElement("dataLen");
		elmDataLen.InnerText = Convert.ToString(field.DataLength);
		elmField.AppendChild(elmDataLen);

		XmlElement elmInputLength = doc.CreateElement("inputLen");
		elmInputLength.InnerText = Convert.ToString(field.InputLength);
		elmField.AppendChild(elmInputLength);

		XmlElement elmCanNull = doc.CreateElement("canNull");
		elmCanNull.InnerText = field.CanNull ? "Y" : "N";
		elmField.AppendChild(elmCanNull);

		XmlElement elmIsPrimaryKey = doc.CreateElement("isPrimaryKey");
		elmIsPrimaryKey.InnerText = field.IsPrimaryKey ? "Y" : "N";
		elmField.AppendChild(elmIsPrimaryKey);

		XmlElement elmIsIdentity = doc.CreateElement("identity");
		elmIsIdentity.InnerText = field.IsIdentity ? "Y" : "N";
		elmField.AppendChild(elmIsIdentity);

		XmlElement elmInputType = doc.CreateElement("inputType");
		elmInputType.InnerText = field.InputType;
		elmField.AppendChild(elmInputType);

		XmlElement elmDesc = doc.CreateElement("fieldDesc");
		elmDesc.InnerText = field.Desc;
		elmField.AppendChild(elmDesc);

		XmlElement elmRefLookup = doc.CreateElement("refLookup");
		elmRefLookup.InnerText = field.RefLookup;
		elmField.AppendChild(elmRefLookup);

		XmlElement elmRows = doc.CreateElement("rows");
		elmRows.InnerText = Convert.ToString(field.Rows);
		elmField.AppendChild(elmRows);

		XmlElement elmCols = doc.CreateElement("cols");
		elmCols.InnerText = Convert.ToString(field.Cols);
		elmField.AppendChild(elmCols);

		XmlElement elmIsShowListClient = doc.CreateElement("showListClient");
		elmIsShowListClient.InnerText = field.IsShowListClient ? Convert.ToString(field.Seq) : String.Empty;
		elmField.AppendChild(elmIsShowListClient);

		XmlElement elmIsFormListClient = doc.CreateElement("formListClient");
		elmIsFormListClient.InnerText = field.IsFormListClient ? Convert.ToString(field.Seq) : String.Empty;
		elmField.AppendChild(elmIsFormListClient);

		XmlElement elmIsQueryListClient = doc.CreateElement("queryListClient");
		elmIsQueryListClient.InnerText = field.IsQueryListClient ? Convert.ToString(field.Seq) : String.Empty;
		elmField.AppendChild(elmIsQueryListClient);

		XmlElement elmIsShowList = doc.CreateElement("showList");
		elmIsShowList.InnerText = field.IsShowList ? Convert.ToString(field.Seq) : String.Empty;
		elmField.AppendChild(elmIsShowList);

		XmlElement elmIsFormList = doc.CreateElement("formList");
		elmIsFormList.InnerText = field.IsFormList ? Convert.ToString(field.Seq) : String.Empty;
		elmField.AppendChild(elmIsFormList);

		XmlElement elmIsQueryList = doc.CreateElement("queryList");
		elmIsQueryList.InnerText = field.IsQueryList ? Convert.ToString(field.Seq) : String.Empty;
		elmField.AppendChild(elmIsQueryList);

		XmlElement elmParamKind = doc.CreateElement("paramKind");
		elmParamKind.InnerText = field.ParamKind;
		elmField.AppendChild(elmParamKind);

		if (field.ClientDefault.Key != null)
		{
			XmlElement elmClientDefault = doc.CreateElement("clientDefault");

			XmlElement elmType = doc.CreateElement("type");
			elmType.InnerText = Convert.ToString(field.ClientDefault.Key);
			elmClientDefault.AppendChild(elmType);

			XmlElement elmSet = doc.CreateElement("set");
			elmSet.InnerText = Convert.ToString(field.ClientDefault.Value);
			elmClientDefault.AppendChild(elmSet);

			elmField.AppendChild(elmClientDefault);
		}

        if (field.showTypeStr != null)
        {
            XmlElement elmShowTypeStr = doc.CreateElement("showTypeStr");

            elmShowTypeStr.InnerText = field.showTypeStr;
            elmField.AppendChild(elmShowTypeStr);
        
        }

		return elmField;
	}
}
