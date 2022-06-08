<%@ CodePage = 65001 %>
<%
'-----930511新聞局xmlSpec總修改
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

Set fso = server.CreateObject("Scripting.FileSystemObject")

xPath = "/site/"+session("mySiteID")+"/GipDSD/"
Set fldr = fso.GetFolder(server.MapPath(xPath))
	
if fldr.Files.count > 0 then
    for each sf in fldr.Files
    	if (Left(sf.name,5)="CuDTx" or Left(sf.name,7)="CtUnitX") and Left(sf.name,12)<>"CtUnitXOrder" then
'    	if sf.name="930510新聞局216.xml" or sf.name="930510新聞局235.xml" then
		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true            	
		LoadXML = server.MapPath(xPath+sf.name)            	
    		xv = htPageDom.load(LoadXML)
  		if htPageDom.parseError.reason <> "" then 
    			Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    			Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    			Response.End()
  		end if 
'----修改xKeyword長度  		
	    	if nullText(htPageDom.selectSingleNode("//fieldList/field[fieldName='xKeyword']"))<>"" then	
			Set fieldNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='xKeyword']")			
  			fieldNode.selectSingleNode("dataLen").text = 80
  			fieldNode.selectSingleNode("inputLen").text = 60
		end if
'----iEditor中clientDefault多組問題
	    	if nullText(htPageDom.selectSingleNode("//fieldList/field[fieldName='iEditor']"))<>"" then	
			Set fieldNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='iEditor']")			
			for each param in fieldNode.selectNodes("field[fieldName='clientDefault']")	
				fieldNode.removeChild param
			next
			for each param in fieldNode.selectNodes("field[fieldName='DBSave']")	
				fieldNode.removeChild param
			next
    			set nxml = server.createObject("microsoft.XMLDOM")
    			nxml.LoadXML("<clientDefault><type>clientFunc</type><set>date()</set></clientDefault>")
    			set nxmlnewNode = nxml.documentElement 	
    			fieldNode.appendChild nxmlnewNode			
    			set nxml2 = server.createObject("microsoft.XMLDOM")
    			nxml2.LoadXML("<DBSave><type>DoNothing</type><set></set></DBSave>")
    			set nxmlnewNode2 = nxml2.documentElement 	
    			fieldNode.appendChild nxmlnewNode2			
		end if		
'----iEditor中clientDefault多組問題
	    	if nullText(htPageDom.selectSingleNode("//fieldList/field[fieldName='iEditor']"))<>"" then	
			Set fieldNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='iEditor']")			
			for each param in fieldNode.selectNodes("clientDefault")	
				fieldNode.removeChild param
			next
			for each param in fieldNode.selectNodes("DBSave")	
				fieldNode.removeChild param
			next
    			set nxml = server.createObject("microsoft.XMLDOM")
    			nxml.LoadXML("<clientDefault><type>session</type><set>userID</set></clientDefault>")
    			set nxmlnewNode = nxml.documentElement 	
    			fieldNode.appendChild nxmlnewNode			
    			set nxml2 = server.createObject("microsoft.XMLDOM")
    			nxml2.LoadXML("<DBSave><type>session</type><set>userID</set></DBSave>")
    			set nxmlnewNode2 = nxml2.documentElement 	
    			fieldNode.appendChild nxmlnewNode2			
		end if		
'----dEditDate中clientDefault多組問題
	    	if nullText(htPageDom.selectSingleNode("//fieldList/field[fieldName='dEditDate']"))<>"" then	
			Set fieldNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='dEditDate']")			
			for each param in fieldNode.selectNodes("clientDefault")	
				fieldNode.removeChild param
			next
			for each param in fieldNode.selectNodes("DBSave")	
				fieldNode.removeChild param
			next
    			set nxml = server.createObject("microsoft.XMLDOM")
    			nxml.LoadXML("<clientDefault><type>clientFunc</type><set>date()</set></clientDefault>")
    			set nxmlnewNode = nxml.documentElement 	
    			fieldNode.appendChild nxmlnewNode			
    			set nxml2 = server.createObject("microsoft.XMLDOM")
    			nxml2.LoadXML("<DBSave><type>SQLFunc</type><set>getdate()</set></DBSave>")
    			set nxmlnewNode2 = nxml2.documentElement 	
    			fieldNode.appendChild nxmlnewNode2			
		end if		
'----Created_Date中clientDefault多組問題
	    	if nullText(htPageDom.selectSingleNode("//fieldList/field[fieldName='Created_Date']"))<>"" then	
			Set fieldNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='Created_Date']")			
			for each param in fieldNode.selectNodes("clientDefault")	
				fieldNode.removeChild param
			next
			for each param in fieldNode.selectNodes("DBSave")	
				fieldNode.removeChild param
			next
    			set nxml = server.createObject("microsoft.XMLDOM")
    			nxml.LoadXML("<clientDefault><type>clientFunc</type><set>date()</set></clientDefault>")
    			set nxmlnewNode = nxml.documentElement 	
    			fieldNode.appendChild nxmlnewNode			
    			set nxml2 = server.createObject("microsoft.XMLDOM")
    			nxml2.LoadXML("<DBSave><type>DoNothing</type><set></set></DBSave>")
    			set nxmlnewNode2 = nxml2.documentElement 	
    			fieldNode.appendChild nxmlnewNode2			
		end if		
'----修改showType/fileDownLoad/refID多組問題		
		set CuDTGenericNode=htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDTGeneric']/fieldList")
		for each param in CuDTGenericNode.selectNodes("field[fieldName='showType']")	
			CuDTGenericNode.removeChild param
		next
		for each param in CuDTGenericNode.selectNodes("field[fieldName='fileDownLoad']")	
			CuDTGenericNode.removeChild param
		next
		for each param in CuDTGenericNode.selectNodes("field[fieldName='refID']")	
			CuDTGenericNode.removeChild param
		next
		set fxml = server.createObject("microsoft.XMLDOM")
		fxml.async = false
		fxml.setProperty "ServerHTTPRequest", true
		LoadXML = server.mappath("/GipDSD/schemaField2.xml")
		xv = fxml.load(LoadXML)
  		if fxml.parseError.reason <> "" then 
    			Response.Write("fxml parseError on line " &  fxml.parseError.line)
    			Response.Write("<BR>Reasonxx: " &  fxml.parseError.reason)
    			Response.End()
  		end if  		
		set xFieldNode = fxml.selectSingleNode("//DataSchemaField").cloneNode(true)		
		for each param in xFieldNode.selectNodes("field")	
			CuDTGenericNode.appendChild param
		next 
		'----只有CtUnitX235與323xml之fileDownLoad.formList='05',其餘為空白
		if not (sf.name="CtUnitX235.xml" or sf.name="CtUnitX323.xml")	then
	    	    if nullText(htPageDom.selectSingleNode("//fieldList/field[fieldName='fileDownLoad']"))<>"" then	
			Set fieldNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='fileDownLoad']")			
  			fieldNode.selectSingleNode("formList").text = ""
		 end if		
		end if	
'----xPostDateE改成 xPostDateEnd 		
	    	if nullText(htPageDom.selectSingleNode("//fieldList/field[fieldName='xPostDateE']"))<>"" then  		
			Set fieldNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='xPostDateE']")		
  			fieldNode.selectSingleNode("fieldName").text = "xPostDateEnd"
  		end if
'----加xImgFile
	    	if nullText(htPageDom.selectSingleNode("//fieldList/field[fieldName='xImgFile']"))="" then	
		    set CuDTGenericNode=htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDTGeneric']/fieldList")
		    set fxml = server.createObject("microsoft.XMLDOM")
		    fxml.async = false
		    fxml.setProperty "ServerHTTPRequest", true
		    LoadXML = server.mappath("/GipDSD/schema0.xml")
		    xv = fxml.load(LoadXML)
  		    if fxml.parseError.reason <> "" then 
    			Response.Write("fxml parseError on line " &  fxml.parseError.line)
    			Response.Write("<BR>Reasonxx: " &  fxml.parseError.reason)
    			Response.End()
  		    end if 
		    set xFieldNode = fxml.selectSingleNode("//dsTable/fieldList/field[fieldName='xImgFile']").cloneNode(true)		
		    CuDTGenericNode.appendChild xFieldNode		
  		end if
'----加前台樣版呈現修改tag  			  		 		  		
		set root = htPageDom.selectSingleNode("DataSchemaDef")
		set formClientCatNode = root.selectNodes("formClientCat")
		if formClientCatNode.length = 0 then
    			set nxml0 = server.createObject("microsoft.XMLDOM")
    			nxml0.LoadXML("<formClientCat></formClientCat>")
    			set newNode = nxml0.documentElement 
    			root.insertBefore newNode,root.childNodes.Item(0)		
		end if
		set showClientSqlOrderByNode = root.selectNodes("showClientSqlOrderBy")
		if showClientSqlOrderByNode.length = 0 then
    			set nxml0 = server.createObject("microsoft.XMLDOM")
    			nxml0.LoadXML("<showClientSqlOrderBy></showClientSqlOrderBy>")
    			set newNode = nxml0.documentElement 
    			root.insertBefore newNode,root.childNodes.Item(0)		
		end if
		set formClientStyleNode = root.selectNodes("formClientStyle")
		if formClientStyleNode.length = 0 then
    			set nxml0 = server.createObject("microsoft.XMLDOM")
    			nxml0.LoadXML("<formClientStyle>STD</formClientStyle>")
    			set newNode = nxml0.documentElement 
    			root.insertBefore newNode,root.childNodes.Item(0)		
		end if
		set showClientStyleNode = root.selectNodes("showClientStyle")
		if showClientStyleNode.length = 0 then
    			set nxml0 = server.createObject("microsoft.XMLDOM")
    			nxml0.LoadXML("<showClientStyle>STD</showClientStyle>")
    			set newNode = nxml0.documentElement 
    			root.insertBefore newNode,root.childNodes.Item(0)		
		end if
'----回存  		
		htPageDom.save(server.MapPath(xPath+sf.name))
		response.write sf.name+"<br>"
   	end if
    next
end if


'----sp_rename 'CuDTGeneric.[xPostDateE]', 'xPostDateEnd', 'COLUMN'



response.write "<br>DONE3!"
%>