<%@ CodePage = 65001 %>
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

Set fso = server.CreateObject("Scripting.FileSystemObject")

xPath = "/site/TCsafe/GipDSD/"

Set fldr = fso.GetFolder(server.MapPath(xPath))
	
if fldr.Files.count > 0 then
    for each sf in fldr.Files
    	if Left(sf.name,5)="CuDTx" or Left(sf.name,7)="CtUnitX" then
'    		on error resume next
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
'----930423修改fileDownLoad.seq=05/xPostDateE改成 xPostDateEnd 		
	    	if nullText(htPageDom.selectSingleNode("//fieldList/field[fieldName='fileDownLoad']"))<>"" then	
			Set fieldNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='fileDownLoad']")
  			fieldNode.selectSingleNode("fieldSeq").text = "05"
  			fieldNode.selectSingleNode("formList").text = "05"
  		end if
	    	if nullText(htPageDom.selectSingleNode("//fieldList/field[fieldName='xPostDateE']"))<>"" then  		
			Set fieldNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='xPostDateE']")		
  			fieldNode.selectSingleNode("fieldName").text = "xPostDateEnd"
  		end if
'----930426修改showType/fileDownLoad/refID多組問題		
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
'----回存  		
		htPageDom.save(server.MapPath(xPath+sf.name))
		response.write sf.name+"<br>"
   	end if
    next
end if

'response.write "<XMP>"+htPageDom.xml+"</XMP>"
'response.end



response.write "<br>DONE!"
%>