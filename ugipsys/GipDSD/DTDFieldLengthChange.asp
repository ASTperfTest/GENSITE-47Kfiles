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
'    	if sf.name="CuDTx4.xml" then
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
'----930417修改xKeyword長度  		
'		Set fieldNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='xKeyword']")	
'		response.write fieldNode.nodetypeString+"<br>"
'		
'  			fieldNode.selectSingleNode("dataLen").text = 80
'  			fieldNode.selectSingleNode("inputLen").text = 60
'  			htPageDom.save(server.MapPath(xPath+sf.name))
'  			response.write sf.name+"<br>"

'----930413GipEdit修改
  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")	
	set CuDTGenericNode=htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDTGeneric']/fieldList")
	'----修改sTitle之showListClient/showList值為空
	    if nullText(allModel.selectSingleNode("//fieldList/field[fieldName='sTitle']"))<>"" then	    	
		Set fieldNode = allModel.selectSingleNode("//fieldList/field[fieldName='sTitle']")
  		fieldNode.selectSingleNode("showListClient").text = ""
  		fieldNode.selectSingleNode("showList").text = ""
  	    end if
	'----修改dEditDate(刪除updateDefault/增加DBSave/資料與輸入長度為20)
	    if nullText(allModel.selectSingleNode("//fieldList/field[fieldName='dEditDate']"))<>"" then	
		set dEditDateNode=allModel.selectSingleNode("//fieldList/field[fieldName='dEditDate']")
		if nullText(dEditDateNode.selectSingleNode("updateDefault"))<>"" then					
			dEditDateNode.removeChild dEditDateNode.selectSingleNode("updateDefault")		
		end if
		dEditDateNode.selectSingleNode("dataLen").text = 20
		dEditDateNode.selectSingleNode("inputLen").text = 20		
    		set nxml2 = server.createObject("microsoft.XMLDOM")
    		nxml2.LoadXML("<DBSave><type>SQLFunc</type><set>getdate()</set></DBSave>")
    		set nxmlnewNode2 = nxml2.documentElement 	
    		dEditDateNode.appendChild nxmlnewNode2			
	    end if		
	'----修改iEditor(刪除updateDefault/增加DBSave)
	    if nullText(allModel.selectSingleNode("//fieldList/field[fieldName='iEditor']"))<>"" then	
		set iEditorNode=allModel.selectSingleNode("//fieldList/field[fieldName='iEditor']")
		if nullText(iEditorNode.selectSingleNode("updateDefault"))<>"" then						
			iEditorNode.removeChild iEditorNode.selectSingleNode("updateDefault")		
		end if		
    		set nxml2 = server.createObject("microsoft.XMLDOM")
    		nxml2.LoadXML("<DBSave><type>session</type><set>userID</set></DBSave>")
    		set nxmlnewNode2 = nxml2.documentElement 	
    		iEditorNode.appendChild nxmlnewNode2			
	    end if		
	'----修改Created_Date(刪除updateDefault/inputType改為hidden/增加clientDefault/增加DBSave/formListClient與formList值補上/資料與輸入長度為20)
	    if nullText(allModel.selectSingleNode("//fieldList/field[fieldName='Created_Date']"))<>"" then	
		set CreatedDateNode=allModel.selectSingleNode("//fieldList/field[fieldName='Created_Date']")
		if nullText(CreatedDateNode.selectSingleNode("updateDefault"))<>"" then			
			CreatedDateNode.removeChild CreatedDateNode.selectSingleNode("updateDefault")		
		end if		
		CreatedDateNode.selectSingleNode("dataLen").text = 20
		CreatedDateNode.selectSingleNode("inputLen").text = 20		
		CreatedDateNode.selectSingleNode("inputType").text = "hidden"
		CreatedDateNode.selectSingleNode("formListClient").text = CreatedDateNode.selectSingleNode("fieldSeq").text
		CreatedDateNode.selectSingleNode("formList").text = CreatedDateNode.selectSingleNode("fieldSeq").text
    		set nxml = server.createObject("microsoft.XMLDOM")
    		nxml.LoadXML("<clientDefault><type>clientFunc</type><set>date()</set></clientDefault>")
    		set nxmlnewNode = nxml.documentElement 	
    		CreatedDateNode.appendChild nxmlnewNode			
    		set nxml2 = server.createObject("microsoft.XMLDOM")
    		nxml2.LoadXML("<DBSave><type>DoNothing</type><set></set></DBSave>")
    		set nxmlnewNode2 = nxml2.documentElement 	
    		CreatedDateNode.appendChild nxmlnewNode2			
	    end if		
	'----加3個field.....
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
		htPageDom.save(server.MapPath(xPath+sf.name))    			
    	end if
    next
end if

'response.write "<XMP>"+htPageDom.xml+"</XMP>"
'response.end



response.write "<br>DONE!"
%>