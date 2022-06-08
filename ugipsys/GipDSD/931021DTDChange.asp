<%@ CodePage = 65001 %>
<%
'-----931021特定分眾修改預設勾選
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

'----修改DTD/DSD xml預設勾選
Set fso = server.CreateObject("Scripting.FileSystemObject")
xPath = "/site/"+session("mySiteID")+"/GipDSD/"
Set fldr = fso.GetFolder(server.MapPath(xPath))
	
if fldr.Files.count > 0 then
    for each sf in fldr.Files
    	if (Left(sf.name,5)="CuDTx" or Left(sf.name,7)="CtUnitX") and Left(sf.name,12)<>"CtUnitXOrder" then
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
'----刪除特定分眾
		set CuDTGenericNode=htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDTGeneric']/fieldList")		
		if nullText(CuDTGenericNode.selectSingleNode("field[fieldName='vGroup']"))<>"" then
			set fieldNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='vGroup']")
			CuDTGenericNode.removeChild fieldNode
		end if
'----回存  
		htPageDom.save(server.MapPath(xPath+sf.name))
		response.write sf.name+"<br>"		
   	end if
    next
end if

response.write "<br>DONE931021!"
%>