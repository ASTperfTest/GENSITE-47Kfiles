<%@ CodePage = 65001 %>
<%
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
	    	if nullText(CuDTGenericNode.selectSingleNode("field[fieldName='xImgFile']"))="" then  		
			CuDTGenericNode.appendChild xFieldNode		
'----回存  
			htPageDom.save(server.MapPath(xPath+sf.name))
			response.write sf.name+"<br>"
  		end if	  		 		


   	end if
    next
end if

'response.write "<XMP>"+htPageDom.xml+"</XMP>"
'response.end



response.write "<br>DONE!"
%>