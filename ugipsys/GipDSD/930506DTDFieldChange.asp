<%@ CodePage = 65001 %>
<%
'-----930506修改DSD
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

'response.write "<XMP>"+htPageDom.xml+"</XMP>"
'response.end



response.write "<br>DONE!"
%>