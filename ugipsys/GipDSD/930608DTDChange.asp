<%@ CodePage = 65001 %>
<%
'-----930608多向出版修改
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
'    	if sf.name="CtUnitX221.xml"  then
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
  		for each fieldNode in htPageDom.selectNodes("//fieldList/field")
  		    '----加fieldRefEditYN
  		    if nullText(fieldNode.selectSingleNode("fieldRefEditYN"))="" then	
    			set nxml = server.createObject("microsoft.XMLDOM")
  		      	if fieldNode.selectSingleNode("fieldName").text="avBegin" or fieldNode.selectSingleNode("fieldName").text="avEnd" then
     				nxml.LoadXML("<fieldRefEditYN>Y</fieldRefEditYN>")
 		     	else
    				nxml.LoadXML("<fieldRefEditYN>N</fieldRefEditYN>")
    		      	end if 	
    			set nxmlnewNode = nxml.documentElement
    			fieldNode.appendChild nxmlnewNode
  		    end if
  		next
'----回存  		
		htPageDom.save(server.MapPath(xPath+sf.name))
		response.write sf.name+"<br>"
   	end if
    next
end if

response.write "<br>DONE3!"
%>