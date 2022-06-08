<%@ CodePage = 65001 %>
<% 
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="資料審稿"
HTProgCode="GC1AP2"
HTProgPrefix="DsdXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%	
	session("SortFlag")=false
	session("CodeID")=request.querystring("iBaseDSD")
	if 1=2 then 	session("SortFlag")=true
	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true		
    	'----找出對應的CtUnitX???? xmlSpec檔案(若找不到則抓default), 並依fieldSeq排序成物件存入session
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(session("CtUnitID")) & ".xml") 	
    	if fso.FileExists(filePath) then
    		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(session("CtUnitID")) & ".xml")
    	else
    		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & session("iBaseDSD") & ".xml")
    	end if 
'	response.write LoadXML & "<HR>"
'	response.end
	xv = htPageDom.load(LoadXML)
'	response.write xv & "<HR>"
  	if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  	end if 	
    	set root = htPageDom.selectSingleNode("DataSchemaDef")
    	'----Load XSL樣板
    	set oxsl = server.createObject("microsoft.XMLDOM")
   	oxsl.async = false
   	xv = oxsl.load(server.mappath("/GipDSD/xmlspec/CtUnitXOrder.xsl"))    		
    	'----複製Slave的dsTable,並依順序轉換
	set DSDNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='"&session("sBaseTableName")&"']").cloneNode(true)    
    	set DSDNodeXML = server.createObject("microsoft.XMLDOM")
   	DSDNodeXML.appendchild DSDNode
    	set nxml = server.createObject("microsoft.XMLDOM")
    	nxml.LoadXML(DSDNodeXML.transformNode(oxsl))
    	set nxmlnewNode = nxml.documentElement    
    	DSDNode.replaceChild nxmlnewNode,DSDNode.selectSingleNode("fieldList")
    	root.replaceChild DSDNode,root.selectSingleNode("dsTable[tableName='"&session("sBaseTableName")&"']")
    	'----複製CuDTGeneric的dsTable,並依順序轉換
    	set GenericNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDTGeneric']").cloneNode(true)    
    	set GenericNodeXML = server.createObject("microsoft.XMLDOM")
    	GenericNodeXML.appendchild GenericNode
   	set nxml2 = server.createObject("microsoft.XMLDOM")
    	nxml2.LoadXML(GenericNodeXML.transformNode(oxsl))
    	set nxmlnewNode2 = nxml2.documentElement    
    	GenericNode.replaceChild nxmlnewNode2,GenericNode.selectSingleNode("fieldList")
    	root.replaceChild GenericNode,root.selectSingleNode("dsTable[tableName='CuDTGeneric']")       	


  	set session("codeXMLSpec") = htPageDom
  	'----混合field順序
	set nxml0 = server.createObject("microsoft.XMLDOM")
	nxml0.LoadXML(htPageDom.transformNode(oxsl))
	set session("codeXMLSpec2") = nxml0
'    response.write "<XMP>"+nxml0.xml+"</XMP>" 
'    response.end    	
%>
<script language=vbs>
	window.navigate "DsdXMLList.asp?iBaseDSD=<%=session("CodeID")%>"
'	window.navigate "DsdXMLAdd.asp"
</script>
<%  	response.end
%>
