<%@ CodePage = 65001 %>
<% 
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
'------ Modify History List (begin) ------
' 2006/1/6	92004/Chirs
'	1. check CtNodeX???.xml before CtUnitX???.xml
'
'------ Modify History List (begin) ------
HTProgCap="資料上稿"
HTProgCode="GC1AP1"
HTProgPrefix="DsdXML" 
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbfunc.inc" -->
<%	
	session("SortFlag")=false
	session("CodeID")=request.querystring("iBaseDSD")
	if 1=2 then 	session("SortFlag")=true
	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true		
    	'----找出對應的CtUnitX???? xmlSpec檔案(若找不到則抓default), 並依fieldSeq排序成物件存入session
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtNodeX" & request("CtNodeId") & ".xml")
if not fso.FileExists(LoadXML) then
	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(session("CtUnitID")) & ".xml")
	if not fso.FileExists(LoadXML) then
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & session("iBaseDSD") & ".xml")
	end if
end if
	'response.write LoadXML & "<HR>"
	'response.end
	xv = htPageDom.load(LoadXML)

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
'response.write "DataSchemaDef/dsTable[tableName='"&session("sBaseTableName")&"']"
'response.end 	
	set xDSDNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='"&session("sBaseTableName")&"']")    
	set xDSDNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable")    
'response.write LoadXML 
'response.write xDSDNode.xml
'response.end 	
	set DSDNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='" & session("sBaseTableName") & "']").cloneNode(true)    
    set DSDNodeXML = server.createObject("microsoft.XMLDOM")
   		DSDNodeXML.appendchild DSDNode
    set nxml = server.createObject("microsoft.XMLDOM")
    	nxml.LoadXML(DSDNodeXML.transformNode(oxsl))
    set nxmlnewNode = nxml.documentElement    
    	DSDNode.replaceChild nxmlnewNode,DSDNode.selectSingleNode("fieldList")
    	root.replaceChild DSDNode,root.selectSingleNode("dsTable[tableName='"&session("sBaseTableName")&"']")
    '----複製CuDTGeneric的dsTable,並依順序轉換
    set GenericNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDtGeneric']").cloneNode(true)    
    set GenericNodeXML = server.createObject("microsoft.XMLDOM")
    	GenericNodeXML.appendchild GenericNode
   	set nxml2 = server.createObject("microsoft.XMLDOM")
    	nxml2.LoadXML(GenericNodeXML.transformNode(oxsl))
    set nxmlnewNode2 = nxml2.documentElement    
    	GenericNode.replaceChild nxmlnewNode2,GenericNode.selectSingleNode("fieldList")
    	root.replaceChild GenericNode,root.selectSingleNode("dsTable[tableName='CuDtGeneric']")       	

  	set session("codeXMLSpec") = htPageDom
  	'----混合field順序
	set nxml0 = server.createObject("microsoft.XMLDOM")
	nxml0.LoadXML(htPageDom.transformNode(oxsl))
	set session("codeXMLSpec2") = nxml0
'		response.write "<?xml version=""1.0"" encoding=""utf-8"" ?>"
'    response.write "<XMP>"+nxml0.xml+"</XMP>" 
'    response.end    	
%>
<script language=vbs>
'	window.navigate "DsdXMLList.asp?iBaseDSD=<%=session("CodeID")%>&ctNodeId=<%=request.querystring("ctNodeId")%>"
'	window.navigate "DsdXMLAdd.asp"
	window.navigate "old_DsdXMLList.asp"
</script>
<%  	response.end
%>
