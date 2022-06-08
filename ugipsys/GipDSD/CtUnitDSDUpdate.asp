﻿<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題單元管理"
HTProgFunc="欄位選取更新"
HTProgCode="GE1T11"
HTProgPrefix="CtUnitDSD" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<%
sub nodeCheck(xNode,xField,fieldValue,seqValue,funStr)
	set testNode=xNode.selectSingleNode(xField)
	if typeName(testNode) = "Nothing" then
		if fieldValue<>"" or (funStr="form" and xNode.selectSingleNode("canNull").text="N") then
			'----新增element且放入值
			set objNewNode=oxml.createElement(xField)
			set objTextNode=oxml.createTextNode(seqValue)
			objNewNode.appendChild(objTextNode)
			xNode.appendChild(objNewNode)
		else
			'----不做任何事
		end if
	else
		if fieldValue<>"" or (funStr="form" and xNode.selectSingleNode("canNull").text="N") then
			'----更新值
			testNode.text = seqValue
		else
			'----更新為空值
			testNode.text = ""
		end if	
	end if
end sub

set oxml = server.createObject("microsoft.XMLDOM")
oxml.async = false
oxml.setProperty "ServerHTTPRequest", true
'----xmlSpec檔案檢查
Set fso = server.CreateObject("Scripting.FileSystemObject")
LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX"+request("ctUnit")+".xml")
if not fso.FileExists(LoadXML) then
	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & request("baseDSD") & ".xml")
end if
xv = oxml.load(LoadXML)
if oxml.parseError.reason <> "" then
	Response.Write("XML parseError on line " &  oxml.parseError.line)
	Response.Write("<BR>Reason: " &  oxml.parseError.reason)
	Response.End()
end if 
set root = oxml.selectSingleNode("DataSchemaDef")

root.selectSingleNode("showClientStyle").text = request("showClientStyle")
root.selectSingleNode("formClientStyle").text = request("formClientStyle")
OrderByStr=""
if request("showClientSqlOrderBy")<>"" then
	OrderByStr=OrderByStr+"Order by "+request("showClientSqlOrderBy")
	if request("showClientSqlOrderByType")<>"" then
		OrderByStr=OrderByStr+" "+request("showClientSqlOrderByType")
	end if
end if

root.selectSingleNode("showClientSqlOrderBy").text = OrderByStr
root.selectSingleNode("formClientCat").text = request("formClientCat")
if request("formClientCat")<>"" then
	oxml.selectSingleNode("//DataSchemaDef/dsTable/fieldList/field[fieldName='"+root.selectSingleNode("formClientCat").text+"']/refLookup").text=request("formClientCatRefLookup")
end if
For Each x In Request.Form
    if left(x,12)="htx_fieldSeq" then    
       	xn=mid(x,13) 
	set fieldNode = oxml.selectSingleNode("//DataSchemaDef/dsTable/fieldList/field[fieldName='"+request("htx_fieldName"&xn)+"']")              
	nodeCheck fieldNode,"fieldSeq",request("htx_fieldSeq"&xn),request("htx_fieldSeq"&xn),""
	fieldNode.selectSingleNode("fieldLabel").text = request("htx_fieldLabel"&xn)
	fieldNode.selectSingleNode("fieldDesc").text = request("htx_fieldDesc"&xn)
	nodeCheck fieldNode,"showListClient",request("ckboxshowListClient"&xn),request("htx_fieldSeq"&xn),""
	nodeCheck fieldNode,"formListClient",request("ckboxformListClient"&xn),request("htx_fieldSeq"&xn),"formClient"
	nodeCheck fieldNode,"queryListClient",request("ckboxqueryListClient"&xn),request("htx_fieldSeq"&xn),""
	nodeCheck fieldNode,"showList",request("ckboxshowList"&xn),request("htx_fieldSeq"&xn),""
	nodeCheck fieldNode,"formList",request("ckboxformList"&xn),request("htx_fieldSeq"&xn),"form"
	nodeCheck fieldNode,"queryList",request("ckboxqueryList"&xn),request("htx_fieldSeq"&xn),""
    end if
Next
'response.write "<XMP>"+oxml.xml+"</XML>"
'response.end
oxml.save(server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX"+request("ctUnit")+".xml"))
'response.write "完成"
'response.end
%>
    <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>編修表單</title>
     </head>
     <body>
<script Language=VBScript>
	alert "編修完成！"
	window.navigate "ctUnitEdit.asp?CtUnitID=<%=request("ctUnit")%>&phase=edit"
</script>
     </body>
     </html>
