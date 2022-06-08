<%@ CodePage = 65001 %>
<% Response.Expires = 0
'------ Modify History List (begin) ------
' 2006/1/6	92004/Chirs
'	1. CtNodeDND.asp created from /GipDSD/CtUnitDSD.asp
'
'------ Modify History List (begin) ------
HTProgCap="目錄樹管理"
HTProgFunc="欄位選取更新"
HTProgCode="GE1T21"
HTProgPrefix="CtUnitNode" %>
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

	sqlCom = "SELECT htx.*, u.ibaseDsd " _
		& " ,(Select count(*) from CatTreeNode where dataParent=htx.ctNodeId) childCount" _
		& " FROM CatTreeNode AS htx JOIN CtUnit AS u ON u.ctUnitId=htx.ctUnitId" _
		& " JOIN BaseDsd AS b ON b.ibaseDsd=u.ibaseDsd" _
		& " WHERE htx.ctNodeId=" & pkStr(request.queryString("ctNodeId"),"")
	Set RSreg = Conn.execute(sqlcom)
	if RSreg.eof then
			showDoneBox("資料不存在！")
			response.end
	end if

	myNodeID = RSreg("ctNodeId")
	myUnitID = RSreg("ctUnitId")
	myBaseID = RSreg("ibaseDsd")
	
'----1.0依ctNodeID檢查/GipDSD/xmlSpec下是否存在CtNodeX???.xml(???=ctNodeID)
'----1.1依ctUnitID檢查/GipDSD/xmlSpec下是否存在CtUnitX???.xml(???=ctUnitID)
'------1.1.1若不存在, GipDSD/CuDTx???(???=baseDSD)
'----2.1開XMLDOM物件, load xml
'----3.1PSetListEdit方式批次編修

'----1.0xmlSpec檔案檢查
Set fso = server.CreateObject("Scripting.FileSystemObject")
LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtNodeX" & myNodeID & ".xml")
if not fso.FileExists(LoadXML) then
	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & myUnitID & ".xml")
	if not fso.FileExists(LoadXML) then
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & myBaseID & ".xml")
	end if
end if

set oxml = server.createObject("microsoft.XMLDOM")
oxml.async = false
oxml.setProperty "ServerHTTPRequest", true

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
oxml.save(server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtNodeX" & myNodeId & ".xml"))
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
	window.navigate "ctUnitNodeEdit.asp?CtNodeID=<%=myNodeId%>&phase=edit"
</script>
     </body>
     </html>
