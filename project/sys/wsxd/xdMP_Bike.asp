<%@ CodePage = 65001 %>
<% Response.AddHeader "content-type", "text/xml; charset=utf-8" %>
<?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->
<% 
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true
		'response.write session("mySiteID")
		'response.write server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & request("mp") & ".xml"
		'response.end
	
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & request("mp") & ".xml"
		xv = htPageDom.load(LoadXML)
'		response.write xv & "<HR>"
  		if htPageDom.parseError.reason <> "" then 
    		Response.Write(LoadXML & " file parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  		end if

  	set refModel = htPageDom.selectSingleNode("//MpDataSet")
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	response.write "<LoadXML>" & LoadXML & "</LoadXML>"
	response.write "<MenuTitle>"&nullText(refModel.selectSingleNode("MenuTitle"))&"</MenuTitle>"
	response.write "<mp>" & request("mp") & "</mp>"  	
	response.write "<myStyle>"&nullText(refModel.selectSingleNode("MpStyle"))&"</myStyle>"
		
	%>
<!--#include file="gensite.inc" -->
	<%
  	for each xDataSet in refModel.selectNodes("DataSet")
  		processXDataSet
	next

  myTreeNode = 0
  upParent = 0

function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function
'--------------頁尾維護單位---------------------
footer_sql = "select footer_dept, footer_dept_url from cattreeroot left join nodeinfo on cattreeroot.ctrootid = nodeinfo.ctrootid where cattreeroot.pvXdmp ='"& request("mp") &"'"
set footer_rs = conn.Execute(footer_sql)
Response.Write "<footer_dept>" & footer_rs("footer_dept") & "</footer_dept>"
Response.Write "<footer_dept_url>" & footer_rs("footer_dept_url") & "</footer_dept_url>"

'--------------頁尾維護單位-end---------------

response.write "<qStr>?site=2&amp;mp=" & request.querystring("mp") & "</qStr>"
%>

<!--#include file="x1Menus.inc" -->
<%	
	sql = "SELECT * FROM counter where mp='" & request("mp") & "'"
	Set rs = conn.Execute(sql)
	If Not rs.EOF Then
		count = rs("counts") + 1
		sql = "UPDATE counter SET counts = counts + 1  WHERE mp='" & request("mp") & "'"
	Else
		count = 1
		sql="INSERT INTO counter (mp, counts) VALUES ('" & request("mp") & "','1')"
	End If
	Response.Write "<Counter>" & count & "</Counter>"
	Set rs = conn.Execute(sql)
	
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	sql = "SELECT max(xpostDate) FROM CuDtGeneric AS htx JOIN CatTreeNode AS n ON n.CtUnitID=htx.iCtUnit" _
		& " AND n.CtRootID=" & xRootID
	set rs = conn.execute(sql)
	'新增ctTreeRoot 20070608
	
%>
<lastupdate><% =year(rs(0)) & "/" & month(rs(0)) &"/" & day(rs(0)) %></lastupdate>
<etoday><%=date()%></etoday>
<today></today>
</hpMain>
