<%@ CodePage = 65001 %>
<?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<!--#Include virtual = "/inc/MSclient.inc" -->
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
if not footer_rs.eof then
Response.Write "<footer_dept>" & footer_rs("footer_dept") & "</footer_dept>"
Response.Write "<footer_dept_url>" & footer_rs("footer_dept_url") & "</footer_dept_url>"
end if
'--------------頁尾維護單位-end---------------

response.write "<qStr>?site=2&amp;mp=" & request.querystring("mp") & "</qStr>"
%>

<!--#include file="x1Menus.inc" -->
<%	
	sql = "SELECT * FROM counter where mp='" & request("mp") & "'"
	Set rs = conn.Execute(sql)
	If Not rs.EOF Then
		count = rs("counts") + 1
		'sql = "UPDATE counter SET counts = counts + 1  WHERE mp='" & request("mp") & "'"
	Else
		count = 1
		'sql="INSERT INTO counter (mp, counts) VALUES ('" & request("mp") & "','1')"
	End If
	'把count移到viewcounter.aspx 
	'Set rs = conn.Execute(sql)
	sql = " SELECT a.ctrootid , a.viewcount AS allview,b.ViewCount AS thisYearView, "
	sql = sql & " c.viewcount AS thisMonthView FROM( "
	sql = sql & "	Select ctRootId, sum(ViewCount) as ViewCount "
	sql = sql & "		 from CounterForSubjectByDate WHERE CtRootId = '" & request("mp") & "' "
	sql = sql & "		 GROUP BY  ctRootId) AS a "
    sql = sql & "     LEFT JOIN (   "
	sql = sql & "	Select ctRootId, sum(ViewCount) as ViewCount "
	sql = sql & "		from CounterForSubjectByDate WHERE CtRootId = '" & request("mp") & "' "
	sql = sql & "		and YEAR(ymd) = YEAR(GETDATE()) GROUP BY  ctRootId ) b ON a.ctRootId = b.ctRootId "
    sql = sql & "   LEFT JOIN ( "
    sql = sql & "   Select ctRootId, sum(ViewCount) as ViewCount "
	sql = sql & "		from CounterForSubjectByDate WHERE CtRootId = '" & request("mp") & "' "
	sql = sql & "		and MONTH(ymd) = MONTH(GETDATE()) GROUP BY  ctRootId )c ON a.ctRootId = c.ctRootId " 
	Set rs = conn.Execute(sql)
	If Not rs.EOF Then
		if Not IsNull(rs("allview"))  then
			countAll = CLng(rs("allview"))
		end If
		if Not IsNull(rs("thisYearView")) then
			countThisYear =CLng(rs("thisYearView"))
		end If
		if Not IsNull(rs("thisMonthView"))then
			countThisMoth = CLng(rs("thisMonthView"))
		Else
			countThisMoth = 1
		end If
	Else
		countAll = 1
		countThisYear = 1
		countThisMoth = 1
	End IF
	Response.Write "<CounterAll>" & countAll & "</CounterAll>"
	Response.Write "<CounterThisYear>" & countThisYear & "</CounterThisYear>"
	Response.Write "<CounterThisMonth>" & countThisMoth & "</CounterThisMonth>"
	Response.Write "<Counter>" & count & "</Counter>"
	
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
