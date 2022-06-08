<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="GC1AP9"
HTProgPrefix="newkpi" %>
<!--#INCLUDE FILE="kpiListParam.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
response.charset = "buf-8"
response.codepage = "65001"
dim nodeArray(999)
dim cols, nodeTree
Set cols=Server.CreateObject("Scripting.Dictionary")
Set nodeTree=Server.CreateObject("Scripting.Dictionary")

sub outCode(xstr)
	xmlStr = xmlStr & xStr
end sub

sub outCode1(xstr)
	excelStr = excelStr & xStr
end sub

function ncTabChar(n)
	if cTabchar = "" then
		ncTabChar = ""
	else
		ncTabChar = String(n,cTabchar)
	end if
end function

sub traverseTree (parent)
	SqlCom = "SELECT CtNodeID, CatName, CtNodeKind, DataLevel, dCondition, unit.CtUnitID FROM CatTreeNode as node " _
		& " left join ctunit as unit on unit.ctunitid=node.ctunitid " _
		& " WHERE node.CtRootID = "& PkStr(CtRoot,"") &" AND dataParent=" & parent & " and node.inUse='Y'  " _
		& " and unit.redirectURL is null " _
		& " Order by node.CatShowOrder "
  
	set RSt = conn.execute(SqlCom)

	while not RSt.eof
		nodeTree.Add nodeCount, ncTabchar(Cint(RSt("DataLevel"))-1)&RSt("CatName")	
		Set nodeArray(nodeCount) = server.CreateObject("Scripting.Dictionary") 
		if RSt("CtUnitID")<>"" then
			if request("htx_func")="dept" then
				sql = "select  iDept, count(*) AS xCount, min(dept.abbrName) AS colName  from cudtgeneric as htx " _
					& " LEFT JOIN infoUser AS u ON htx.iEditor=u.userID  " _
					& " LEFT JOIN dept ON dept.deptID = htx.iDept  " _
					& " where htx.ictunit=" & RSt("CtUnitID") _
					& " AND htx.dEditDate BETWEEN '" & replace(htx_IDateS, "'", "''") & "' and '" & replace(htx_IDateE, "'", "''") & "' "
				if RSt("dCondition")<>"" then
					sql = sql & " AND " & RSt("dCondition")
				end if
				sql = sql & " GROUP BY htx.iDept " _
					& " ORDER BY htx.iDept DESC"
				set rsx = conn.execute(sql)
				'response.write sql
				while not rsx.eof
					if rsx("colName")<>"" then
						
						nodeArray(nodeCount).Add trim(rsx("colName")), trim(rsx("xCount"))
						'response.write nodeCount & RSt("CatName") & " == " & trim(rsx("colName")) & rsx("xCount") & "<br>"
						'response.write trim(rsx("colName")) & "=" & nodeArray(nodeCount).item(trim(rsx("colName"))) & "<br>"
						'dic.Add rs("colName"), rs("xCount")
						If not cols.Exists(trim(rsx("colName"))) Then
						'	cols.Add trim(rsx("iDept")), trim(rsx("colName"))
							cols.Add trim(rsx("colName")), trim(rsx("colName"))
						'	response.write rsx("iDept") & " == " & rsx("colName") & "<br>"
						end if
					end if
					rsx.moveNext
				wend
			else
				sql = "select  count(*) AS xCount, htx.iEditor, min(u.username) AS colName  from cudtgeneric as htx " _
					& " LEFT JOIN infoUser AS u ON htx.iEditor=u.userID  " _
					& " where htx.ictunit=" & RSt("CtUnitID") _
					& " AND htx.dEditDate BETWEEN '" & replace(htx_IDateS, "'", "''") & "' and '" & replace(htx_IDateE, "'", "''") & "' "
				if RSt("dCondition")<>"" then
					sql = sql & " AND " & RSt("dCondition")
				end if
				sql = sql & " GROUP BY htx.iEditor " _
					& " ORDER BY htx.iEditor DESC"
				set rsx = conn.execute(sql)		
				while not rsx.eof
					if rsx("colName")<>"" then
						nodeArray(nodeCount).Add trim(rsx("colName")), trim(rsx("xCount"))
						If not cols.Exists(trim(rsx("colName"))) Then
							cols.Add trim(rsx("colName")), trim(rsx("colName"))
						end if
					end if
					rsx.moveNext
				wend				
			end if
		else
			'nodeArray(nodeCount)
		end if

		nodeCount = nodeCount + 1
		if RSt("CtNodeKind") = "C" then   traverseTree RSt("CtNodeID")
		RSt.moveNext
	wend
end sub	

sub genDataView()
	outCode "<tr><td></td>"
	'response.write "<table border=1 width=3000>"
	'response.write "<tr><td></td>"

	for each k2 in cols
		outCode "<th>" & k2 & "</th>"
		'response.write "<td>" & k2 & "</td>"
	next	
	outCode "</tr><tr>"
	'response.write "</tr><tr>"
	for each k1 in nodeTree
		outCode "<td nowrap style=color:#616161;font-size:13px;background-color:#E4E3E3;><b>" & nodeTree.item(k1) & "</b></td>"
		'response.write "<td>" & nodeTree.item(k1) & "</td>"
		for each k2 in cols
			If nodeArray(k1).Exists(k2) Then	
				outCode "<td align=center>" & nodeArray(k1).item(k2) & "</td>"
				'response.write "<td align=right>" & nodeArray(k1).item(k2) & "</td>"
			else
				outCode "<td align=center>0</td>"
				'response.write "<td align=right>0</td>"
			end if
		next	
		outCode "</tr>"
		'response.write "</tr>"
	next
	'outCode "</table>"
	'response.write "</table>"
	set clos = nothing
	set nodeTree = nothing	
end sub
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 11">
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author></o:Author>
  <o:LastAuthor></o:LastAuthor>
  <o:Created></o:Created>
  <o:LastSaved></o:LastSaved>
  <o:Version>11.5606</o:Version>
 </o:DocumentProperties>
</xml><![endif]-->
<style>
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
@page
	{margin:1.0in .75in 1.0in .75in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;}
tr
	{mso-height-source:auto;
	mso-ruby-visibility:none;}
col
	{mso-width-source:auto;
	mso-ruby-visibility:none;}
br
	{mso-data-placement:same-cell;}
.style0
	{mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	white-space:nowrap;
	mso-rotate:0;
	mso-background-source:auto;
	mso-pattern:auto;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:新細明體, serif;
	mso-font-charset:136;
	border:none;
	mso-protection:locked visible;
	mso-style-name:一般;
	mso-style-id:0;}
td
	{mso-style-parent:style0;
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:新細明體, serif;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border:none;
	mso-background-source:auto;
	mso-pattern:auto;
	mso-protection:locked visible;
	white-space:nowrap;
	mso-rotate:0;}
.xl24
	{mso-style-parent:style0;
	mso-number-format:"Short Date";}
ruby
	{ruby-align:left;}
rt
	{color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:新細明體, serif;
	mso-font-charset:136;
	mso-char-type:none;
	display:none;}
-->
</style>
<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook>
  <x:ExcelWorksheets>
   <x:ExcelWorksheet>
    <x:Name>Sheet1</x:Name>
    <x:WorksheetOptions>
     <x:DefaultRowHeight>330</x:DefaultRowHeight>
     <x:Selected/>
     <x:Panes>
      <x:Pane>
       <x:Number>3</x:Number>
       <x:ActiveRow>2</x:ActiveRow>
      </x:Pane>
     </x:Panes>
     <x:ProtectContents>False</x:ProtectContents>
     <x:ProtectObjects>False</x:ProtectObjects>
     <x:ProtectScenarios>False</x:ProtectScenarios>
    </x:WorksheetOptions>
   </x:ExcelWorksheet>
  </x:ExcelWorksheets>
  <x:WindowHeight>12495</x:WindowHeight>
  <x:WindowWidth>18195</x:WindowWidth>
  <x:WindowTopX>480</x:WindowTopX>
  <x:WindowTopY>105</x:WindowTopY>
  <x:ProtectStructure>False</x:ProtectStructure>
  <x:ProtectWindows>False</x:ProtectWindows>
 </x:ExcelWorkbook>
</xml><![endif]-->
</head>
<%

	dim CtRoot, orgRoot, dtLevel, xmlStr, excelStr
	xmlStr=""
	excelStr=""
	cTabchar = "　　"
	CtRoot = request("htx_CtRootId")
	htx_IDateS = request("htx_IDateS")
	htx_IDateE = request("htx_IDateE")	
	if not isdate(htx_IDateS) or not isdate(htx_IDateE) then
		htx_IDateS = year(date) & "/" & month(date) & "/" & day(date) & " "
		htx_IDateE = htx_IDateS
	end if	
	dim nodeCount : nodeCount = 0		
	dim xcatCount(10)
	traverseTree 0
	genDataView
	if trim(request("submit")) = "匯出" then
		date_str = year(date) & "-" & month(date) & "-" & day(date) & "-" & hour(time) & "-" & minute(time) & "-" & second(time)
		response.AddHeader "Content-disposition","attachment; filename=KPI_" & date_str & ".xls"
%>
		<table><%=xmlStr%></table>
	</body>
</html>
<%
		response.end
	end if
%>
<html>
<head>
<title><%= HTProgCap %></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="http://www.mof.gov.tw/xslgip/mof94/CSS/lpcp.css" rel="stylesheet" type="text/css"/>
<script language="JavaScript" src="ts_picker.js"></script>
</head>
<body leftmargin="0" rightmargin="0" topmargin="10" marginwidth="0" marginheight="0">
 <form name="form1" action="newkpi.asp" method="post">
<!--<form method="POST" name="reg" action="">-->
<INPUT TYPE=hidden name=submitTask value="">
<object data="/inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="60%">
<TR><TD class="eTableLable" colspan=2 align="center">【<%=HTProgCap%>】</TR>
<TR><TD class="eTableLable" align="right">日期範圍：</TD>
<TD class="eTableContent">
<input name="pcShowhtx_IDateS" size="8" readonly onclick="VBS: popCalendar 'htx_IDateS','htx_IDateE'"> ～ 
<input name="htx_IDateS" type=hidden><input name="htx_IDateE" type=hidden>
<input name="pcShowhtx_IDateE" size="8" readonly onclick="VBS: popCalendar 'htx_IDateE',''">
<!--
<input type="text" name="startdate" size="10" value="<%= startdate %>" readonly onclick="javascript:show_calendar('document.form1.startdate', document.form1.startdate.value);">
至
<input type="text" name="enddate" size="10" value="<%= enddate %>" readonly onclick="javascript:show_calendar('document.form1.enddate', document.form1.enddate.value);">
-->
</TD>
</TR>
<tr><td class="eTableLable" align="right">分類樹：</td>
    <td class="eTableContent">
      <select name="htx_CtRootId">
	  <!--<option value="">不限定</option>-->
<%
	sqlR = " select CtRootId, CtRootName from CatTreeRoot " & _
		" where inUse = 'Y' order by CtRootId "
	set rsR = conn.execute(sqlR)
	while not rsR.eof
%>
		<option value="<%= trim(rsR("CtRootId")) %>"<% if CtRoot = trim(rsR("CtRootId")) then %> selected<% end if %>><%= rsR("CtRootName") %></option>
<%
		rsR.movenext
	wend
%>	  
      </select>
</td></tr>

<tr><td class="eTableLable" align="right">評比角度：</td>
    <td class="eTableContent">
      <select name="htx_func">
	  <!--<option value="">不限定</option>-->
		
		<option value="user">人員</option>
		<option value="dept">單位</option>
      </select>
</td></tr>
<tr><td colspan=2 align=center>
	<input type="submit" name="submit" value="查詢">
	
	<input type="button" name="print" value="友善列印" onClick="window.print()" ></td></tr>
</TABLE>
<hr noshade size="1" color="#000080">
</CENTER>
</form>     
<table class="list">
<%
response.write xmlStr
%>
</table>
<script language=vbs>
Dim CanTarget
Dim followCanTarget

sub popCalendar(dateName,followName)        
 	CanTarget=dateName
 	followCanTarget=followName
	xdate = document.all(CanTarget).value
	if not isDate(xdate) then	xdate = date()
	document.all.calendar1.setDate xdate
	
 	If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 	Else        
       ex=window.event.clientX
       ey=document.body.scrolltop+window.event.clientY+10
       if ex>520 then ex=520
       document.all.calendar1.style.pixelleft=ex-80
       document.all.calendar1.style.pixeltop=ey
       document.all.calendar1.style.visibility=""
 	End If              
end sub     

sub calendar1_onscriptletevent(n,o) 
    	document.all("CalendarTarget").value=n         
    	document.all.calendar1.style.visibility="hidden"    
    if n <> "Cancle" then
        document.all(CanTarget).value=document.all.CalendarTarget.value
        document.all("pcShow"&CanTarget).value=document.all.CalendarTarget.value
        if followCanTarget<>"" then
        	document.all(followCanTarget).value=document.all.CalendarTarget.value
        	document.all("pcShow"&followCanTarget).value=document.all.CalendarTarget.value
        end if
	end if
end sub   
</script>