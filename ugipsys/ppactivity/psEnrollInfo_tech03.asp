<% Response.Expires = 0
HTProgCap="活動管理"
HTProgFunc="梯次列表"
HTProgCode="PA001"
HTProgPrefix="paSession" %>
<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
Dim mpKey, dpKey
Dim RSmaster, RSlist
	sqlCom = "SELECT * FROM ppAct WHERE actID=" & pkStr(request.queryString("actID"),"")
	Set RSmaster = Conn.execute(sqlcom)
	mpKey = ""
	mpKey = mpKey & "&actID=" & RSmaster("actID")
	if mpKey<>"" then  mpKey = mid(mpKey,2)
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="95%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
	       <a href="paActList.asp?x=1">回活動清單</a>
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
</table>
<CENTER><TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD class="lightbluetable" align="right">活動名稱：</TD>
<TD class="whitetablebg"><input type="hidden" name="htx_actID">
<input name="htx_actName" size="40" readonly="true" class="rdonly">
</TD>
<TR><TD class="lightbluetable" align="right">對象：</TD>
<TD class="whitetablebg"><input name="htx_actTarget" size="50" readonly="true" class="rdonly">
</TD>
</TR>

</TABLE>
</CENTER>
<%
function qqRS(fldName)
	xValue = RSmaster(fldname)
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
	end if
end function %>

<script language=vbs>
	document.all("htx_actID").value= "<%=qqRS("actID")%>"
	document.all("htx_actName").value= "<%=qqRS("actName")%>"
	document.all("htx_actTarget").value= "<%=qqRS("actTarget")%>"
</script>
<%
	fSql = "SELECT htx.*, xref1.mValue AS xrStatus " _
		& " FROM (paSession AS htx LEFT JOIN CodeMain AS xref1 ON xref1.mCode = htx.aStatus AND xref1.codeMetaID='classStatus')"
	fSql = fSql & " WHERE htx.actID=" & pkStr(RSmaster("actID"),"")
	fSql = fSql & " ORDER BY bDate"
	set RSlist = conn.execute(fSql)
%>
	<INPUT TYPE=button VALUE="新增" onClick="window.navigate('paSessionAdd.asp?<%=mpKey%>')">
<CENTER>
 <TABLE width="95%" cellspacing="1" cellpadding="0" class="bg">
 <tr align="left">
	<td class=lightbluetable>梯次代碼</td>
	<td class=lightbluetable>起始日期</td>
	<td class=lightbluetable>日期時間</td>
	<td class=lightbluetable>正取/備取</td>
	<td class=lightbluetable>狀態</td>
 </tr>
<%
	while not RSlist.eof
		dpKey = ""
		dpKey = dpKey & "&paSID=" & RSlist("paSID")
		if dpKey<>"" then  dpKey = mid(dpKey,2)
%>
	<TD class=whitetablebg><font size=2>
<%=RSlist("paSID")%>
</font></td>
	<TD class=whitetablebg><font size=2>
	<A href="paSessionEdit.asp?<%=dpKey%>">
<%=s7Date(RSlist("bDate"))%>
</A>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSlist("dtNote")%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSlist("pLimit")%>/<%=RSlist("pBackup")%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSlist("xrStatus")%>
</font></td>
    </tr>
    <%
         RSlist.moveNext
     wend
   %>
    </TABLE>
    </CENTER>
       <!-- 程式結束 ---------------------------------->  
     </form>  
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table> 
</body>
</html>                                 

<script Language=VBScript>

</script>
