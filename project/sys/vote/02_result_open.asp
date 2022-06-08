<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	'HTProgCap = "功能管理"
	'HTProgFunc = "問卷調查"
	HTProgCode = "GW1_vote01"
	'HTProgPrefix = "mSession"
	response.expires = 0
%>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<%	
	
	subjectid = request("subjectid")
	questionid = request("questionid")
	answerid = request("answerid")
	
	if subjectid = "" or questionid = "" or answerid = "" then
		response.write "<script language='JavaScript'>alert('empty');history.go(-1);</script>"
		response.end
	end if
	
	
	sql = "" & _
		" select m011_subject, m011_bdate, m011_edate, m011_haveprize " & _
		" from m011 where m011_subjectid = " & replace(subjectid, "'", "''")
	set rs = conn.execute(sql)
	if rs.eof then
		response.write "<script language='JavaScript'>alert('error');history.go(-1);</script>"
		response.end
	end if
	
	
	set ts = conn.execute("select count(*) from m014 where m014_subjectid = " & replace(subjectid, "'", "''"))
	if ts(0) = 0 then
		response.write "<script language='JavaScript'>alert('目前尚無任何答題資料！');history.go(-1);</script>"
		response.end
	end if
	
	
	sql = "select m012_title from m012 where m012_subjectid = " & subjectid & " and m012_questionid = " & questionid
	set rs2 = conn.execute(sql)
	if rs2.eof then
		response.write "<script language='JavaScript'>alert('error2');history.go(-1);</script>"
		response.end
	end if
	
	
	sql = "select m013_title from m013 where m013_subjectid = " & subjectid & " and m013_questionid = " & questionid & " and m013_answerid = " & answerid
	set rs3 = conn.execute(sql)
	if rs3.eof then
		response.write "<script language='JavaScript'>alert('error3');history.go(-1);</script>"
		response.end
	end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF">
<p>問卷調查</p>
<table border="1" cellspacing="0" cellpadding="2" width="100%" bordercolorlight="#999900" bordercolordark="#FFFFFF">
  <tr> 
    <td colspan="2">統計項目： 
      <input type="button" value="答案累計統計" onClick="javascript:location.replace('02_result.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
<%
	'if rs("m011_haveprize") = "1" then
%>
      <input type="button" value="調查對象分析" onClick="javascript:location.replace('02_group.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
      <input type="button" value="調查對象明細" onClick="javascript:location.replace('02_namelist.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
<%
	'end if
%>
      <input type="button" value="意見回覆" onClick="javascript:location.replace('02_feedback_note.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
    </td>
  </tr>

  <!--
  <tr>
    <td>
      <input type="button" value="資料匯出" onClick="javascript:location.replace('02_result_exl.asp?subjectid=<%= subjectid %>');">
      <font color="#993300">[請按資料匯出之後,會出現另存新檔,副檔名請存成xls檔即可]</font>
    </td>
  </tr>
  -->
  
  <tr><td colspan="2">起訖時間：<%= DateValue(rs("m011_bdate")) %> ~ <%= DateValue(rs("m011_edate")) %></td></tr>
  <tr><td colspan="2">調查主題：<%= trim(rs("m011_subject")) %></td></tr>
  <tr><td colspan="2">題目：<%= trim(rs2("m012_title")) %></td></tr>
  <tr><td colspan="2">答案：<%= trim(rs3("m013_title")) %></td></tr>

<%
	sql = "select m016_content from m016 where m016_subjectid = " & subjectid & " and m016_questionid = " & questionid & " and m016_answerid = " & answerid
	set rs = conn.execute(sql)
	i = 0
	while not rs.eof
		i = i + 1
%>
  <tr>
    <td align="center" width="5%" nowrap><%= i %></td>
    <td><%= trim(rs("m016_content")) %></td>
  </tr>
<%
		rs.movenext
	wend
%>

</table>
<p><a href="javascript:history.go(-1);">Previous Page</a></p>
</body>
</html>
