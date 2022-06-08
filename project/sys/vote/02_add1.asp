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
	if subjectid = "" then
		response.write "<body onload='javascript:history.go(-1);'>"
		response.end
	end if
	
	sql = "select m011_questno, m011_jumpquestion from m011 where m011_subjectid = " & subjectid
	set rs = conn.execute(sql)
	if rs.EOF then
		response.write "<body onload='javascript:history.go(-1);'>"
		response.end
	end if
	questno = rs(0)
	jumpquestion = rs(1)
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#ffffff">
問卷調查管理 
<table width="100%" height="95%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td>&nbsp;</td>
    <td colspan="3" align="center" valign="top" > 
      <table border="1" cellspacing="0" cellpadding="2" width="100%" bordercolordark="#FFFFFF">

      <form method="post" name="form1" action="02_add2_jump.asp">
      <input type="hidden" name="subjectid" value="<% =subjectid %>">
      <input type="hidden" name="questno" value="<% =questno %>">
      <input type="hidden" name="jumpquestion" value="<% =jumpquestion %>">
        <tr> 
          <td colspan="2">
            <p><font color="#993300">使用說明：</font></p>
            <ul><li><font color="#993300">請先輸入每一題目的答案數。</font></li></ul>
          </td>
        </tr>
<%
	for i = 1 to questno
		sql = "select m012_answerno from m012 where m012_subjectid = " & subjectid & _
			" and m012_questionid = " & i
		set rs = conn.execute(sql)
		if rs.EOF then
			answerno = 0
		else
			answerno = rs(0)
		end if
%>
        <tr> 
          <td>題目<% =i %>：</td>
          <td>答案數：
            <select name="answerno<% =i %>">
<%
		for j = 1 to 120
			if j = answerno then
				response.write "<option value='" & j & "' selected>" & j & "</option>"
			else
				response.write "<option value='" & j & "'>" & j & "</option>"
			end if
		next
%>
            </select>
          </td>
        </tr>
<%
	next
%>
        <tr> 
          <td colspan="2">
            <input type="submit" name="submit" value="下一步(問卷設計)">
            <input type="button" name="back" value="回上一步" onClick="window.location='02_fix.asp?subjectid=<% =subjectid %>'">
          </td>
        </tr>
      </form>
      </table>
      <p>&nbsp;</p>
    </td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
