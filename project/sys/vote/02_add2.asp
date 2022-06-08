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
	questno = request("questno")
	if subjectid = "" or questno = "" then
		response.write "<body onload='javascript:history.go(-1);'>"
		response.end
	end if
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
    <td align="center" valign="top">
      <p><br>逐題顯示問卷輸入問題題目與答案</p>
      <table border="1" cellspacing="0" cellpadding="2" width="100%" bordercolordark="#FFFFFF">
      <form method="post" name="form1" action="02_add2_act.asp">
      <input type="hidden" name="functype" value="<%= functype %>">
        <tr> 
          <td colspan="2"><a name=note></a>
            <p><font color="#993300">使用說明：</font></p>
            <ul>
              <li><font color="#993300">請逐一輸入每一題目的問題，以及答案選項。</li>
               <li>按“重新調查此份問卷”，刪除原作答紀錄重新調查，並儲存修改資料，回列表。</li>
               <li>按“上一步”，回上一個修改頁面</li>
               <li>按“修改資料”，將進行單一題目修改功能，不刪除原作答紀錄，在儲存資料後，回列表。</font></li></ul>
          </td>
        </tr>

<%
	sql = "select m012_title, m012_type from m012 where m012_subjectid = " & subjectid & _
		" and m012_questionid <= " & questno & " order by m012_questionid "
	set rs = conn.execute(sql)	

	for i = 1 to questno		
		title = ""
		m012_type = "1"
		if not rs.EOF then
			title = replace(trim(rs("m012_title")), chr(34), "&quot;")
			m012_type = rs("m012_type")
			rs.MoveNext
		end if		
%>        
        <tr> 
          <td rowspan="2">題目<% =i %>：</td>
          <td>
            本題答題方式： 
            <input type="radio" name="type<% =i %>" value="1"<% if m012_type = "1" then %> checked<% end if %>> 單選&nbsp;
            <input type="radio" name="type<% =i %>" value="2"<% if m012_type = "2" then %> checked<% end if %>>複選
          </td>
        </tr>
        <tr> 
          <td>
            <b>題目<% =i %>：</b><input type="text" name="title<% =i %>" value="<% =title %>"><br>
<%
		answerno = request("answerno" & i)

		sql = "" & _
			" select m013_title, m013_default from m013 where m013_subjectid = " & subjectid & _
			" and m013_questionid = " & i & " and m013_answerid <= " & answerno & " order by m013_answerid "
		set rs2 = conn.execute(sql)

		for j = 1 to answerno			
			title = ""
			m013_default = "N"
			if not rs2.EOF then
				title = replace(trim(rs2("m013_title")), chr(34), "&quot;")
				m013_default = rs2("m013_default")
				rs2.MoveNext
			end if
%>
            答案 <% =j %>：<input type="text" name="answer<% =i %>-<% =j %>" value="<% =title %>">&nbsp;&nbsp;
            設定為預設值 <input type="checkbox" name="default<% =i %>-<% =j %>" value="Y"<% if m013_default = "Y" then %> checked<% end if %>><br>
<%
		next
%>
          </td>
        </tr>
<%
	next
%>

        <tr> 
          <td colspan="2">
            <input type="submit" name="submit" value="確定">
            <input type="button" name="back" value="上一步" onClick="return history.go(-1);">
          </td>
        </tr>
      </form>
      </table>
      <p><a href="javascript:history.go(-1);">Previous Page</a></p>
    </td>
  </tr>
</table>
</body>
</html>
