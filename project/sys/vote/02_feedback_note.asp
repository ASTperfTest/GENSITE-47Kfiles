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
		response.write "<script language='JavaScript'>history.go(-1);</script>"
		response.end
	end if
	
	
	sql = "" & _
		" select m011_subject, m011_bdate, m011_edate, m011_haveprize from m011 where " & _
		" m011_subjectid = " & subjectid
	set rs = conn.execute(sql)
	if rs.EOF then
		response.write "<script language='JavaScript'>history.go(-1);</script>"
		response.end
	end if
	
	
	set ts = conn.execute("select count(*) from m014 where m014_subjectid = " & subjectid)
	if ts(0) = 0 then
		response.write "<script language='JavaScript'>alert('目前尚無任何答題資料！');history.go(-1);</script>"
		response.end
	end if
	base_answer_no = ts(0)
	
	
	sql = "" & _
		" select count(*) from m014 where m014_subjectid = " & subjectid & " and " & _
		" isNull(m014_reply, '') not like '' "
	set ts = conn.Execute(sql)
	
	totalrecord = ts(0)
	if totalrecord > 0 then              'ts代表筆數
		totalpage = totalrecord \ 10
		if (totalrecord mod 10) <> 0 then
			totalpage = totalpage + 1
		end if
	else
		totalpage = 1
	end if
	
	if request("page") = empty then
		page = 1
	else
		page = request("page")
	end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function Page_Onchange(form) {
	location.replace("02_feedback_note.asp?page=" + form.select.value + "&subjectid=<%= subjectid %>&functype=<%= functype %>");
}
-->
</script>
</head>

<body bgcolor="#FFFFFF">
<p>問卷調查分析</p>
<p>僅列出有留意見的資料</p>
<table border="1" cellspacing="0" cellpadding="2" width="100%" bordercolordark="#FFFFFF">
  <tr> 
    <td colspan="3">統計項目： 
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
  
  <tr>
    <td colspan="3">
      <input type="button" value="資料匯出" onClick="javascript:location.replace('02_feedback_note_exl.asp?subjectid=<%= subjectid %>');">
      <font color="#993300">[請按資料匯出之後,會出現另存新檔,副檔名請存成xls檔即可]</font>
    </td>
  </tr>  
  <tr><td colspan="3">起訖時間：<%= DateValue(rs("m011_bdate")) %> ~ <%= DateValue(rs("m011_edate")) %></td></tr>
  <tr><td colspan="3">調查主題：<%= trim(rs("m011_subject")) %></tr>  

  <form>
  <tr> 
    <td colspan="3">
      共有 <% =totalpage %> 頁，目前在第 <% =page %> 頁，跳到： 
      <select name="select" onChange='Page_Onchange(this.form)'>
<%
	n = 1
	While n <= totalpage 
		Response.Write "<option value='" & n & "'"
		If Int(n) = Int(Page) Then 
			Response.Write " selected"
		End If
		Response.Write ">" & n & "</option>"
		n = n + 1
	WEnd
%>
      </select>
      頁【每頁10筆】
    </td>
  </tr>
  </form>
  
  <tr> 
    <td bgcolor="#EFEFEF" width="10%">姓名</td>
    <td bgcolor="#EFEFEF" width="20%">email</td>
    <td bgcolor="#EFEFEF">意見內容</td>
  </tr>
  
<%
	sql = "" & _
		" select m014_name, m014_email, m014_reply from m014 " &_
		" where m014_subjectid = " & subjectid & _
		" and isNull(m014_reply, '') not like '' " & _
		" order by m014_polldate desc "
	set rs = conn.Execute(sql)
	while not rs.EOF
%>
  <tr> 
    <td bgcolor="#EFEFEF" width="10%"><%= trim(rs("m014_name")) %>&nbsp;</td>
    <td bgcolor="#EFEFEF" width="20%"><a href="mailto:<%= trim(rs("m014_email")) %>"><%= trim(rs("m014_email")) %></a>&nbsp;</td>
    <td bgcolor="#EFEFEF"><%= trim(rs("m014_reply")) %></td>
  </tr>
<%
		rs.MoveNext
	wend
%>

</table>
<p><a href="javascript:history.go(-1);">Previous Page</a></p>
</body>
</html>
