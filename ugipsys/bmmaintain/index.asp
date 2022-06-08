<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
HTProgCap="意見信箱"
HTProgFunc="意見信箱維護"
HTProgCode="BM010"
HTProgPrefix="mSession"
response.expires = 0 
%>
<!-- #include virtual = "/inc/server.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	
	set rs2 = conn.execute("select * from mailbox")
	if not rs2.eof then	email = trim(rs2("email"))	
	status = trim(Request("status"))
	If status = "" Then	status = "0"
	
	Select Case status
		Case "0":
			sql2 = "select count(*) from Prosecute where mailord = 1"
		Case "1"
			sql2 = "select count(*) from Prosecute where mailord = 1 and sendflag = '1'"
		Case "2"
			sql2 = "select count(*) from Prosecute where mailord = 1 and (sendflag is null or sendflag='')"
	End Select
	
	set ts = conn.execute(sql2)
	Recordcount = ts(0)
	
	Select Case status
		Case "0":
			sql2 = "select Id ,Name, Email, DATE, replydate, isnull(Reply,'') Reply, docid,sendflag from Prosecute where mailord = 1 order by Id desc"
		Case "1"
			sql2 = "select Id ,Name, Email, DATE, replydate, isnull(Reply,'') Reply, docid,sendflag from Prosecute where mailord = 1 and sendflag = '1' order by Id desc"
		Case "2"
			sql2 = "select Id ,Name, Email, DATE, replydate, isnull(Reply,'') Reply, docid,sendflag from Prosecute where mailord = 1 and  (sendflag is null or sendflag='') order by Id desc"
	End Select	
	set rs = conn.Execute(sql2)
	
	perpagerecords = 10
	NowPage = Request("Page")
	totalPage = Int(Recordcount / perpagerecords)
	If NowPage = "" Then	NowPage = 1	
	If Recordcount mod perpagerecords <> 0 Then	totalPage = totalPage + 1
%>
<html>
<head>
<title><%=session("mySiteName")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
</head>

<body>
<table border="0" cellspacing="1" cellpadding="0" width="80%">
  <tr>
    <td><%=session("mySiteName")%>意見信箱</td>
  </tr>
  <tr>
    <td><hr noshade size="1" color="#000080"></td>
  </tr>
</table>
<table cellpadding="2" cellspacing="0" width="80%">
  <form action="mailbox_fix.asp" method="post">
  <input type="hidden" name="mailord" value="1">
  <tr>
    <td class=whitetablebg>
      意見信箱：<input type="text" name="email" value="<%= email %>" size="15" maxlength="50">
      <input type="submit" name="Submit62" value="修改">
    </td>
  </tr>
  </form>
  <form action="index.asp" method="post" name="c">
      <tr> 
         <td class=whitetablebg>回覆狀態：
          <select onchange="javascript:location.replace('index.asp?status=' + this.value);">
          <option value="0"
<%		
		If status = "0" Then	Response.Write " selected" end if
%>
		>全部</option>
		 <option value="1"
<%		
		If status = "1" Then	Response.Write " selected" end if
%>
		>已回信</option>
		 <option value="2"
<%		
		If status = "2" Then	Response.Write " selected" end if
%>
		>未回信</option>
      </select>        
               
         </td>
      </tr>
   </form>  
  <tr>
    <td><hr noshade size="1" color="#000080"></td>
  </tr>
</table>
<table cellpadding="2" cellspacing="0" border="0" width="80%">
  <form>
  <tr>
    <td colspan="11">
      <font size="2" color="rgb(63,142,186)">
      共有<%= totalPage %>頁，目前在第<%= NowPage %>頁　跳到：
      <select onchange="javascript:location.replace('index.asp?Page=' + this.value + '&status=' + <%=status%>);">
<%
	PageNo = 1
	While PageNo <= totalPage
		Response.Write "<option value=" & PageNo
		If Int(PageNo) = Int(NowPage) Then	Response.Write " selected"
		Response.Write ">" & PageNo & "</option>"
		PageNo = PageNo + 1
	WEnd
%>
      </select>
      頁【每頁<%= perpagerecords %>筆】 
<% if NowPage > 1 then 
%> 
<a href="index.asp?Page=<%= NowPage - 1 %>&status=<%=status%>"> 上一頁  </a> 
<% end if %> 
&nbsp;&nbsp;
<% if (totalPage - NowPage) > 0 then 	
%>  
  <a href="index.asp?Page=<%= NowPage + 1 %>&status=<%=status%>">  下一頁 </a>
<% end if 
%>  
      </font>
    </td>
  </tr>
  </form>
</table>
<table cellspacing="1" cellpadding="8" class=bg bgcolor=navy width="80%">
  <tr align=left bgcolor=white>
    <td align="center" class=lightbluetable nowrap width="5%">編號</td>
    <td align="right" class=lightbluetable nowrap width="8%">來信日期</td>
    <td class=lightbluetable nowrap>姓名</td>
  <!--<td align="center" class=lightbluetable nowrap width="10%">觀看內容</td> -->
    <td align="center" class=lightbluetable nowrap width="10%">Email</td>  
    <td align="right" class=lightbluetable nowrap width="8%">回信日期</td>
    <td align="center" class=lightbluetable nowrap>輸入／觀看回信內容</td>
    <td align="center" class=lightbluetable nowrap width="5%">狀態</td>
   </tr>
<%
	ii = 1
	While Not rs.Eof
		If ii>(NowPage-1) * perpagerecords and ii <= NowPage*perpagerecords Then
%>
   <tr bgcolor=white>
    <td class=whitetablebg align=center><%= ii %></td>
    <td class=whitetablebg align=right nowrap><%= datevalue(rs("DATE")) %></td>
    <td class=whitetablebg nowrap><%= trim(rs("Name")) %></td>
   <!-- <td class=whitetablebg align=center nowrap><a href="detail.asp?id=<%= rs("id") %>&mailord=1&page=<%= NowPage %>">觀看內容</a></td> -->
    <td class=whitetablebg align=center><%= rs("Email") %></td> 
    <td class=whitetablebg align=right nowrap><% if rs("replydate") <> "" then response.write datevalue(rs("ReplyDate")) else response.write "&nbsp;" end if %></td>
    <td class=whitetablebg align="center" nowrap>
<%
        	If trim(rs("Reply")) = "" Then
        		response.write "尚未回信,<a href='input.asp?id=" & rs("id") & "&mailord=1&page=" & NowPage & "'>輸入回信內容</a>"
		else
			response.write "<a href='input.asp?id=" & rs("id") & "&mailord=1&page=" & NowPage & "'>觀看回信內容</a>"
		end if
%>
    </td>
    <td class=whitetablebg align=center><% if trim(rs("sendflag")) = "1" then response.write "寄出" else response.write "&nbsp;" end if  %></td>
  </tr>
<%
		end if
		rs.movenext
		ii = ii + 1
	wend
%>
</table>
<p><a href="javascript:history.go(-1);"><b>回上頁</b></a></p>
</body>
</html>
