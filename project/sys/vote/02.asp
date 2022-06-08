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
	set ts = conn.execute("select count(*) from m011")

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
</head>

<body bgcolor="#ffffff">
問卷調查管理
<hr/>
       			<div align="right"> <input type="button" value="新增調查主題" onClick="window.location='02_add.asp'">	</div>

<font size="2"> 
      共有<font color="red"><b> <% =totalpage %></b></font> 頁，目前在第<font color="red"><b> <% =page %></b></font> 頁，跳到：
      <select onChange="javascript:location.replace('02.asp?page=' + this.value);">
<%
	n = 1
	while n <= totalpage
		response.write "<option value='" & n & "'"
		if int(n) = int(page) then
			response.write " selected"
		end if
		response.write ">" & n & "</option>"
		n = n + 1
	wend
%>
      </select>
      頁【每頁10筆】
</font>
      
<table width="100%" height="95%" border="0" cellpadding="0" cellspacing="0">
  <form>
 
 	
  <tr>
    <td>&nbsp;</td>
    <td colspan="3" valign="top" >
      <table border="1" cellspacing="0" cellpadding="2" width="100%" bordercolordark="#FFFFFF">

       	
       	<!--
        <tr>
          <td colspan="6">
            <input type="button" value="新增調查主題" onClick="window.location='02_add.asp'">
          </td>
          <td>
            &nbsp; <a href="./front/vote.asp" target="_blank">前端</a>
          </td>
        </tr>
        -->
        <tr>
          <td width="5%" align="center">編號</td>
          <td width="20%">問卷起訖日期</td>
          <td width="7%" align="center">上線(入口網)</td>
          <td width="7%" align="center">上線(加值網)</td>
          <td width="7%" align="center">答題</td>
          <td>調查主題</td>
          <td width="8%">觀看結果</td>
          <td width="7%" align="center">問卷預覽</td>
          <td width="15%">亂數抽獎</td>
        </tr>

<%
	sql = "" & _
		" select m011_subjectid, m011_subject, m011_bdate, m011_edate, " & _
		" isNull(m011_online, '0') m011_online, m011_jumpquestion, " & _
		" m011_haveprize, isNull(m011_pflag, '0') m011_pflag, " & _
		" isNull(m011_km_online, '0') m011_km_online " & _
		" from m011 order by m011_bdate desc "
	set rs = conn.execute(sql)
	i = 1
	kmweburl = session("kmUrl") & "coa/vote/previewvote.aspx"
	Function GetGuid() 
        Set TypeLib = CreateObject("Scriptlet.TypeLib") 
        GetGuid = Left(CStr(TypeLib.Guid), 38) 
        Set TypeLib = Nothing 
    End Function 
	while not rs.eof
		if i <= (page * 10) and i > (page - 1) * 10 then
%>
        <tr valign="Top" height="30">
          <td align="center"><%= i %></td>
          <td><% =rs("m011_bdate") %> 至 <% =rs("m011_edate") %></td>
          <td align="center"><% if rs("m011_online") = "1" then %>是<% else %>否<% end if %></td>
          <td align="center"><% if rs("m011_km_online") = "1" then %>是<% else %>否<% end if %></td>
          <td align="center"><% if rs("m011_jumpquestion") = "1" then %>跳題<% else %>一般<% end if %></td>
          <td><a href="02_fix.asp?subjectid=<% =rs("m011_subjectid") %>"><%= trim(rs("m011_subject")) %></a></td>
          <td><a href="02_result.asp?subjectid=<% =rs("m011_subjectid") %>">結果</a></td>
          <td><a href="<%= kmweburl %>?subjectid=<% =rs("m011_subjectid") %>&view=<%= GetGuid() %>" target="_blank">view</a></td>
<%
			if rs("m011_haveprize") = "1" then
				if rs("m011_pflag") = "0" then
%>
          <td>
            <select name="prizeno<%= rs("m011_subjectid") %>">
<%
					for j = 1 to 5
						response.write "              <option value='" & j & "'>" & j & "</option>" & vbCrLf
					next
%>
            </select>
            <input type="button" value="抽獎" onClick="javascript:location.href='02_random.asp?subjectid=<% =rs("m011_subjectid") %>&prizeno<%= rs("m011_subjectid") %>=' + form.prizeno<%= rs("m011_subjectid") %>.value">
          </td>
<%
				else
%>
          <td>
            <a href="02_random.asp?subjectid=<% =rs("m011_subjectid") %>">得獎名單</a>
          </td>
<%
				end if
			else
				response.write "<td>此題不抽獎</td>"
			end if
%>        
        </tr>
<%
		end if
		i = i + 1
		rs.movenext
	wend
%>
      </table>

      <table border="1" cellspacing="0" cellpadding="2" width="100%" bordercolordark="#FFFFFF">
        <tr>
          <td colspan="6">
            <p><font color="#993300">使用說明：</font></p>
            <ul>
              <li><font color="#993300">點選“新增調查主題”，可以增加調查項目。</font></li>
              <li><font color="#993300">點選“調查主題”，可以進行修改刪除。其中，可以修改問卷起訖日期，控制問卷是否上線接受調查，與控制問卷結束，僅供使用者觀看結果的時間。（若要求從前端顯示消除，必須刪除本問卷。）</font></li>
              <li><font color="#993300">點選“觀看結果”，可以檢視目前問卷統計結果（如果問卷不調查個人基本資料，則統計資料中無法顯示對象分析，並無法亂數抽獎。）</font></li>
            </ul>
          </td>
        </tr>
      </table>
    </td>
    <td>&nbsp;</td>
  </tr>
  </form>
</table>
</body>
</html>
