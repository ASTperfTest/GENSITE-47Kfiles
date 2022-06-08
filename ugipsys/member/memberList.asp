<%@ CodePage = 65001 %>
<%
Response.Expires = 0 
HTProgCode = "HT002"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc"-->
<%
'nowPage=Request.QueryString("nowPage")
'現在頁數
nowPage=1
if len(trim(request("nowPage"))) > 0 then
	nowPage=cint(trim(request("nowPage")))
end if

'每頁筆數
mpp=20
if len(trim(request("mpp"))) > 0 then
	mpp=cint(trim(request("mpp")))
end if
if mpp <> 10 and mpp <> 20 and mpp <> 30 and mpp <> 50 and mpp <> 100 and mpp <> 200 then
	mpp=20
end if

'總筆數
m_num=0
sql="select count(*) from Member"
set rs=conn.execute(sql)
m_num=cint(trim(rs(0)))

'總頁數
p_num = fix((m_num-1)/mpp) + 1

if nowPage > p_num then
	nowPage=p_num
end if
if nowPage < 1 then
	nowPage=1
end if

%>
<html><head>
<title>會員列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"; cache=no-caches;>
<link rel="stylesheet" href="../inc/setstyle.css">
</head>
<body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<form name="thisform" method="get" action="memberList.asp">
<table border="0" width="100%" cellspacing="1" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2">
      會員列表
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>  
  <tr>
    <td width="100%" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2" align=center>
      <table border="0" width="500" cellspacing="0" cellpadding="0">
        <tr>
          <td align=center>
            <font size="2" color="rgb(63,142,186)">
              第
              <font size="2" color="#FF0000"><%= nowPage %> / <%= p_num %></font>
              頁
              |
              共
              <font size="2" color="#FF0000"><%= m_num %></font>
              筆
              |
              跳至第
              <select name="nowPage" size="1" style="color:#FF0000" onChange="submit();">
<%
	for iPage = 1 to p_num
%>
                <option value="<%= iPage %>"
<%
		if iPage=nowPage then
%>
                SELECTED
<%
		end if
%>
                ><%= iPage %></option>
<%
	next
%>
              </select>
              頁
<%
	if nowPage > 1 then
%>
              |
              <a href="memberList.asp?nowPage=<%= (nowPage-1) %>&mpp=<%= mpp %>">上一頁</a>
<%
	end if
	if nowPage < p_num then
%>
              |
              <a href="memberList.asp?nowPage=<%= (nowPage+1) %>&mpp=<%= mpp %>">下一頁</a>
<%
	end if
%>
              |
              每頁筆數：
              <select name="mpp" size="1" style="color:#FF0000" onChange="submit();">
                <option value="10"<%if mpp=10 then%> selected<%end if%>>10</option>
                <option value="20"<%if mpp=20 then%> selected<%end if%>>20</option>
                <option value="30"<%if mpp=30 then%> selected<%end if%>>30</option>
                <option value="50"<%if mpp=50 then%> selected<%end if%>>50</option>
                <option value="100"<%if mpp=100 then%> selected<%end if%>>100</option>
                <option value="200"<%if mpp=200 then%> selected<%end if%>>200</option>
              </select>
            </font>
          </td>
          <td align=right>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <CENTER>
        <TABLE width=90% cellspacing="1" cellpadding="3" class="bluetable">
          <tr align=left>
            <td width=5% align=center class=lightbluetable>&nbsp;</td>
            <td align=center class="lightbluetable">帳號</td>
            <td align=center class="lightbluetable">姓名</td>
            <td align=center class="lightbluetable">EMAIL</td>
            <td align=center class="lightbluetable">註冊日期</td>
            <td align=center class="lightbluetable">修改</td>
            <td align=center class="lightbluetable">刪除</td>
          </tr>
<%
	num_start = (nowPage-1) * mpp + 1
	num_end = nowPage * mpp

	sql="select account, realname, email, createtime from Member order by createtime desc"
	set rs=conn.execute(sql)
	i=0
	while not rs.EOF
		i=i+1
		if i >= num_start and i <= num_end then
%>
          <tr>
            <TD align=center class="whitetablebg"><%= i %></TD>
            <TD class="whitetablebg"><%= trim(rs("account")) %></TD>
            <TD class="whitetablebg"><%= trim(rs("realname")) %></TD>
            <TD class="whitetablebg"><%= trim(rs("email")) %></TD>
            <TD class="whitetablebg"><%= trim(rs("createtime")) %></TD>
            <TD align=center class="whitetablebg"><a href="member_edit.asp?account=<%= trim(rs("account")) %>">修改</a></TD>
            <TD align=center class="whitetablebg"><a href="member_edit_act.asp?account=<%= trim(rs("account")) %>&submit=delete" onClick="return confirm('確定刪除？');">刪除</a></TD>
          </tr>
<%
		end if
		rs.movenext
	wend
%>
        </TABLE>
      </CENTER>
      <p align="right">
      <img src="../images/management/titlehr.gif" width="570" height="1" vspace="3">
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2" align="center" class="English">
	  <!--#include virtual = "/inc/Footer.inc" --> 
    </td>
  </tr>
</table>
</form>
</body>
</html>
