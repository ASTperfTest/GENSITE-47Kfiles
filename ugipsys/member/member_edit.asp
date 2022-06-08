<%@ CodePage = 65001 %>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
	account = request("account")

	'取得帳號資料
	sql = "select * from Member where account = N'" & account & "'"
	set rs = conn.execute(sql)

	passwd = trim(rs("passwd"))
	realname = trim(rs("realname"))
	homeaddr = trim(rs("homeaddr"))
	phone = trim(rs("phone"))
	mobile = trim(rs("mobile"))
	email = trim(rs("email"))
	mcode = trim(rs("mcode"))
%>
<html>
<head>
<title>財政部全球資訊網 - Ministry of Finance,R.O.C.</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="../inc/setstyle.css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%"  border="0" cellspacing="0" cellpadding="0" summary="主要資訊表格">
<form method="post" action="member_edit_act.asp">
  <input type="hidden" name="account" value="<%= account %>">
  <tr> 
    <td valign="top">
      <table width="95%"  border="0" align="center" cellpadding="3" cellspacing="0">
        <tr>
          <td valign="top" class="NewsTitle">會員資料修改</td>
        </tr>
        <tr>
          <td align="right" valign="top"><hr align="left" width="100%" size="1" color="#cccccc"></td>
        </tr>
        <tr>
          <td valign="top">
            <table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#D8E7EA" summary="會員欄位">
              <tr>
                <td width="18%" align="right" bgcolor="#EBF2F3" class="whitetablebg"><strong>＊帳　號：</strong></td>
                <td bgcolor="#FFFFFF"><%=account%></td>
              </tr>
              <tr>
                <td align="right" bgcolor="#EBF2F3" class="whitetablebg"><strong>＊密　碼：</strong></td>
                <td bgcolor="#FFFFFF"><input name="passwd1" type="password" class="InputBox" size="30" value="<%=passwd%>"></td>
              </tr>
              <tr>
                <td align="right" bgcolor="#EBF2F3" class="whitetablebg"><strong>＊密碼確認：</strong></td>
                <td bgcolor="#FFFFFF"><input name="passwd2" type="password" class="InputBox" size="30" value="<%=passwd%>"></td>
              </tr>
              <tr>
                <td align="right" bgcolor="#EBF2F3" class="whitetablebg"><strong>＊姓　名：</strong></td>
                <td bgcolor="#FFFFFF"><input name="realname" type="text" class="InputBox" value="<%=realname%>" size="30"></td>
              </tr>
              <tr>
                <td align="right" bgcolor="#EBF2F3" class="whitetablebg"><strong>住　址：</strong></td>
                <td bgcolor="#FFFFFF"><input name="homeaddr" type="text" class="InputBox" value="<%=homeaddr%>" size="50"></td>
              </tr>
              <tr>
                <td align="right" bgcolor="#EBF2F3" class="whitetablebg"><strong>電　話：</strong></td>
                <td bgcolor="#FFFFFF"><input name="phone" type="text" class="InputBox" value="<%=phone%>" size="30"></td>
              </tr>
              <tr>
                <td align="right" bgcolor="#EBF2F3" class="whitetablebg"><strong>手　機：</strong></td>
                <td bgcolor="#FFFFFF"><input name="mobile" type="text" class="InputBox" value="<%=mobile%>" size="30"></td>
              </tr>
              <tr>
                <td align="right" bgcolor="#EBF2F3" class="whitetablebg"><strong>＊電子信箱：</strong></td>
                <td bgcolor="#FFFFFF"><input name="email" type="text" class="InputBox" value="<%=email%>" size="30"></td>
              </tr>
              <tr>
                <td align="right" bgcolor="#EBF2F3" class="whitetablebg"><strong>＊身分群組：</strong></td>
                <td bgcolor="#FFFFFF">
                  <select class="InputBox" name="mcode">
<%
	sql2 = "select mcode,mvalue from CodeMain where codeMetaId=N'xvGroup' order by msortValue"
	set ts=conn.Execute(sql2)
	while not ts.eof
		if mcode = trim(ts("mcode")) then
%>
                    <option value="<%=trim(ts("mcode"))%>" selected><%=trim(ts("mvalue"))%></option>
<%
	else
%>
                    <option value="<%=trim(ts("mcode"))%>"><%=trim(ts("mvalue"))%></option>
<%
		end if
		ts.movenext
	wend
%>
                  </select>
                </td>
              </tr>
              <tr>
                <td bgcolor="#EBF2F3">&nbsp;</td>
                <td height="50" bgcolor="#FFFFFF">
                  <input name="submit" type="submit" id="submit" value="確定">
                  <input name="cancel" type="button" id="cancel" value="取消" onClick="history.back();">
                </td>
              </tr>
            </table>
          </td>
        </tr>
      <tr>
        <td align="left" valign="top">&nbsp;</td>
      </tr>
      <tr>
        <td align="right" valign="top"><hr align="left" width="100%" size="1" color="#cccccc"></td>
      </tr>
      <tr>
        <td align="right" valign="top">&nbsp;</td>
      </tr>
    </table>
    </td>
  </tr>
</form>
</table>
</body>
</html>
