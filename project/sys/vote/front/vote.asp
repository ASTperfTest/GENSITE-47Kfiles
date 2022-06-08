<%
	Response.Expires = 0
	'HTProgCap = "功能管理"
	'HTProgFunc = "問卷調查"
	HTProgCode = "GW1_vote01"
	'HTProgPrefix = "mSession"
	response.expires = 0
%>
<!-- #include file = "../../../../inc/server.inc" -->
<!-- #include file = "../../../../inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<%
Function chgDate(year, month, day)
	dim newmonth, newday

	if CInt(month) < 10 then
		newmonth = "0" & month
	else
		newmonth = month
	end if

	if CInt(day) < 10 then
		newday = "0" & day
	else
		newday = day
	end if

	chgDate = year & "/" & newmonth & "/" & newday

End Function
%>
<html>
<head>
<title>問卷調查</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script LANGUAGE="JavaScript">
<!--
ie4 = ((navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >= 4 ));
NN4 = (navigator.appName == "Netscape" && parseInt(navigator.appVersion) >= 4)
z = "xyz";
function toggle(targetId) {
	if ( NN4 || ie4 ) {
		if ( z != "xyz" ) {
			if ( z != document.all(targetId) )
				z.style.display = "none";
		}
		target = document.all(targetId);
		if ( target.style.display == "none" ) {
			target.style.display = "";
			z = document.all(targetId);
		} else {
			target.style.display = "none";
			z = "xyz";
		}
	}
}
ie4 = ((navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >= 4 ));
NN4 = (navigator.appName == "Netscape" && parseInt(navigator.appVersion) >= 4)
z = "xyz";
//-->
</script>
<noscript>您的瀏覽器不支援script</noscript>
<link href="template.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">

  <tr>
    <td>
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>

          <td align="center" valign="top">

            <table width="90%" border="0" cellpadding="0" cellspacing="1" style="margin-top:20px;">
              <tr>
                <td height="20">
                  <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="16" valign="middle" bgcolor="A4D265"><img src="../images/Service/service05.gif" width="16" height="20" alt="裝飾圖"></td>
                      <td width="80" valign="middle" bgcolor="A4D265" class="ContentTitle">問卷調查</td>
                      <td valign="middle" bgcolor="8DC73F" class="ContentTitle">&nbsp;</td>
                      <td width="16" align="right" bgcolor="8DC73F"><img src="../images/Service/service07.gif" width="16" height="20" alt="裝飾圖"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr>
                <td height="30" bgcolor="eeeeee" class="SubTitle">歡迎使用問卷調查</td>
              </tr>
              <tr>
                <td height="80" align="center">
                  <table width="100%" border="0" cellspacing="1" cellpadding="2">
                    <tr bgcolor="666666">
                      <td height="20" align="center" class="IntroTitle"><font color="ffffff">編號</font></td>
                      <td height="20" class="IntroTitle"><font color="ffffff">主題</font></td>
                      <td height="20" class="IntroTitle"><font color="ffffff">起訖時間</font></td>
                      <td height="20" class="IntroTitle"><font color="ffffff">參與調查</font></td>
                      <td height="20" class="IntroTitle"><font color="ffffff">觀看結果</font></td>
                    </tr>
<%
	curdate = chgDate(year(now), month(now), day(now))

	set rs = conn.execute("select * from m011 where m011_online = '1' order by m011_subjectid desc")

	page = request(page)
	if page = "" then
		page = 1
	end if

	i = 1
	while not rs.eof
		if i <= (page * 10) and i > (page - 1) * 10 then
			if (i mod 2) = 0 then
				color = "eeeeee"
			else
				color = "dddddd"
			end if
		
			bdate = chgDate(year(rs("m011_bdate")), month(rs("m011_bdate")), day(rs("m011_bdate")))
			edate = chgDate(year(rs("m011_edate")), month(rs("m011_edate")), day(rs("m011_edate")))
%>
                    <tr bgcolor="<%= color %>">
                      <td align="center" class="IntroContent"><%= i %>.</td>
                      <td class="IntroContent"><%= trim(rs("m011_subject")) %> </td>
                      <td class="IntroContent"><%= bdate %> ~ <br><%= edate %></td>
<%
			if curdate >= bdate and curdate <=edate then
				if rs("m011_jumpquestion") = "0" then
%>
                      <td align="center"><a href="vote01.asp?subjectID=<%= rs("m011_subjectid") %>"><img src="../images/Service/vote.gif" width="25" height="25" hspace="5" border="0" align="absmiddle" alt="參與調查">參與調查</a></td>
                      <td align="center"><a href="vote02.asp?subjectID=<%= rs("m011_subjectid") %>"><img src="../images/Service/result.gif" width="25" height="25" hspace="5" border="0" align="absmiddle" alt="觀看結果">觀看結果</a></td>
<% 
				else 
%>
                      <td align="center"><a href="vote01_jump.asp?subjectID=<%= rs("m011_subjectid") %>"><img src="../images/Service/vote.gif" width="25" height="25" hspace="5" border="0" align="absmiddle" alt="參與調查">參與調查</a></td>
                      <td align="center"><a href="vote02.asp?subjectID=<%= rs("m011_subjectid") %>"><img src="../images/Service/result.gif" width="25" height="25" hspace="5" border="0" align="absmiddle" alt="觀看結果">觀看結果</a></td>
<%
				end if
			else
%>
		      <td align="center">&nbsp;</td>
            	      <td align="center"><a href="vote02.asp?subjectID=<%= rs("m011_subjectid") %>"><img src="../images/Service/result.gif" width="25" height="25" hspace="5" border="0" align="absmiddle" alt="觀看結果">觀看結果</a></td>
<%
			end if
%>
                    </tr>
<%
		end if
		i = i + 1
		rs.movenext
	wend

	set ts = conn.execute("select count(*) from m011 where m011_online = '1'")
	if ts(0) > 0 then              'ts代表筆數
		totalpage = ts(0) \ 10
		if (ts(0) mod 10) <> 0 then
			totalpage = totalpage + 1
		end if
	else
		totalpage = 1
	end if
%>
              
                    <tr bgcolor="666666">
                      <td height="20" align="center" class="IntroTitle" colspan="5">&nbsp;</td>
                    </tr>

                  </table>
                </td>
              </tr>

            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
