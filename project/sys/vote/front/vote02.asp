<%
	Response.Expires = 0
	'HTProgCap = "￥\¯aoT2z"
	'HTProgFunc = "°Y‥÷?O?d"
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
  subjectID=request("subjectID")

  
  sql="select m011_subject from m011 where m011_subjectid=" & subjectID
  set rs=conn.execute(sql)
  if not rs.EOF then
     Subject=rs("m011_subject")
  end if
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
<noscript>
  您的瀏覽器不支援script
</noscript>
<link href="template.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="center" valign="top">
          
            <table width="90%" border="0" cellpadding="0" cellspacing="1" style="margin-top:20px;">
              <tr> 
                <td height="20"> <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="16" valign="middle" bgcolor="A4D265"><img src="../images/Service/service05.gif" width="16" height="20" alt="裝飾圖"></td>
                      <td width="80" valign="middle" bgcolor="A4D265" class="ContentTitle">問卷調查</td>
                      <td valign="middle" bgcolor="8DC73F" class="ContentTitle">&nbsp;</td>
                      <td width="16" align="right" bgcolor="8DC73F"><img src="../images/Service/service07.gif" width="16" height="20" alt="裝飾圖"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td height="40" bgcolor="eeeeee" class="SubTitle">歡迎使用問卷調查</td>
              </tr>
              <tr> 
                <td height="80" align="center">
                   <table width="100%" border="0" cellspacing="1" cellpadding="2">
<%
sql2="select * from m012 where m012_subjectid=" & subjectID & " order by m012_questionid"
set rs2=conn.execute(sql2)
if not rs2.eof then
%>
                    <tr bgcolor="666666"> 
                      <td height="25" colspan="5" class="IntroTitle"><font color="ffffff"><%=Subject%></font></td>
                    </tr>
<%
i=1
while not rs2.EOF

'sum_sql="select * from m013 where m013_subjectid=" & subjectID & " and m013_questionid=" & rs2("m012_questionid") & " order by m013_answerid"
%>                    
                   
                    <tr bgcolor="dddddd"> 
                      <td align="center" class="IntroContent"><img src="../images/Service/Q.gif" width="25" height="25" alt="問題圖示"></td>
                      <td colspan="4" class="IntroContent"><%=i%>. <%=trim(rs2("m012_title"))%>
                      </td>
                    </tr>
                    <tr bgcolor="eeeeee"> 
                      <td align="center" valign="top" class="IntroContent"><img src="../images/Service/A.gif" width="25" height="25" alt="答案圖示"></td>
                      <td colspan="4" valign="top" bgcolor="dddddd" class="IntroContent"> 
                        <table width="100%" border="0" cellspacing="1" cellpadding="0">
                        <%
                          sqls="select * from m013 where m013_subjectid=" & subjectID & " and m013_questionid=" & rs2("m012_questionid")
                          set srs=conn.execute(sqls)             
                          sum=0
                          while not srs.EOF       
                            sum=sum+srs("m013_no") 
                          srs.MoveNext
                          wend    
                          if sum=0 then
                             sum=1
                          end if

                          sql3="select * from m013 where m013_subjectid=" & subjectID & " and m013_questionid=" & rs2("m012_questionid") & " order by m013_answerid"
                          set rs3=conn.execute(sql3)             
                          while not rs3.EOF    
     
                          num=rs3("m013_no")*100.00/sum
                          'num=FormatNumber(rs3("m013_no")*100/sum , 2, -1)
%>                             
                          <tr bgcolor="eeeeee"> 
                            <td class="IntroContent"><%=trim(rs3("m013_title"))%></td>
                            <td class="IntroContent"><img src="../images/Service/bar_chart.gif" width="<%=Int(num*100.00)/100.00%>%" height="12" hspace="3" alt="統計結果圖"><%=rs3("m013_no")%>人(<%=Int(num*100.00)/100.00%>%)</td>
                          </tr>
                        <%
                           rs3.MoveNext
                           wend
%>         
                        </table></td>
                    </tr>
<%
i=i+1
rs2.MoveNext
wend
end if
%>     
                  </table> </td>
              </tr>
              <tr> 
                <td height="50" align="center" bgcolor="eeeeee"> <a href="vote.asp"><img src="../images/Service/back.gif" width="60" height="20" border="0" alt="回列表"></a></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>

</body>
</html>
