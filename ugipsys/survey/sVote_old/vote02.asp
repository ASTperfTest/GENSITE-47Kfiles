<%@ CodePage = 65001 %>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbFunc.inc" -->
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
<title>title</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
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

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<font><a>歡迎使用問卷調查</a></font>
               
                   <table width="55%" border="0" cellspacing="0" cellpadding="0">
                   
                    <tr bgcolor="666666"> 
                      <td height="25" colspan="5" class="IntroTitle"><font><%=trim(rs("m011_subject"))%></font></td>
                    </tr>
<%
sql2="select * from m012 where m012_subjectid=" & subjectID & " order by m012_questionid"
set rs2=conn.execute(sql2)
if not rs2.eof then

i=1
while not rs2.EOF

'sum_sql="select * from m013 where m013_subjectid=" & subjectID & " and m013_questionid=" & rs2("m012_questionid") & " order by m013_answerid"
%>                    
                   
                    <tr bgcolor="dddddd"> 
                      <td align="center" class="IntroContent"><img src="images/Service/Q.gif" width="25" height="25" alt="問題圖示"></td>
                      <td colspan="4" class="IntroTitle"><p><%=i%>. <%=trim(rs2("m012_title"))%></p></td>
                    </tr>
                    <tr bgcolor="eeeeee"> 
                      <td align="center" valign="top" class="IntroContent"><img src="images/Service/A.gif" width="25" height="25" alt="答案圖示"></td>
                      <td colspan="4" valign="top" bgcolor="dddddd" class="IntroContent"> 
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
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
                            <td class="IntroContent" border="0" width="50%"><%=trim(rs3("m013_title"))%></td>
                            <td class="IntroContent" width="50%"><img src="svote/images/Service/bar_chart.gif" width="<%=Int(num*100.00)/100.00%>-1%" height="12" hspace="3" alt="統計結果圖"><%=rs3("m013_no")%>人(<%=Int(num*100.00)/100.00%>%)</td>
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
                  </table> 
  

</body>
</html>
