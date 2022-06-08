
<% response.expires = 0 
HTProgPrefix="ppAct" %>
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

%>
<% Session("ActXID")="D" %>


<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>科學博物館科學教育解說中小學教師研習會</title>
<style type="text/css">
<!--

.txt       { font-family: "新細明體", "細明體"; font-size: 11pt; color: gray; line-height: 175%; letter-spacing: 1pt;left-spacing:20pt;}
A          { color:#ffffff; text-decoration: none; font-size: 10pt;}
A:hover    { color:#6600FF; text-decoration: none; background-color: #F2DEFE;}

-->
</style>


</head>

<body bgcolor="#B8C6AE">

<div align="center">
  <center>
  <table border="0" cellpadding="10" cellspacing="1" width="750" height="482" bgcolor="#858617">
    <tr>
      <td bgcolor="#FFFFFF" valign="top" align="center" height="460">
        <div align="center">
          <center>
          <table border="0" cellpadding="0" cellspacing="1" bgcolor="#858617" width="720" height="62">
            <tr>
              <td width="619" height="60" bgcolor="#FFFFFF" valign="middle" align="center">
                <div align="center">
                  <center>
                  <table border="0" cellpadding="0" cellspacing="0" width="720">
                    <tr>
                      <td colspan="7" width="718" align="right" valign="top"><embed width="718" height="62" src="eduworkshop/images/title.swf"></td>
                    </tr>
                    <tr>
                      <td width="264" bgcolor="#E8A438" height="20">　</td>
                      <td width="88" bgcolor="#858617" height="20" valign="middle" align="center">&nbsp;                                       <a href="eduworkshop/index.htm">活動說明</a></td>   
                      <td width="89" bgcolor="#858617" height="20" valign="middle" align="center"><a href="eduworkshop/schedule.htm">時程表</a></td>
                      <td width="89" bgcolor="#858617" height="20" valign="middle" align="center"><a href="enroll.htm">我要報名</a></td>
                      <td width="89" bgcolor="#858617" height="20" valign="middle" align="center"><a href="enroll.htm">查詢/取消</a></td>
                      <td width="89" bgcolor="#858617" height="20" valign="middle" align="center"><a href="http://www.nmns.edu.tw/">回首頁</a></td>
                    </tr>
                  </table>
                  </center>
                </div>
              </td>
            </tr>
          </table>
          </center>
        </div>
        　
        <table border="0" cellpadding="0" cellspacing="1" bgcolor="#9F9C93" width="720" height="354">
          <tr>
            <td  valign="top" align="center" bgcolor="#BAB8B2" height="298">
              <table border="0" cellpadding="0" cellspacing="0" width="718" height="330">
                <tr>
                  <td width="470" bgcolor="#E7D17C" valign="middle" align="left" class=txt height="42">
                  　.::<b>活動說明::.</b></td>
                  <td bgcolor="#E7D17C" valign="middle" align="right" class=txt height="42">
                  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="45" height="30">
	              <param name=movie value="eduworkshop/images/white.swf">
    	          <param name=quality value=high>
        	      <param name="wmode" value="transparent">
            	  <embed src="eduworkshop/images/white.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="30" height="30" wmode="transparent">
	          </embed> 
	          </object>
                  </td>
                </tr>
                <tr>
                  <td width="716"  bgcolor="#E7D17C" valign="top" align="center" colspan="2" height="288">
                    <table border="0" cellpadding="20" cellspacing="1" width="700" bgcolor="#E8A438" height="262">
                      <tr>
                        <td width="100%" valign="top" align="left" bgcolor="#F4F4E8" class=txt height="220">　ertyretyre
                     <table width="86%" border="0" cellspacing="1" cellpadding="6" align="center" bgcolor="#336699">
                       <tr bgcolor="#D0F79D" align="center"> 
                         <td class="txt2" width="28%" bgcolor="#F2C864">活動名稱</td>
                          <td class="txt2" bgcolor="#F2C864" >梯次日期（<font color=red>請點選欲參加之梯次日期進行報名</a>）</td>
                        </tr>
                                                   
                          <%
			sql = "SELECT htx.*, pe.status AS myStatus, a.actName, a.actDesc, a.actTarget, a.acthrs, a.actfee" _
			& ", (SELECT count(*) FROM paEnroll AS e WHERE e.pasID=htx.pasID) AS erCount" _
			& ", (SELECT sum(ppCount) FROM paEnroll AS p WHERE p.pasID=htx.pasID) AS xppCount" _
			& " FROM paSession AS htx JOIN ppAct AS a ON a.actID=htx.actID" _
			& " LEFT JOIN paEnroll AS pe on pe.pasID=htx.pasID AND pe.psnID=" & pkStr(Session("wuxPID"),"") _
			& " WHERE actCat='M1'" _
			& " ORDER BY htx.actID, htx.bDate"
	
			'	response.write sql 
			'	response.end
	
			set RS = conn.execute(sql)
			newActID = ""
			cnt = 0
			
			while not RS.eof
			if RS("actID") <> newActID then
				if newActID <> "" then
					response.write "</table></td></tr>"
				end if
				count = 0
				newActID = RS("actID")
				cnt=cnt+1
			%>
			<TR bgcolor="#E7E7E7">

			 <td width=40% class="txt2" align="center" title="<%=RS("actDesc")%>" bgcolor="#FCF0D3"><%=RS("actName")%></td>
			 <td class="txt2" valign=top bgcolor="#FCF0D3">
			<table border=0  width=100%>
			<%
			end if
			count = count +1
			xStyle = "cursor:hand; font-size: 12px; text-decoration: underline; "
			
			if RS("myStatus") <> "" then		xStyle= xStyle & " background:pink;"
			if RS("aStatus")="A"  AND (isNull(RS("xppCount")) OR RS("xppCount")<RS("pLimit")) then
			
			%>
			<TR >
				<td  class="txt2" width=90% align=center><!--第<%=count%>梯:   -->
				<span onMouseover="VBS: me.style.background = '#dae8c6'" 
      				onMouseOut="VBS: me.style.background = '#E7E7E7'"			
				onClick="VBS: xDoEnroll <%=RS("paSID")%>" style="<%=xStyle%>"> 
				<!--<%=RS("dtnote")%><font size=-1 color=red>[<%=RS("erCount")%>]-->
				<%=RS("dtnote")%></span>
				<font size=-1>(已報名人數: <%=RS("xppcount")%>人)<font color=#336699> </font>   
				<a href=../ppActivity/psEnrollList.asp?paSID=<%=RS("paSID")%>>錄取名單</a></font>
				
				
				</td>
			</TR>
			<%
				else
			%>
			<TR>
			<td  class="txt2" width=90%><A--第<%=count%>梯:   -->
				<span> <font size=-1>
				<%=RS("dtnote")%></span><font size=-1 color=red>[已額滿]　<a href=../ppActivity/psEnrollList.asp?paSID=<%=RS("paSID")%>>錄取名單</a></font>
              </font>
			</td>
			</TR>
			<%
			end if
			RS.moveNext
			wend
			response.write "</table></td></tr>"
			'	response.end
			%>
			
			</table><!-- first talbe -->                        
                        
                        
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td bgcolor="#858617" height="20" class=txt valign="middle" align="center">
              <font color="#FFFFFF">國立自然科學博物館．The       
              National Museum of Nature Science</font>      
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>



<%
Function chnWeekday(xDay)
	if isDate(xDay) then
'   Get_weekday_num = WeekDay(GWeekDaystr)
   	  Select case weekDay(xDay)
    	case 1 : chnWeekday = "日"
    	case 2 : chnWeekday = "一"
    	case 3 : chnWeekday = "二"
    	case 4 : chnWeekday = "三"
    	case 5 : chnWeekday = "四"
    	case 6 : chnWeekday = "五"
    	case 7 : chnWeekday = "六"
   	  End Select 
   	else
   		chnWeekday = ""
   	end if
End Function
%>
<script language=vbs>
sub xReturn (xValue)
	msgBox xValue
	document.location.reload
end sub

sub xDoEnroll(xValue)
	window.navigate "enrollAct3.asp?paSID=" & xValue

end sub

</script>





