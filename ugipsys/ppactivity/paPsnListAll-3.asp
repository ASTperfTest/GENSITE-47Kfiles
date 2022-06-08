<% Response.Expires = 0
HTProgCap="活動管理"
HTProgFunc="報名學員清單"
HTProgCode="PA005"
HTProgPrefix="psEnroll" %>
<% response.expires = 0 %>
<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
       session("paSID") = request.queryString("paSID")
	sql = "SELECT htx.*, a.actName, a.actCat, a.actDesc, a.actTarget, xref2.mValue AS xrActCat" _
		& ", (SELECT count(*) FROM paEnroll AS e WHERE e.pasID=htx.pasID) AS erCount" _
		& " FROM paSession AS htx JOIN ppAct AS a ON a.actID=htx.actID" _
		& " LEFT JOIN CodeMain AS xref2 ON xref2.mCode = a.actCat AND xref2.codeMetaID='ppActCat'" _
		& " WHERE htx.paSID=" & session("paSID")
	set RS = conn.execute(sql)

%>
<html>

<head>
<title>植物與美術-0712</title>
<meta content="text/html; charset=big5" http-equiv="Content-Type">
<link rel="stylesheet" href="../inc/setstyle.css">
<style type="text/css">.txt {
	COLOR: black; FONT-FAMILY: "新細明體", "細明體"; FONT-SIZE: 11pt; LETTER-SPACING: 1pt; LINE-HEIGHT: 160%
}
.tt {
	COLOR: black; FONT-FAMILY: "新細明體", "細明體"; FONT-SIZE: 10pt; LETTER-SPACING: 1pt; LINE-HEIGHT: 140%
}
.txt2 {
	COLOR: #493fd1; FONT-FAMILY: "新細明體", "細明體"; FONT-SIZE: 10pt; LETTER-SPACING: 1pt; LINE-HEIGHT: 165%
}
.t1 {
	COLOR: #660099; FONT-FAMILY: "新細明體", "新細明體"; FONT-SIZE: 13pt; FONT-WEIGHT: bold; LETTER-SPACING: 2pt; LINE-HEIGHT: 180%
}
A {
	COLOR: black; TEXT-DECORATION: none
}
A:hover {
	BACKGROUND-COLOR: #c5ebfe; COLOR: blue; TEXT-DECORATION: none
}
</style>
<meta content="MSHTML 5.00.3314.2100" name="GENERATOR">
</head>

<body bgColor="#ffccdd" leftMargin="2" text="#000000" topMargin="0" marginwidth="2">

<table border="0" cellPadding="0" cellSpacing="0" width="770">
<TBODY>
  <tr>
    <td vAlign="top"><table border="0" cellPadding="0" cellSpacing="0" width="100%">
<TBODY>
      <tr>
        <td colSpan="2"><img height="12" src="images/icon2.gif" width="770"></td>
      </tr>
      <tr>
        <td width="270"><img height="78" src="images/icon.gif" width="270"></td>
        <td bgColor="#f5f5f5" width="526"><table border="0" cellPadding="0" cellSpacing="0"
        width="100%">
<TBODY>
          <tr>
            <td width="60"></td>
                <td class="txt2" width="463">▋國立自然科學博物館<br>
                  ▋植物公園九十一年暑期研習活動<br>
                  ▋[水果奶奶這一家] 錄取名單<br>
                </td>
          </tr>
</TBODY>
        </table>
        </td>
      </tr>
      <tr>
        <td colSpan="2"><table background="images/bg6.gif" border="0" cellPadding="10"
        cellSpacing="0" width="100%">
<TBODY>
          <tr>
            <td width="5"></td>
                <td width="745">&nbsp;
                  <table border="0" width="95%" bgcolor="#CCFFFF" align="center" height="131">
                    <tr valign=top>
                      <td width="11%" height="129"><font color="#FF0000">請注意：</font></td>
                      <td width="89%" class=txt height="129"> 
                        <ol>
                          <li>繳費期限：請於91年6月26日（三）16:00前完成繳費手續，逾時視為自動放棄，名額由備取者遞補。（將以電話另行通知備取者）</li>
                          <li>繳費地點及時間：植物公園研究教育中心，9:00-16:00，請攜帶身分證明文件（戶口名簿或健保卡）。 
                          </li>
                          <li>洽詢電話：04-23285328。</li>
                        </ol>
                        </td> 
    </tr>
  </table>
                  <table border="0" width="95%" cellspacing="1" cellpadding="0" align="center">
                    <tr> 
                      <td class="FormName" width="52%" height="30"> 
                        <div align="left"><font size=2>◎ <%=RS("actName")%>&nbsp;<%=RS("dtNote")%>&nbsp;【正取名單】</font></div>
                      </td>
                      <td class="FormName" width="47%" height="30"> <font size=2>【備取名單】 
                        　　　　　　　　　<a href="Javascript:window.history.back();">回前頁</a> 
                        </font> </td>
                      <td class="FormLink" valign="top" align=right width="1%" height="30">&nbsp; 
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="3"> 
                        <hr noshade size="1" color="#000080">
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="3" class="FormRtext"></td>
                    </tr>
                    <tr> 
                      <td colspan="3" valign=top> 
                        <CENTER>
                          <table border="0" cellspacing="0" cellpadding="1" width="100%" height="24">
                            <tr> 
                              <td width="52%" valign="top" height="60"> 
                                
                              <table width=84% cellspacing="0" cellpadding="2" class=bg border=1 bordercolor="#003399">
                                <tr align=left> 
                                    <td class=lightbluetable width="36%"> 
                                      <div align="center">姓名</div>
                                    </td>
                                    <td class=lightbluetable width="64%"> 
                                      <div align="center">身份證號</div>
                                    </td>
                                  </tr>
                                  <%                   
	fSql = "SELECT htx.paSID, htx.ckValue, htx.erDate, htx.status, htx.psnID AS xPsnID, p.*" _
		& " FROM (paEnroll AS htx LEFT JOIN paPsnInfo AS p ON p.psnID = htx.psnID)" _
		& " WHERE htx.status='Y' AND htx.pasID=" & session("paSID")
	set RSreg = conn.execute(fSql)

    while not RSreg.eof           
%>
                                  <tr> 
                                    <td class=whitetablebg align=center width="36%"><font size=2><%=RSreg("pName")%></font></td>
                                    <td class=whitetablebg align=center width="64%"><font size=2><%=RSreg("xpsnID")%></font></td>
                                  </tr>
                                  <%
         RSreg.moveNext
	wend
   %>
                                </table>
                              </td>
                              <td width="48%" valign="top" height="60"> 
                                
                              <table width=82% cellspacing="0" cellpadding="2" class=bg border=1 bordercolor="#003399">
                                <tr align=left> 
                                    <td class=lightbluetable width="41%" height="19"> 
                                      <div align="center">姓名</div>
                                    </td>
                                    <td class=lightbluetable width="59%" height="19"> 
                                      <div align="center">身份證號</div>
                                    </td>
                                  </tr>
                                  <%                   
	fSql = "SELECT htx.paSID, htx.ckValue, htx.erDate, htx.status, htx.psnID AS xPsnID, p.*" _
		& " FROM (paEnroll AS htx LEFT JOIN paPsnInfo AS p ON p.psnID = htx.psnID)" _
		& " WHERE htx.status='B' AND htx.pasID=" & session("paSID")
	set RSreg = conn.execute(fSql)

    while not RSreg.eof           
%>
                                  <tr> 
                                    <td class=whitetablebg align=center width="41%"><font size=2><%=RSreg("pName")%></font></td>
                                    <td class=whitetablebg align=center width="59%"><font size=2><%=RSreg("xpsnID")%></font></td>
                                  </tr>
                                  <%
         RSreg.moveNext
	wend
   %>
                                </table>
                              </td>
                            </tr>
                          </table>
                       
                  </table>
                </td>
          </tr>
</TBODY>
        </table>
        </td>
      </tr>
</TBODY>
    </table>
    </td>
  </tr>
  <tr>
    <td>
      <table background="images/bg6.gif" border="0" cellPadding="6" cellSpacing="0"
    width="100%">
        <TBODY> 
        <tr> 
          <td width="5"></td>
          <td>&nbsp;</td>
        </tr>
        </TBODY> 
      </table>
    </td>
  </tr>
</TBODY>
</table>
</body>
</html>
