<% Response.Expires = 0
HTProgCap="志工回報單"
HTProgFunc="新增"
HTProgCode="ap01"
HTProgPrefix="paAct" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<html>
<head><STYLE type="text/css">
<!--
BODY {
scrollbar-face-color:#E1F7FD;
scrollbar-highlight-color:#C8FBFA;
scrollbar-3dlight-color:#BDAEF7;
scrollbar-darkshadow-color:#A4A6F0;
scrollbar-shadow-color:#C4F1FB;
scrollbar-arrow-color:#BCC9F5;
scrollbar-track-color:#F1F0FD;
}
-->
</STYLE>

<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="css.css" type="text/css">

</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellspacing="0" cellpadding="0">
  
  <tr> 
    
    <td width="400" valign="top"> 
      <table width="400" border="0" cellspacing="8" cellpadding="0" align="center">
        <tr>
          <td valign="top">志工回報單</td>
        </tr>
        <tr> 
          <td valign="top"><br>
         
        <table>
        <tr>
           <td>請輸入志工編號:<input type="text" ><br>
           年度:<select>
           <option>93</option>
           <option>94</option>
           </select><br>
           <input type="submit" value="確定"></td>
        </tr> 
        </table>
        <hr>
       <TABLE summary='活動' border="2">
  <TR>
    <Td nowrap>志工編號</Td>
    <Td nowrap>場次</Td>
    <Td  nowrap>時數</Td>
    <Td  nowrap>活動地點</Td>
  </tr>  
  <tr>  
    <td>12345</td>
    <TD ><A href="v7_add.htm" >國家森林志工步道認養活動</A></TD>
    <TD >3 hr</TD>
    <TD >南澳工作站 </TD>
  </TR>
   <tr>  
    <td>12345</td>
    <TD ><A href="v7_add.htm" >國家森林志工步道認養活動2</A></TD>
    <TD >4 hr</TD>
    <TD >南澳工作站2 </TD>
  </TR>
  
</TABLE><br>
          </td>
        </tr>
      </table>
    <p>　</p></td>
  
  </tr>
 
</table>
</body>
</html>