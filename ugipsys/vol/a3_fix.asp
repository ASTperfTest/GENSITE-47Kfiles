<% Response.Expires = 0
HTProgCap="志工資料"
HTProgFunc="志工資料"
HTProgCode="ap03"
HTProgPrefix="mSession" %>
<% response.expires = 0 %>
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
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<!--#include virtual = "/inc/hyftdGIP.inc" -->
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellspacing="0" cellpadding="0" width="720" height="496">
 <form action="a3_fix_act.asp" method="post">
  <tr> 
    <td width="400" valign="top"> 
<%
   id=replace(request("id"),"'","''")
   sql="select * from volitient where id='" & id & "'"
   set rs=conn.Execute(sql)
   if not rs.eof then 
%>
      <table width="400" border="0" cellspacing="8" cellpadding="0" align="center">
        <tr>
          <td valign="top">基本資料</td>
        </tr>
        <tr> 
          <td valign="top"><br>
            <table>
        <tr>
           <td>志工編號:<% =trim(rs("id")) %>
           <input type="hidden" name="id" value="<% =id %>">
           </td>
        </tr> 
        <tr>
           <td>姓名:<input type="text" name="name" value="<% =trim(rs("name")) %>"></td>
        </tr> 
         <tr>
           <td>密碼:<input type="password" name="password"  value="<% =trim(rs("passwd")) %>"></td>
        </tr> 
         <tr>
           <td>生日:<input type="text" size="4" name="yy" >年<input type="text" size="4" name="mm">月<input type="text" size="4" name="dd">日</td>
        </tr>   
        <tr>
           <td>教育程度:
          <select name="grade">
            <option <% if trim(rs("grade"))="大學" then response.write "selected" end if %>>大學</option>
            <option <% if trim(rs("grade"))="高中" then response.write "selected" end if %>>高中</option> 
          </select></td>
        </tr>   
        <tr>
           <td>隸屬林管處:
<select name="unit">
<% 
   sql2="select * from codemain where codemetaid='forest_unit'"
   set rs2=conn.Execute(sql2)
   do while not rs2.eof 
%>
<option  value="<% =trim(rs2("mcode")) %>" <% if trim(rs("unit"))=trim(rs2("mcode")) then response.write "selected" end if %>><% =trim(rs2("mvalue")) %></option>
<%
   rs2.Movenext
   loop
%>

</select></td>
        </tr>   
          <tr>
           <td><font color="red">*請提供基本資料需要欄位</font></td>
        </tr>   
        </table>
        <center><input type="submit" value="修改">&nbsp;&nbsp;<input type="button" value="刪除"></center>
          </td>
        </tr>
       
      </table>
<%
   end if
%>
    <p>　</p></td>
  
  </tr>
   </form>
</table>
</body>
</html>