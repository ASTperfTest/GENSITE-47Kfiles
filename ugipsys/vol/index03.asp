<% Response.Expires = 0
HTProgCap="志工資料"
HTProgFunc="志工資料"
HTProgCode="ap03"
HTProgPrefix="mSession" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<!--#include virtual = "/inc/hyftdGIP.inc" -->

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
<table cellspacing="0" cellpadding="0" >
 
  <tr> 
    
    <td width="400" valign="top"> 
      <table width="400" border="0" cellspacing="8" cellpadding="0" align="center">
        <tr>
          <td valign="top">志工資料</td>
        </tr>
        <tr> 
          <td valign="top"><br>
          <p><A href="a3_add.asp" >新增志工資料</A></p>
<%
  sql="select * from volitient order by unit"
  sqlcnt="select count(*) from volitient"

  set rs=conn.Execute(sql)
  set rscnt=conn.Execute(sqlcnt)

  Dim NowPage, Page
  NowPage = Request("Page")
  If NowPage = "" Then
        NowPage = 1
  End If

  Page = Int(rscnt(0) / 10)
  If rscnt(0) mod 10 <> 0 Then
        Page = Page + 1
  End If
  if page=0 then
   page=1
  end if   
%>
<table>
<form>
<tr>
  <td>
      共有<%Response.Write Page%>頁，目前在第<%Response.Write NowPage%>頁　跳到： 
        <select name="select" onchange='Page_Onchange(document.select.value)'>
          <%
        Dim PageNo
        PageNo = 1
        While PageNo<=Page 
                Response.Write "<option "
                If Int(PageNo) = Int(NowPage) Then
                        Response.Write "selected"
                End If
                Response.Write " value=" & PageNo & " >" & PageNo & "</option>"
                PageNo = PageNo + 1
        WEnd
                
                
       %>
        </select>
        頁【每頁10筆】 
   </td>
  </tr>
</form>
</table>
             <TABLE summary='活動' border="2">
  <TR>
    <Td nowrap>志工編號</Td>
    <Td  nowrap>姓名</Td>
    <Td  nowrap>隸屬林管處</Td>
  </tr>  
 <%
    Dim Count
    Count=1
    While Not rs.Eof
     
      If Count>(NowPage-1) * 10 and Count <= NowPage*10 Then
    %>  
  <tr>  
    <TD ><A href="a3_fix.asp?id=<% =trim(rs("id")) %>" ><% =trim(rs("id")) %></A></TD>
    <TD ><% =trim(rs("name")) %></TD>
    <TD ><% 
sql2="select * from codemain where codemetaid='forest_unit' and mcode='" & trim(rs("unit"))  & "'"
   set rs2=conn.Execute(sql2)
  if not rs2.eof then 
   response.write trim(rs2("mvalue")) 
  end if%></TD>
  </TR>
<%        
      End If
      rs.MoveNext
      Count = Count + 1
    WEnd    
%>    
</TABLE>
            
          </td>
        </tr>
      </table>
    <p>　</p></td>
  
  </tr>
  
</table>
</body>
</html>
<script language="JavaScript">

<!--

function Page_Onchange(value1)
 {
            location.replace("index03.asp?Page="+value1)
 }

//-->

</script>         