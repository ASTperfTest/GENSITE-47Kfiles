<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "Pn50M06" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
if request("task") = "查詢" then
     Sqlcom = "Select * From Ugrp where 1=1"
     for each xItem in request.form
         if request(xItem) <> "" then
              if left(xItem,2)="fg" then
                    Sqlcom = Sqlcom & "AND " & mid(xItem,3) & " LIKE N'%" & request(xItem) & "%'"
              end if
         end if
     next
     Sqlcom=Sqlcom & " Order by ugrpId"
     session("ApQuery")=Sqlcom     
else
     Sqlcom = session("ApQuery")
end if

%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>代碼權限群組查詢清單</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<table border="0" width="95%" cellspacing="1" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2"><%=Title%></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>  
  <tr>
    <td class="Formtext" width="60%">【代碼權限群組查詢清單】</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right">
      <% if (HTProgRight and 2)=2 then %><a href="Querygroup.asp">重設查詢</a>　<% End If %>&nbsp;</td>
  </tr>
  <tr>
    <td class="Formtext" width="100%" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
	<form name="reg" method="post"> 
  <center>
  <table border="0" width="95%" cellspacing="1" cellpadding="0" class="bluetable">
    <tr>
      <td width="5%" class="lightbluetable" height="28">
        <p align="center">&nbsp;</p>
      </td>
      <td width="10%" class="lightbluetable" height="28">
        <p align="center">群組ID</p>
      </td>
      <td width="35%" class="lightbluetable" height="28">
        <p align="center">群　組　名　稱</p>
      </td>
      <td width="50%" class="lightbluetable" height="28">
        <p align="center">說　　　　明</p>
      </td>
    </tr>
<%   set rs=conn.execute(Sqlcom)
     if rs.eof then %>
    <tr class="whitetablebg">
      <td colspan="4" height="150">
        <p align="center">查無資料</p>
      </td>
    </tr>
  </table>
  </center>
    <% Else
       rec=0 
       while not rs.eof 
        rec=rec+1  %>
    <tr class="whitetablebg">
      <td colspan="4">
        <table border="0" width="100%" cellspacing="1" cellpadding="2" bgcolor="#C0C0C0">
          <tr class="whitetablebg">
		     <td width="5%" align="center"><input type="radio" value="<%=rs("ugrpId")%>" name="ugrpId"></td>
		     <td width="10%" height="20"><%=rs("ugrpId")%></td>
		     <td width="35%" height="20"><%=rs("ugrpName")%></td>
	             <td width="50%" height="20"><%=rs("Remark")%></td>
          </tr>
        </table>
      </td>
    </tr>
<%    RS.movenext
      wend %>
  </table>
   <p align="center">  
   <img src="../images/management/titlehr.gif" width="570" height="1" vspace="3"><br>  
   <% if (HTProgRight and 8)=8 then %><input type="button" value="編修代碼權限" class="cbutton" OnClick="REdit()">&nbsp;&nbsp;
   <% End If %>  
  <% End IF %>
</form>
    </td>
  </tr>
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      <!--#include virtual = "/inc/Footer.inc" -->
      </td>                                         
  </tr>                 
</body></html>                 
<script language=VBScript>
Sub REdit()                                                                       
<% if rec>1 then %>                                     
  for i=0 to reg.ugrpId.length-1                                     
    if reg.ugrpId(i).checked then                                     
      radiochk="YES"                                     
      exit for                                     
    end if                                     
  next                                     
<% else %>                                     
  if reg.ugrpId.checked then                                     
    radiochk="YES"                                     
  end if                                     
<% end if %>                                     
  if radiochk<>"YES" then                                     
    alert("尚未選定!!")                                     
  else                    
     reg.action="EditGroupByGroup.asp"
     reg.submit                                     
  end if                    
End Sub
</script>              






