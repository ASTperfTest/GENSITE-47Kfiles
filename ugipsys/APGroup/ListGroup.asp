<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "HT001" %>
<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
else
     Sqlcom = "Select * From Ugrp where 1=1 "
end if
	if Instr(session("ugrpId")&",", "HTSD,") = 0 then  Sqlcom = Sqlcom & " AND ugrpId<>N'HTSD'"
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>

<body>
<div id="FuncName">
	<h1><%=Title%></h1>
	<div id="Nav">
<% if (HTProgRight and 4)=4 then %><a href="Addgroup.asp">新增</a>　<% End If %>
<% if (HTProgRight and 2)=2 then %><a href="Querygroup.asp">重設查詢</a>　<% End If %>
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">權限群組查詢結果</div>

<form id="Form2" name="reg" method="post"> 
  <center>
  <table cellspacing="0" id="ListTable">
    <tr>
			<th class="First" scope="col">選擇</th>
			<th><p align="center">群組ID</p></th>
			<th><p align="center">群組名稱</p></th>
			<th><p align="center">公開</p></th>
			<th><p align="center">說明</p></th>
    </tr>
<%   set rs=conn.execute(Sqlcom)
       rec=0 
       while not rs.eof 
        rec=rec+1  %>
		<tr>
			<td class="Center">
				<input type="radio" value="<%=rs("ugrpId")%>" name="ugrpId">    </td>
			<td class="Center"><%=rs("ugrpId")%></td>
			<td class="Center"><%=rs("ugrpName")%></td>
			<td class="Center"><%=rs("isPublic")%></td>
			<td><%=rs("Remark")%></td>
		</tr>
<%    RS.movenext
      wend %>
  </table>
   <% if (HTProgRight and 8)=8 and rec <>0 then %>
   	<input type="button" value="編修群組" class="cbutton" OnClick="Edit()">
   	<input type="button" value="編修權限" class="cbutton" OnClick="REdit()">
   <% End If %>  
</form>

</body></html>                 
<script language=VBScript>             
Sub Edit()            
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
     reg.action="EditGroup.asp"    
     reg.submit                           
    end if          
End Sub             
    
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
     reg.action="EditRegGroup.asp"
     reg.submit                                     
  end if                    
End Sub    
                       
</script>              






