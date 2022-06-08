<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "HT001" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
   ugrpID = Request("ugrpID") 
   SQLCom = "SELECT count(*) FROM InfoUser" 
   set rs = conn.Execute(SQLCom)

if request("page_no")="" then
  pno=0
else
  pno=request("page_no")-1
end if

  Set RSreg = Server.CreateObject("ADODB.RecordSet")
  
if request("page_no")="" then
	sql = "SELECT * FROM InfoUser WHERE ugrpID Like N'%"& ugrpID & "%'"	
	session("querySQL") = sql
else
	sql = session("querySQL")
end if
	session("queryPage_no") = pno+1


'----------HyWeb GIP DB CONNECTION PATCH----------
'        RSreg.Open sql,Conn,1,1
Set RSreg = Conn.execute(sql)

'----------HyWeb GIP DB CONNECTION PATCH----------

        recno=RSreg.recordcount	

%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>權限群組成員清單</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<table border="0" width="100%" cellspacing="1" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2"><%=Title%> </td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" width="60%"><%=Remark%></td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right"><a href="Querygroup.asp">重設查詢</a>&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
<Form name=lform method="POST" action=ListUser.asp>
<table border="0" width="90%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" colspan="3"><font size="2">目前系統人員數 <font color=red><%=RS(0)%></font> 人</font>  
    <% if RSreg.eof then %>  
      <p align="center"><font color="#FF0000" face="標楷體" size="4">此群組無成員</font> 
    <% End If %> 
    </td>       
  </tr>       
<% if not RSreg.eof then %>                             
  <tr>       
    <td width="25%">       
     <p align="left"><input type=submit value="前往第" class="cbutton"><input type=text name=page_no size=3 value="<%=pno+1%>" style="text-align:center"><font size="3">頁</font>      
    </td>    
    <td width="50%" class="whitetablebg">       
<p align="center">      
　此群組共 <font color="#FF0000"><b> <%=recno%> </b></font>                
位人員　第 <font color="#FF0000">          
<%                             
  response.write pno*5+1 & "～"                             
  if pno*5+5 > recno then                             
    response.write recno                             
  else                             
    response.write pno*5+5                             
  end if                             
  pgno=int((recno-1)/5)          
  pointer=pno*5+1                              
%>          
</font>位人員       
    </td>       
    <td width="25%">       
 <p align="right">　<% if pno>0 then %><input type=button value="上五筆" onclick="VBScript: lform.page_no.value='<%=pno%>' : lform.submit" class="cbutton"><% end if %>     
   <% if pno<pgno then %><input type=button value="下五筆" onclick="VBScript: lform.page_no.value='<%=pno+2%>' : lform.submit" class="cbutton"><% end if %></p>     
    </td>       
  </tr>       
</table>       
</FORM>          
    </td>          
  </tr>          
  <tr>          
    <td width="100%" colspan="2">          
<CENTER>    
<TABLE width=80% cellspacing="1" cellpadding="3" class="bluetable">                       
<tr align=left>        
	<td align=center class="lightbluetable" width="35%">使用者帳號</td>    
	<td align=center class="lightbluetable" width="35%">使用者姓名</td>    
	<td align=center class="lightbluetable" width="15%">最近到訪</td>    
	<td align=center class="lightbluetable" width="15%">到訪次數</td>    
</tr>	                                    
<%                      
RSreg.absoluteposition=pointer                      
for i=1 to 5                      
    if not RSreg.eof then                      
%>                      
<tr>                      
	<TD class="whitetablebg" width="35%"><%=RSreg("UserID")%></TD>   
	<TD class="whitetablebg" width="35%">  
      <%=RSreg("UserName")%></TD>  
	<TD class="whitetablebg" width="15%"> 
      <p align="center"><%=d7date(RSreg("LastVisit"))%></TD>  
	<TD class="whitetablebg" width="15%"> 
      <p align="right"><%=RSreg("VisitCount")%></TD>  
</tr>  
<%  
    RSreg.moveNext  
    end if  
next  
%>  
</TABLE>  
</CENTER>  
            <p align="right">   
            <img src="../images/management/titlehr.gif" width="570" height="1" vspace="3">   
    </td>               
  </tr>                            
</table>                
</body></html>                
<% End IF %>     
<script language=VBScript>         
function lform_onsubmit()         
  if not isnumeric(lform.page_no.value) or cint(lform.page_no.value)<1 or cint(lform.page_no.value)><%=pgno+1%> then         
    alert("請輸入 1～<%=pgno+1%> 的數字")         
    lform_onsubmit=false         
    lform.page_no.focus         
  else         
    lform_onsubmit=true         
  end if         
end function         
</script>         
