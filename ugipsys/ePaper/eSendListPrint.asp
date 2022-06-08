<%@ CodePage = 65001 %>
<%
Response.AddHeader "content-disposition","attachment; filename=ePaper.xls"
Response.Charset ="utf-8"
Response.ContentType = "Content-Language;content=utf-8" 
Response.ContentType = "application/vnd.ms-excel"
Response.Buffer = False
%>
<% Response.Expires = 0
HTProgCap="電子報管理"
HTProgFunc="發送清單"
HTProgCode="GW1M51"
HTProgPrefix="eSend" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
	Server.ScriptTimeOut = 5000
 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 fSql=session("epaper_sql")
 
'-- Modify By Leo  新增判斷式(發送清單、重複清單)  --- Start ---
	    if request("repeat") <> "" then
			repeat = request("repeat")
		end if
'-- Modify By Leo  新增判斷式(發送清單、重複清單)  ---  End  ---


nowPage=Request.QueryString("nowPage")  '現在頁數


'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      PerPageSize=9999999'cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=50  
      end if 
      
      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
      'strSql=server.URLEncode(fSql)
   end if    
end if   

%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<title></title>
</head>
<body>
<table border="1" width="100%" cellspacing="1" cellpadding="1" align="center">
  <tr>
		<!-- Modify By Leo  新增判斷式(發送清單、重複清單)  --- Start --->
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【 <% If repeat = "1" then HTProgFunc = "重複清單" end if %> <%=HTProgFunc%> 】</td>
		<!-- Modify By Leo  新增判斷式(發送清單、重複清單)  ---  End  --->
	  </tr>
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
  <p align="center">  
<%If not RSreg.eof then%> 
    <CENTER>
     <TABLE width=100% border="1" cellspacing="1" cellpadding="1" class=bg>                   
     <tr align=left>    
	<td class=eTableLable>eMail</td>
	<td class=eTableLable>發送日期</td>
	<td class=eTableLable>會員ID</td>
	<td class=eTableLable>姓名</td>
    </tr>	                                
<%     <!-- for j=1 to 3  //為啥要跑三次迴圈?  edited by Leo-->    
 fSql=session("epaper_sql")
Set RSreg = Conn.execute(fSql)         
    for i=1 to PerPageSize                  
%>                 
<tr>                  
<%pKey = ""
pKey = pKey & "&ePubID=" & RSreg("ePubID")
if pKey<>"" then  pKey = mid(pKey,2)%>
	<TD class=eTableContent><font size=2>
<%=RSreg("eMail")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("sendDate")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("account")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("realname")%>
</font></td>
    </tr>
    <%
         RSreg.moveNext
         if RSreg.eof then exit for 
    next 
	<!--next-->
   %>
    </TABLE>
    </CENTER>
       <!-- 程式結束 ---------------------------------->
    </td>
  </tr>  
</table> 
</form>
</body>
</html>
<%end if%>