<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   response.charset="utf-8"
   HTProgCode = "webgeb1"%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
'response.write request("id")
SQLCom = "SELECT count(*) FROM InfoUser Where userType= 'H'"
set rs = conn.Execute(SQLCom)
Set RSreg = Server.CreateObject("ADODB.RecordSet")
nowPage=Request.QueryString("nowPage")  '現在頁數  
if nowpage="" then
	session("querySQL") = ""
	    SQL = "SELECT I.*, d.deptName FROM InfoUser I LEFT JOIN Dept AS d on d.deptId=I.deptId " _
	    	& " WHERE 1=1 " _
	    	& " AND I.deptId LIKE N'" & session("deptId") & "%'"
	for each x in request.form
	 if request(x) <> "" then
	  if mid(x,2,3) = "fx_" then
		select case left(x,1)
		  case "s"
		  	if x="sfx_ugrpId" then
				sql = sql & " AND (''''+replace(I.ugrpId,', ',''',''')+'''' LIKE N'%''" & request(x) & "''%')"	
		  	elseif x="sfx_deptId" then	  
				sql = sql & " AND I." & mid(x,5) & " LIKE N'" & request(x) & "%'"
		  	else		  
				sql = sql & " AND " & mid(x,5) & "=" & pkStr(request(x),"")
			end if
		  case else
			sql = sql & " AND " & mid(x,5) & " LIKE N'%" & request(x) & "%'"
		end select
	  end if
	 end if
	next
	sql = sql & " ORDER BY I.deptId"
	session("querySQL") = SQL
else
	SQL = session("querySQL")
end if
'response.write sql & "<HR>"
'response.end
'RSreg.Open sql,Conn,1,1
Set RSreg = Conn.execute(sql)

        if RSreg.EOF then%>
        <script language=VBS>
        	alert "找不到資料, 請重設查詢!"
        	window.navigate "UserQuery.asp?id=<%=request("id")%>&type=<%=request("type")%>"
        </script>
<%
		response.end
	else

   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=9999 
      end if 
      
      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
      strSql=server.URLEncode(fSql)
   end if    
 	
	end if
	
'檢查主題館上稿權限的使用者
sql_ch = "select subject_user,owner from NodeInfo where CtrootID='"&request("id")&"'"
set rs_ch = conn.execute(sql_ch)
subject_user= rs_ch("subject_user")
owner= rs_ch("owner")
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"; cache=no-caches;>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>查詢結果清單</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<Form name=reg method="POST" action="UserListSave.asp?id=<%=request("id")%>">
<input type=hidden name="type" value="<%=request("type")%>">
<input type=hidden name="deptid" value="<%=request("sfx_deptId")%>">
<table border="0" width="100%" cellspacing="1" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2">系統管理／主題館權限設定 </td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>  
  <tr>
    <td class="Formtext" width="60%">【使用者資料查詢清單】</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right">
      	<%if (HTProgRight and 2)=2 then%>      
      		<a href="UserQuery.asp?id=<%=request("id")%>&type=<%=request("type")%>">重設查詢</a>&nbsp;
        <%end if%>
      	</td>      		
  </tr>
  <tr>
    <td width="100%" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2" align=center>
  <tr>         
    <td width="100%" colspan="2">         
<CENTER>   
<TABLE width=90% cellspacing="1" cellpadding="3" class="bluetable">                      
<tr align=left>      
	<td width=5% align=center class=lightbluetable>選取</td>       
	<td align=center class="lightbluetable">帳號</td>    
	<td align=center class="lightbluetable">姓名</td> 
	<td align=center class="lightbluetable">單位</td>  	
</tr>	                                                      
<%                   
    for i=1 to PerPageSize                  
%>                 
<tr>   
	<%if request("type") = 1 then%>
	<td class=whitetablebg><p align=center><input type=radio name=ckbox value="<%=RSreg("UserID")%>" <%if instr(owner,RSreg("UserID"))>0 then response.write "checked" end if%>></td>
	<%elseif request("type") = 2 then%>
	<td class=whitetablebg><p align=center><input type=checkbox name=ckbox value="<%=RSreg("UserID")%>" <%if instr(subject_user,RSreg("UserID"))>0 then response.write "checked" end if%>></td>
	<%end if%>
	<TD class="whitetablebg"><%=RSreg("UserID")%></TD>   
	<TD class="whitetablebg"><%=RSreg("UserName")%></TD>  
	<TD class="whitetablebg"><%=RSreg("deptName")%></TD>  
</tr>  
    <%
         RSreg.moveNext
         if RSreg.eof then exit for 
    next 
   %>
<tr>
   <td class="whitetablebg" colspan=4 align="center">
   <%if request("type") = 1 then%>
   <input type = "button" value = "轉移確定" onclick = "check()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
   <%elseif request("type") = 2 then%>
   <input type = "button" value = "確定" onclick = "check()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
   <%end if%>
   <input type = "button" value = "上一頁" onclick="javascript:history.go(-1);">
   </td>
</tr>
</TABLE>  
</CENTER>  
    </td>               
  </tr>               
</table> 
<center>
</form>              
</body></html>
<script language=vbscript>
Sub check1()
	for i=0 to reg.ckbox.length-1
		if reg.ckbox(i).checked =true then
			reg.Submit
			exit sub
		end if
	next
	alert "請選擇至少一名專家帳號！"
End Sub

Sub check()
	reg.Submit
End Sub
</script> 