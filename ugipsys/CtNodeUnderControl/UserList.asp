<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "HT002"%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
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

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open sql,Conn,1,1
Set RSreg = Conn.execute(sql)

'----------HyWeb GIP DB CONNECTION PATCH----------


        if RSreg.EOF then%>
        <script language=VBS>
        	alert "找不到資料, 請重設查詢!"
        	window.navigate "UserQuery.asp"
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
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"; cache=no-caches;>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>查詢結果清單</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<Form name=reg method="POST" action=UserListSave.asp>
<input type=hidden name=submittask>
<input type=hidden name=tfx_ugrpName>  
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
    <td class="Formtext" colspan="2">【使用者資料查詢清單】</td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2" align=center>
<table border="0" width="500" cellspacing="0" cellpadding="0">
   <tr>
    <td align=center> 
     <font size="2" color="rgb(63,142,186)"> 第
     <font size="2" color="#FF0000"><%=nowPage%>/<%=totPage%></font>頁|                      
        共<font size="2" color="#FF0000"><%=totRec%></font>
       <font size="2" color="rgb(63,142,186)">筆| 跳至第       
         <select id=GoPage size="1" style="color:#FF0000">
             <%For iPage=1 to totPage%> 
                   <option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
             <%Next%>   
         </select>      
         頁</font>           
       <% if cint(nowPage) <>1 then %>             
            |<a href="UserList.asp?nowPage=<%=(nowPage-1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">上一頁</a> 
       <%end if%>      
       
        <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
            |<a href="UserList.asp?nowPage=<%=(nowPage+1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">下一頁</a> 
        <%end if%>     
        | 每頁筆數:
       <select id=PerPage size="1" style="color:#FF0000">        
             <option value="9999"<%if PerPageSize=9999 then%> selected<%end if%>>全部列出</option>                                            
             <option value="100"<%if PerPageSize=100 then%> selected<%end if%>>100</option>                       
             <option value="200"<%if PerPageSize=200 then%> selected<%end if%>>200</option>
             <option value="300"<%if PerPageSize=300 then%> selected<%end if%>>300</option>
        </select>     
     </font></td>
    <td align=right>
    </td>       
    </tr>
    </table>
  <tr>         
    <td width="100%" colspan="2">         
<CENTER>   
<TABLE width=90% cellspacing="1" cellpadding="3" class="bluetable">                      
<tr align=left>      
	<td width=5% align=center class=lightbluetable><input type=button value ="全選" class="cbutton"  name=ckall onClick="ChkAll"></td>       
	<td align=center class="lightbluetable">帳號</td>    
	<td align=center class="lightbluetable">姓名</td> 
	<td align=center class="lightbluetable">單位</td> 
	<td align=center class="lightbluetable">最近到訪</td>    
	<td align=center class="lightbluetable">到訪次數</td>       	
</tr>	                                                      
<%                   
    for i=1 to PerPageSize                  
%>                 
<tr>   
        <td class=whitetablebg><p align=center><input type=checkbox name=ckbox<%=i%>>
           <input type=hidden name=pfx_UserID<%=i%> value="<%=RSreg("UserID")%>">                       
	<TD class="whitetablebg"><a href="CtNodeControl.asp?UserID=<%=RSreg("UserID")%>"><%=RSreg("UserID")%></a></TD>
	<TD class="whitetablebg"><%=RSreg("UserName")%></TD>  
	<TD class="whitetablebg"><%=RSreg("deptName")%></TD>  
	<TD class="whitetablebg"><p align="center"><%=d7date(RSreg("LastVisit"))%></TD>  
	<TD class="whitetablebg"><p align="right"><%=RSreg("VisitCount")%></TD>  
</tr>  
    <%
         RSreg.moveNext
         if RSreg.eof then exit for 
    next 
   %>
</TABLE>  
</CENTER>  
            <p align="right">   
            <img src="../images/management/titlehr.gif" width="570" height="1" vspace="3">   
    </td>               
  </tr>               
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      </td>                                         
  </tr>             
</table> 
<center>
<div id="div2" class="divpop"  style="position: absolute; left: 250; top: 160;display:none">  
  <table width="200" height="110" bordercolordark="#000066" bordercolorlight="#000066" border="2">
    <tr>
      <td> 
        <table cellpadding="5" cellspacing="1" bgcolor="#000066" align="center">
               <tr>                    
            <td align="right" height="25" width="100" class="lightbluetable" valign="top">權限群組</td>                             
                  
            <td height="25" width="100" class="lightbluetable" valign="top"> 
              <select name=sfx_ugrpId size=4 multiple>   
	          <% SQL = "Select * From Ugrp where isPublic='Y'"                                                                                                                                
	             set rs1 = conn.Execute(SQL)
	            If rs1.EOF Then%>                                                                                                                                
	          <option value="" style="color:red">目前無群組</option>
    	      <% Else %>
	          <option value="" style="color:blue">請選擇...</option> 
    	      <% Do While Not rs1.EOF %> 
        	  <option value="<%=rs1("ugrpId")%>"><%=rs1("ugrpName")%></option> 
          	  <% rs1.MoveNext 
            	 Loop 
            	End If %>
        	  </select>
            </td>  
               </tr>                       
          <tr align="center"> 
            <td colspan="2" height="37" class="lightbluetable"> 
              <input type=button name=div2sure  class=cbutton value="確定">  
                <input type=button name=div2Close class=cbutton value="取消">
            </td>  
               </tr>  
        </table>            
      </td>  
    </tr>  
  </table>  
</div>
</form>              
</body></html>               
<script language=VBScript>  
      Dim chkCount
      chkCount=0            '記錄checkbox 被勾數   
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="UserList.asp?nowPage=" & newPage & "&strSql=" & strSql & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="UserList.asp?nowPage=<%=nowPage%>" & "&strSql=<%=strSql%>&pagesize=" & newPerPage                    
     end sub       
sub div2close_onClick
		document.all("div2").style.display="none" 	
end sub   

sub div2Sure_onClick
msg2 = "請點選權限群組，不得為空白！"
  If reg.sfx_ugrpId.value = Empty Then
     MsgBox msg2, 64, "Sorry!"
     reg.sfx_ugrpId.focus
     Exit Sub
  End if
	reg.submittask.value="批次存檔"	
	reg.submit
end sub 

sub sfx_ugrpId_onChange
    if reg.sfx_ugrpId.value="" then
    	alert "請點選權限群組"
    	exit sub
    end if
    reg.tfx_ugrpName.value=""    
    for i=0 to reg.sfx_ugrpId.length-1
	if reg.sfx_ugrpId.options(i).selected=true then
		reg.tfx_ugrpName.value=reg.tfx_ugrpName.value&reg.sfx_ugrpId.options(i).text&","
	end if
    next 
    reg.tfx_ugrpName.value=left(reg.tfx_ugrpName.value,len(reg.tfx_ugrpName.value)-1)    
end sub 

sub button1_onClick
	document.all("RunJob").style.visibility="hidden"
	document.all("div2").style.display="block"
end sub     
     
    sub document_onClick           'checkbox 被勾計數
         set sObj=window.event.srcElement
         if sObj.tagName="INPUT" then 
            if sObj.type="checkbox"  then 
                if sObj.checked then 
                   chkCount=chkCount+1
                else
                   chkCount=chkCount-1                
                end if                                          
            end if
             '
            if chkCount=0 then 
               document.all("RunJob").style.visibility="hidden"
		document.all("div2").style.display="none"               
            else
               document.all("RunJob").style.visibility="visible"
            end if
         end if        
    end sub          
    
   sub Chkall
       chkCount=0     
       if document.all("ckall").value="全選" then           '全勾
          for i=0 to reg.elements.length-1
               set e=document.reg.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=true
                   chkCount=chkCount+1
              end if     
          next                 
          document.all("RunJob").style.visibility="visible"
          document.all("ckall").value="全不選"
      elseif document.all("ckall").value="全不選" then        '全不勾
          for i=0 to reg.elements.length-1
               set e=document.reg.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=false
               end if     
          next                 
          document.all("RunJob").style.visibility="hidden"
          document.all("ckall").value="全選"          
       end if
   end sub        
</script>        
