<% Response.Expires = 0
HTProgCap="�ǭ���ƺ޲z"
HTProgFunc="���C"
HTProgCode="PA010"
HTProgPrefix="paPsn" %>
HTProgPrefix="psEnroll" %>
<!--#INCLUDE FILE="paPsnListParam.inc" -->
<% Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("ODBCDSN")
Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------

%>
<%

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 fSql=Request.QueryString("strSql")

if fSql="" then
	fSql = "SELECT htx.psnID, htx.pName, htx.birthDay, htx.eMail, htx.tel, htx.myPassword, htx.addr" _
		& " FROM paPsnInfo AS htx" _
		& " WHERE 1=1"
	xpCondition
end if

nowPage=Request.QueryString("nowPage")  '�{�b����


'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '�`����
   if totRec>0 then 
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=10  
      end if 
      
      RSreg.PageSize=PerPageSize       '�C������

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '�`����
      strSql=server.URLEncode(fSql)
   end if    
end if   
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>�i<%=HTProgFunc%>�j</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top" align=right>
		<%if (HTProgRight and 1)=1 then%>
			<A href="paPsnQuery.asp" title="���]�d�߱���">�d��</A>
		<%end if%>
			<A href="Javascript:window.history.back();" title="�^�e��">�^�e��</A> 
    </td>    
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
 <Form name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
  <p align="center">  
<%If not RSreg.eof then%>     
     <font size="2" color="rgb(63,142,186)"> ��
     <font size="2" color="#FF0000"><%=nowPage%>/<%=totPage%></font>��|                      
        �@<font size="2" color="#FF0000"><%=totRec%></font>
       <font size="2" color="rgb(63,142,186)">��| ���ܲ�       
         <select id=GoPage size="1" style="color:#FF0000">
             <%For iPage=1 to totPage%> 
                   <option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
             <%Next%>   
         </select>      
         ��</font>           
       <% if cint(nowPage) <>1 then %>             
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">�W�@��</a> 
       <%end if%>      
       
        <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">�U�@��</a> 
        <%end if%>     
        | �C������:
       <select id=PerPage size="1" style="color:#FF0000">            
             <option value="10"<%if PerPageSize=10 then%> selected<%end if%>>10</option>                       
             <option value="20"<%if PerPageSize=20 then%> selected<%end if%>>20</option>
             <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
             <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
        </select>     
     </font>     
    <CENTER>
     <TABLE width=100% cellspacing="0" cellpadding="5" class=bg border=0>                   
     <tr align=left>    
	<td class=eTableLable>�����Ҹ�</td>
	<td class=eTableLable>�m�W</td>
	<td class=eTableLable>�X�ͤ�</td>
	<td class=eTableLable>eMail</td>
	<td class=eTableLable>�q��</td>
	<td class=eTableLable>�K�X</td>
	<td class=eTableLable>��}</td>
    </tr>	                                
<%                   
    for i=1 to PerPageSize                  
%>                  
<tr>                  
<%pKey = ""
pKey = pKey & "&psnID=" & RSreg("psnID")
if pKey<>"" then  pKey = mid(pKey,2)%>
	<TD class=eTableContent><font size=2>
	<A href="paPsnEdit.asp?<%=pKey%>">
<%=RSreg("psnID")%>
</A>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("pName")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("birthDay")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("eMail")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("tel")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("myPassword")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("addr")%>
</font></td>
    </tr>
    <%
         RSreg.moveNext
         if RSreg.eof then exit for 
    next 
   %>
    </TABLE>
    </CENTER>
       <!-- �{������ ---------------------------------->  
     </form>  
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table> 
</form>
</body>
</html>                                 
<%else%>
      <script language=vbs>
           msgbox "�䤣����, �Э��]�d�߱���!"
'	       window.history.back
      </script>
<%end if%>

<script Language=VBScript>
  dim gpKey
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&strSql=" & strSql & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>" & "&strSql=<%=strSql%>&pagesize=" & newPerPage                    
     end sub 

     sub setpKey(xv)
     	gpKey = xv
     end sub

	sub butAction(k) 
	  if gpKey = "" then 
	    alert("�Х���ܭӮ�!") 
	  else
    	select case k 
    	end select 
	  end if 
	end sub 

</script>
