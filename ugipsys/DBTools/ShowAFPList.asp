<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "HT011"  %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Function TranComp(groupcode,value)
   if (groupcode and value )=value then
      response.write "checked"
   end if
End function

Function TranComp1(groupcode,value)
   if (groupcode and value ) <> value then
      	response.write "none"
   else
   	if value=64 or value=128 then
   		response.write "block; cursor: hand;"
   	else
		response.write "''"   	
   	end if   
   end if
End function

 
   dbId = Request("dbId")
   SQL="Select dbName, dbDesc from HtDdb where dbId=" & pkStr(dbId,"")
   set RSU=conn.execute(SQL)

	SQLCom = "SELECT htx.*,(CASE WHEN NOT EXISTS (Select agrpID from agrpFp WHERE" _
		& " FPcode=htx.tableName) Then '1' ELSE '0' END) " _
		& " FROM htDEntity AS htx WHERE dbId=" & pkStr(dbId,"") _
		& " Order By tableName"		
	Set RS = Conn.execute(SQLCom)


	IF Not rs.EOF Then
	 APCat = rs.getrows()
	 Datack = "Y"
	Else
	 DataCk = ""
	End If
	
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>資料來源管理</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<table border="0" width="100%" cellspacing="1" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2"><%=Title%></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" width="60%">【資料表 - 程式 存取關聯】</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right"><% if (HTProgRight and 2)=2 then %><a href="Querygroup.asp">查詢群組</a>　<% End IF %><a href="VBScript: history.back">回上一頁</a>&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td align="right" class="whitetablebg">[ <font color=red><%=RSU("dbName")%> - <%=RSU("dbDesc")%></font>] 程式存取關聯
    </td>
    <td>
      <p align="right">
      <input type="button" value="全部關閉" Name="OpenAll" class="cbutton" OnClick="chdisplayall()">
      </p>
    </td>
  </tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>  
  <tr>
    <td width="100%" colspan="2" id=rightsTable>
    <form method="POST" action="EditRegGroup.asp" name="reg">
    <input type="hidden" name="ugrpID" value="<%=ugrpID%>"> 
<% 
	
	If DataCk = "Y" Then 
       cno = 0 
        for i=0 to ubound(APCat,2)  
          cno = 1 + cno %> 
      <table border="0" width="100%" class="bluetable" cellspacing="0" cellpadding="3"> 
        <tr> 
          <td width="100%">
           <a href="JavaScript: chdisplay(<%=cno%>);"><img border="0" src="images/2.gif" align="absmiddle" id="I<%=cno%>">  
           <SPAN ID="MF<%=cno%>" style="cursor:hand;color:#ffffff"><%=APCat(2,i) & "　" & APCat(3,i)%></SPAN></a></td>
        </tr>
      </table>
    <% 
       SQLCom = "SELECT agrp.*, agrpFp.rights As Urights " & _
       		"FROM agrpFp Left Join agrp ON agrpFp.agrpID= agrp.agrpID " & _
       		"Where agrpFp.FPcode = "& pkStr(APCat(2,i),"")  &" Order By agrpFp.agrpID"
       Set RS = Conn.execute(SQLCom)
        If Not rs.Eof Then 
%>
        <center>
        <table border="0" width="100%" cellspacing="1" class="bluetable" style="Display:'';" ID="M<%=cno%>" cellpadding="2">
          <tr class="lightbluetable">
            <td width="7%" align="center"></td>
            <td width="51%" align="center">程式名稱</td>
            <td width="7%" align="center">讀取</td>
            <td width="7%" align="center">寫入</td>            
            <td width="7%" align="center">建立</td>
            <td width="7%" align="center">刪除</td>
            <td width="7%" align="center">關聯</td>
            <td width="7%" align="center">解聯</td>
          </tr>
           <% do while not rs.eof %>
          <tr>
            <td width="7%" align="center" class="lightbluetable"><%=rs("Urights")%></td>
            <td width="51%" class="whitetablebg">　<%=rs("agrpID")%> - <%=rs("agrpName")%></td>
            <td width="7%" align="center" class="whitetablebg">
            <% if (RS("Urights") AND 1) = 1 then %>
             <input type="checkbox" name="<%=rs("agrpID")%>1" value="Y" checked>
            <% end if %>
            </td>
           <td width="7%" align="center" class="whitetablebg">
             <% if (RS("Urights") AND 2) = 2 then %>
             <input type="checkbox" name="<%=rs("agrpID")%>2" value="Y" checked>
            <% end if %>
            </td>
           <td width="7%" align="center" class="whitetablebg">
            <% if (RS("Urights") AND 4) = 4 then %>
             <input type="checkbox" name="<%=rs("agrpID")%>3" value="Y" checked>
            <% end if %>
            </td>
           <td width="7%" align="center" class="whitetablebg">
            <% if (RS("Urights") AND 8) = 8 then %>
             <input type="checkbox" name="<%=rs("agrpID")%>4" value="Y" checked>
            <% end if %>
            </td>
           <td width="7%" align="center" class="whitetablebg">
            <% if (RS("Urights") AND 16) = 16 then %>
             <input type="checkbox" name="<%=rs("agrpID")%>5" value="Y" checked>
            <% end if %>
            </td>
           <td width="7%" align="center" class="whitetablebg">
            <% if (RS("Urights") AND 32) = 32 then %>
             <input type="checkbox" name="<%=rs("agrpID")%>6" value="Y" checked>
            <% end if %>
            </td>
          </tr>
          <%  rs.movenext      
              loop %>
        </table>
        </center>
     <% Else %>
    <table border="0" width="100%" cellspacing="1" cellpadding="3" class="bluetable" style="Display:none;" ID="M<%=cno%>">
      <tr>
        <td width="100%" class="whitetablebg" align="center">此分類下無應用程式
        </td>
      </tr>
    </table>
    <%  End IF
        Next                                  
       End If %>
    <p align="center">  
    <img src="../images/management/titlehr.gif" width="570" height="1" vspace="3"><br>
    </form>
    </td>
  </tr>                 
</table>               
</body></html>               
<script language=VBScript>         
sub ChkAll(xStr)
	if document.all("ckall"&xStr).value="全選" then           '全勾
          	for i=0 to reg.elements.length-1
               		set e=document.reg.elements(i)
               		if right (e.name,1)="0" AND left(e.name,len(xStr))=xStr then
                   		e.checked=true
                   		for k=0 to ubound(APComArray,2)                  		
                   			if Left(e.name,Len(e.name)-1)=APComArray(1,k) then
                   				chkbox APComArray(1,k),APComArray(2,k),APComArray(3,k),APComArray(4,k)
						for y=1 to 8
							document.all(Left(e.name,Len(e.name)-1)&y).checked = true   
						next                       				
                   				exit for
                   			end if
                   		next
              		end if
          	next
          	document.all("ckall"&xStr).value="全不選"
      	elseif document.all("ckall"&xStr).value="全不選" then        '全不勾
          	for i=0 to reg.elements.length-1
               		set e=document.reg.elements(i)
'               		if right (e.name,1)="0"	then	msgBox e.name & "==>" & xStr
               		if right (e.name,1)="0" AND left(e.name,len(xStr))=xStr then
                   		e.checked=false
                   		for k=0 to ubound(APComArray,2)                  		
                   			if Left(e.name,Len(e.name)-1)=APComArray(1,k) then
                   				chkbox APComArray(1,k),APComArray(2,k),APComArray(3,k),APComArray(4,k)
						for y=1 to 8
							document.all(Left(e.name,Len(e.name)-1)&y).checked = false
						next    
                   				exit for
                   			end if
                   		next
              		end if
          	next
          	document.all("ckall"&xStr).value="全選"
       	end if
end sub   
sub chdisplayall()   
IF document.all("OpenAll").value="全部展開" Then   
 for i=1 to <%=cno%>          
  if document.all("M" & i).style.display="none" then         
    document.all("M" & i).style.display=""         
    document.all("I" & i).src="images/2.gif"         
  end if         
 next    
    document.all("OpenAll").value="全部關閉"    
ElseIF document.all("OpenAll").value="全部關閉" Then   
 for i=1 to <%=cno%>          
  if document.all("M" & i).style.display="" then         
    document.all("M" & i).style.display="none"         
    document.all("I" & i).src="images/1.gif"         
  end if         
 next    
    document.all("OpenAll").value="全部展開"    
End IF            
end sub         
         
sub chdisplay(k)         
  if document.all("M" & k).style.display="" then         
    document.all("M" & k).style.display="none"     
    document.all("I" & K).src="images/1.gif"         
  else         
    document.all("M" & k).style.display=""         
    document.all("I" & K).src="images/2.gif"         
  end if         
end sub         
     
Sub chkbox(x,apmask,spare64,spare128)    
    if reg(x & "0").checked then    
  	for i=1 to 8
  	    if (apmask AND 2^(i-1))=2^(i-1) then    
		document.all("x" & x & i).style.display=""  	    		   
    		if i=7 then 	
			document.all("x" & x & i).style.cursor="hand"   				
			document.all("x" & x & i).title=spare64
    		elseif i=8 then		
    			document.all("x" & x & i).style.cursor="hand"
			document.all("x" & x & i).title=spare128				
    		end if   
    	    end if  
  	next    
    Else    
  	for i=1 to 8    
    		document.all("x" & x & i).style.display="none"        
    		reg(x & i).checked=false    
  	next    
    end if    
end sub    
     
sub rightsTable_onDblClick     
	set xObj = window.event.srcElement    
	if xObj.tagname="INPUT" then 
	if xObj.type = "checkbox" then     
		xname=xObj.name     
		xnPos = cint(right(xname,1))     
		xname = left(xname, len(xname)-1)     
		xBoolean = xObj.checked     
		for i=1 to xnPos     
			document.all(xname&i).checked = xBoolean     
		next     
	end if     
	end if
end sub     
</script>         
