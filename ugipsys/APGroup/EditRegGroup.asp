<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "HT001"  %>
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

If Request.Form("Enter") = "確定存檔" Then     
     ugrpID = Request.Form("ugrpID")
     DataDate = Date()
     SysUser = Session("UserID")
     SQL = "Delete from UgrpAp Where apcode <> N'Pn00M00' And ugrpId =N'"& ugrpID & "'"
     set rs = conn.Execute(SQL)     
     SQLCom = "SELECT apcode, apcat FROM Ap Order By apcat, aporder"
     set RSCom = Conn.execute(sqlcom)
      while not RSCom.eof
           sumvalue=0
           for i=1 to 8
               APcodeName=RSCom("APcode")&i
               if request.Form(APcodeName)="Y" then 
                    if i=1 then
                       sumvalue=sumvalue+1
                    elseif i=2 then
                       sumvalue=sumvalue+2
                    elseif i=3 then
                       sumvalue=sumvalue+4
                    elseif i=4 then
                       sumvalue=sumvalue+8
                    elseif i=5 then
                       sumvalue=sumvalue+16
                    elseif i=6 then
                       sumvalue=sumvalue+32
                    elseif i=7 then
                       sumvalue=sumvalue+64
                    elseif i=8 then
                       sumvalue=sumvalue+128                      
                    end if
               end if
           next                     
	   SQL = SQL & "INSERT INTO UgrpAp (ugrpId, apcode, rights,regdate) VALUES (N'"& ugrpID &"', N'"& RSCom("APCode") &"', "& sumvalue &",N'" & date() & "');"           
     RSCom.movenext
     wend
     if SQL<>"" then rs = conn.execute(SQL)    
     %>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"; cache=no-caches;>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>查詢結果清單</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">     
		<script language=VBScript>
		  alert("編修完成！")
		  window.navigate "ListGroup.asp"
		</script>	
</body>
</html>		
<%     		response.end
End IF
 
   ugrpID = Request.Form("ugrpID")
   SQL="Select ugrpName from Ugrp where ugrpId=N'" & ugrpID & "'"
   set RSU=conn.execute(SQL)

	SQLCom = "SELECT Apcat.*,(CASE WHEN NOT EXISTS (Select Ap.apcode from Ap Inner Join UgrpAp UAP " & _
		"ON Ap.apcode=UAP.apcode where UAP.ugrpId=N'"& ugrpID &"' " & _
		"AND Ap.apcat=Apcat.apcatId and UAP.rights>0) Then '1' ELSE '0' END) " & _
		"FROM Apcat Order By apseq"		
	Set RS = Conn.execute(SQLCom)
	IF Not rs.EOF Then
	 APCat = rs.getrows()
	 Datack = "Y"
	 	SQLCom="Select APC.apcatId,Ap.apcode,Ap.apmask,Ap.spare64,Ap.spare128 from Apcat APC Inner Join Ap ON APC.apcatId=Ap.apcat " & _
	 		"Where Ap.apmask<>0 Order By APC.apseq,Ap.apcode" 		
	 	SET RSCom=conn.execute(SQLCom)
	 	if not RSCom.EOF then
	 		APComArray=RSCom.getrows()
	 	end if	 
	Else
	 DataCk = ""
	End If
	
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>編修權限群組</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
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
    <td class="Formtext" width="60%">【編修權限群組】</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right"><% if (HTProgRight and 2)=2 then %><a href="Querygroup.asp">查詢群組</a>　<% End IF %><a href="VBScript: history.back">回上一頁</a>&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td align="right" class="whitetablebg">[ <font color=red><%=RSU("ugrpName")%></font> ]權限群組權限設定
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
    <% If DataCk = "Y" Then 
       cno = 0 
        for i=0 to ubound(APCat,2)  
          cno = 1 + cno %> 
      <table border="0" width="100%" class="bluetable" cellspacing="0" cellpadding="3"> 
        <tr> 
          <td width="100%">
           <a href="JavaScript: chdisplay(<%=cno%>);"><img border="0" src="images/2.gif" align="absmiddle" id="I<%=cno%>">  
           <SPAN ID="MF<%=cno%>" style="cursor:hand;color:#ffffff"><%=APCat(1,i) & "　" & APCat(2,i)%></SPAN></a></td>
        </tr>
      </table>
    <% 
       SQLCom = "SELECT Ap.*, UgrpAp.rights As Urights " & _
       		"FROM Ap Left Join UgrpAp ON UgrpAp.ugrpId =N'"& ugrpID &"' And Ap.apcode = UgrpAp.apcode " & _
       		"Where Ap.apcat = N'"& APCat(0,i) &"' AND apmask<>0 Order By Ap.aporder"
       Set RS = Conn.execute(SQLCom)
        If Not rs.Eof Then %>
        <center>
        <table border="0" width="100%" cellspacing="1" class="bluetable" style="Display:'';" ID="M<%=cno%>" cellpadding="2">
          <tr class="lightbluetable">
            <td width="7%" align="center"><input type=button value ="<%if APCat(4,i)="1" then%>全選<%else%>全不選<%end if%>" class="cbutton"  name=ckall<%=APCat(0,i)%> onClick="VBS:ChkAll '<%=APCat(0,i)%>'"></td>
            <td width="37%" align="center">程式名稱</td>
            <td width="7%" align="center">查詢</td>
            <td width="7%" align="center">檢視</td>            
            <td width="7%" align="center">新增</td>
            <td width="7%" align="center">修改</td>
            <td width="7%" align="center">刪除</td>
            <td width="7%" align="center">列印</td>
            <td width="7%" align="center">備用A</td>            
            <td width="7%" align="center">備用B</td>
          </tr>
           <% do while not rs.eof %>
          <tr>
            <td width="7%" align="center" class="lightbluetable">
             <input type="checkbox" name="<%=rs("APCode")%>0" value="Y" OnClick="VBScript: chkbox '<%=rs("APCode")%>','<%=rs("APmask")%>','<%=rs("spare64")%>','<%=rs("spare128")%>'" <%if RS("URights") > 0 then%> checked<%end if%>></td>
            <td width="37%" class="whitetablebg">　<%=rs("APNameC")%></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("APCode")%>2" value="Y" <%if RS("URights")=0 then%> style="display:none"<%else%> style="display:<%=TranComp1(RS("APMask"),2)%>"<%end if%> id="x<%=rs("APCode")%>2" <%=TranComp(RS("URights"),2)%>></td>            
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("APCode")%>1" value="Y" <%if RS("URights")=0 then%> style="display:none"<%else%> style="display:<%=TranComp1(RS("APMask"),1)%>"<%end if%> id="x<%=rs("APCode")%>1" <%=TranComp(RS("URights"),1)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("APCode")%>3" value="Y" <%if RS("URights")=0 then%> style="display:none"<%else%> style="display:<%=TranComp1(RS("APMask"),4)%>"<%end if%> id="x<%=rs("APCode")%>3" <%=TranComp(RS("URights"),4)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("APCode")%>4" value="Y" <%if RS("URights")=0 then%> style="display:none"<%else%> style="display:<%=TranComp1(RS("APMask"),8)%>"<%end if%> id="x<%=rs("APCode")%>4" <%=TranComp(RS("URights"),8)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("APCode")%>5" value="Y" <%if RS("URights")=0 then%> style="display:none"<%else%> style="display:<%=TranComp1(RS("APMask"),16)%>"<%end if%> id="x<%=rs("APCode")%>5" <%=TranComp(RS("URights"),16)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("APCode")%>6" value="Y" <%if RS("URights")=0 then%> style="display:none"<%else%> style="display:<%=TranComp1(RS("APMask"),32)%>"<%end if%> id="x<%=rs("APCode")%>6" <%=TranComp(RS("URights"),32)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("APCode")%>7" value="Y" <%if RS("URights")=0 then%> style="display:none"<%else%> style="display:<%=TranComp1(RS("APMask"),64)%>"<%end if%> id="x<%=rs("APCode")%>7" <%=TranComp(RS("URights"),64)%> <%if (RS("APmask") AND 64)=64 then%> title="<%=RS("spare64")%>"<%end if%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("APCode")%>8" value="Y" <%if RS("URights")=0 then%> style="display:none"<%else%> style="display:<%=TranComp1(RS("APMask"),128)%>"<%end if%> id="x<%=rs("APCode")%>8" <%=TranComp(RS("URights"),128)%> <%if (RS("APmask") AND 128)=128 then%> title="<%=RS("spare128")%>"<%end if%>></td>
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
    <% if (HTProgRight and 8)=8 then %><input type="submit" value="確定存檔" Name="Enter" class="cbutton"><% End If %>
    </form>
    </td>
  </tr>                 
</table>               
</body></html>               
<script language=VBScript>         
dim APComArray(4,<%=ubound(APComArray,2)%>)
<%for i=0 to ubound(APComArray,2)%>
    APComArray(0,<%=i%>) = "<%=APComArray(0,i)%>"
    APComArray(1,<%=i%>) = "<%=APComArray(1,i)%>"
    APComArray(2,<%=i%>) = "<%=APComArray(2,i)%>"
    APComArray(3,<%=i%>) = "<%=APComArray(3,i)%>"
    APComArray(4,<%=i%>) = "<%=APComArray(4,i)%>"                
<%next%>  
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
