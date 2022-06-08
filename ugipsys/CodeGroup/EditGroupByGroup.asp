<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "Pn50M06"  %>
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
   end if
End function

If Request.Form("Enter") = "確定存檔" Then     
     ugrpId = Request.Form("ugrpId")
     DataDate = Date()
     SysUser = Session("UserID")
     SQL = "Delete from UgrpCode Where ugrpId =N'"& ugrpId & "'"
     set rs = conn.Execute(SQL)     
     SQLCom = "SELECT codeId from CodeMetaDef order by codeType"
     set RSCom = Conn.execute(sqlcom)
      while not RSCom.eof
           sumvalue=0
           for i=1 to 8
               codeName=RSCom("codeId")&i
               if request.Form(codeName)="Y" then 
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
           mStart = request.Form(RSCom("codeId") & "start")
           mEnd = request.Form(RSCom("codeId") & "end")
           if trim(mEnd)="" then
			   mEnd="2079/6/6"		' SmallDate 最大值
	   end if
	   if sumvalue<>0 then		   		   
           	SQL = "INSERT INTO UgrpCode (ugrpId, codeId, rights,regdate,startdate,enddate) " & _
           		"VALUES (N'"& ugrpId &"', N'"& RSCom("codeId") &"', "& sumvalue &",N'" & date() & "',N'" & mStart & "',N'" & mEnd & "')"           
	   	Set rs = conn.execute(SQL)
	   end if
     RSCom.movenext
'response.write SQL & "<hr>"   
     wend
'response.write SQLDelete & "<hr>"     
'response.end     
     %>
		<script language="VBScript">
		  alert("Down!")
		  window.navigate "ListGroup.asp?pageno=1"
		</script>	
<%     		response.end
End IF
 
   ugrpId = Request("ugrpId")
   SQL="Select ugrpName from Ugrp where ugrpId=N'" & ugrpId & "'"
   set RSU=conn.execute(SQL)

	SQLCom = "Select Distinct codeType from CodeMetaDef order by codeType"	
	Set RS = Conn.execute(SQLCom)
	IF Not RS.EOF Then
	 APCat = RS.getrows()
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
<title>編修代碼權限群組</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
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
    <td class="Formtext" width="60%">【編修代碼權限群組(by 群組)】</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right"><% if (HTProgRight and 2)=2 then %><a href="Querygroup.asp">查詢群組</a>　<% End IF %><a href="VBScript: history.back">回上一頁</a>&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td align="right" class="whitetablebg">[ <font color="red"><%=RSU("ugrpName")%></font> ]代碼權限群組權限設定
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
    <td width="100%" colspan="2" id="rightsTable">
    <form method="POST" action="EditGroupByGroup.asp" name="reg">
    <input type="hidden" name="ugrpId" value="<%=ugrpId%>"> 
    <% If DataCk = "Y" Then 
       cno = 0 
        for i=0 to ubound(APCat,2)  
          cno = 1 + cno %> 
      <table border="0" width="100%" class="bluetable" cellspacing="0" cellpadding="3"> 
        <tr> 
          <td width="100%">
           <a href="JavaScript: chdisplay(<%=cno%>);"><img border="0" src="images/2.gif" align="absmiddle" id="I<%=cno%>">  
           <span ID="MF<%=cno%>" style="cursor:hand;color:#ffffff"><%=APCat(0,i)%></span></a></td>
        </tr>
      </table>
    <% 
	SQLCom="SELECT C2.codeId, C2.codeName,UC.rights Urights,UC.startdate Ustartdate,UC.enddate Uenddate " & _
		"FROM CodeMetaDef AS C2 " & _
		"Left Join UgrpCode UC ON C2.codeId=UC.codeId AND UC.ugrpId=N'"&ugrpId&"' " & _
		"Where C2.showOrNot=N'Y' AND C2.codeType=N'"&APCat(0,i)&"' " & _
		"ORDER BY C2.codeType DESC,convert(Int,C2.codeRank),C2.codeId "
        Set RS = Conn.execute(SQLCom)
        If Not rs.Eof Then %>
        <center>
        <table border="0" width="100%" cellspacing="1" class="bluetable" style="Display:'';" ID="M<%=cno%>" cellpadding="2">
          <tr class="lightbluetable">
            <td width="5%" align="center"></td>
            <td width="30%" align="center">代碼名稱</td>
            <td width="7%" align="center">查詢</td>
            <td width="7%" align="center">新增</td>
            <td width="7%" align="center">修改</td>
            <td width="7%" align="center">刪除</td>
            <td width="7%" align="center">列印</td>
            <td width="7%" align="center">備用A</td>            
            <td width="7%" align="center">備用B</td>
            <td width="8%" align="center">起始日期時間</td>            
            <td width="8%" align="center">終止日期時間</td>
          </tr>
           <% do while not rs.eof %>
          <tr>
            <td width="5%" align="center" class="lightbluetable">
             <input type="checkbox" name="<%=rs("codeId")%>1" value="Y" OnClick="VBScript: chkbox '<%=rs("codeId")%>'" <%=TranComp(RS("Urights"),1)%>></td>
            <td width="32%" class="whitetablebg"><A href="EditGroupByCode.asp?ugrpId=<%=ugrpId%>&codeId=<%=rs("codeId")%>"><%=rs("codeName")%></A></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("codeId")%>2" value="Y" style="display:'<%=TranComp1(RS("Urights"),1)%>'" id="x<%=rs("codeId")%>2" <%=TranComp(RS("Urights"),2)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("codeId")%>3" value="Y" style="display:'<%=TranComp1(RS("Urights"),1)%>'" id="x<%=rs("codeId")%>3" <%=TranComp(RS("Urights"),4)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("codeId")%>4" value="Y" style="display:'<%=TranComp1(RS("Urights"),1)%>'" id="x<%=rs("codeId")%>4" <%=TranComp(RS("Urights"),8)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("codeId")%>5" value="Y" style="display:'<%=TranComp1(RS("Urights"),1)%>'" id="x<%=rs("codeId")%>5" <%=TranComp(RS("Urights"),16)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("codeId")%>6" value="Y" style="display:'<%=TranComp1(RS("Urights"),1)%>'" id="x<%=rs("codeId")%>6" <%=TranComp(RS("Urights"),32)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("codeId")%>7" value="Y" style="display:'<%=TranComp1(RS("Urights"),1)%>'" id="x<%=rs("codeId")%>7" <%=TranComp(RS("Urights"),64)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("codeId")%>8" value="Y" style="display:'<%=TranComp1(RS("Urights"),1)%>'" id="x<%=rs("codeId")%>8" <%=TranComp(RS("Urights"),128)%>></td>
            <td width="7%" class="whitetablebg"><input type="text" name="<%=rs("codeId")%>start" value="<%=rs("Ustartdate")%>" size="12"></td>
            <td width="7%" class="whitetablebg"><input type="text" name="<%=rs("codeId")%>end" value="<%=rs("Uenddate")%>" size="12"></td>
          </tr>
          <%  rs.movenext      
              loop %>
        </table>
        </center>
     <% Else %>
    <table border="0" width="100%" cellspacing="1" cellpadding="3" class="bluetable" style="Display:none;" ID="M<%=cno%>">
      <tr>
        <td width="100%" class="whitetablebg" align="center">此分類下無代碼
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
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      <!--#include virtual = "/inc/Footer.inc" -->
      </td>                                         
  </tr>                    
</table>               
</body></html>               
<script language="VBScript">         
   
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
     
Sub chkbox(x)     
 if reg(x & "1").checked then     
  for i=2 to 8     
    document.all("x" & x & i).style.display=""         
  next     
 Else     
  for i=2 to 8    
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
		for i=2 to xnPos     
			document.all(xname&i).checked = xBoolean     
		next     
	end if     
	end if
end sub     
</script>         
