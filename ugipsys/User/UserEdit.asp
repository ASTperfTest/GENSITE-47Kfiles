<%@ CodePage = 65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<% Response.Expires = 0 
   HTProgCode = "HT002"
   %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%
if request("task")="編修存檔" then
	SQL="Update InfoUser Set password=" & drs("tfx_xpassword") _
		& " ugrpName=" & drs("tfx_ugrpName") _
		& " userName=" & drs("tfx_userName") _
		& " ugrpId=" & drs("sfx_ugrpId") _
		& " deptId=" & drs("sfx_deptId") _
		& " tdataCat=" & drs("sfx_tdataCat") _
		& " email=" & drs("tfx_email") _
		& " telephone=" & drs("tfx_telephone") _
		& " uploadPath=" & drs("tfx_uploadPath") _
		& " jobName=" & pkStr(request("tfx_jobName"),",")
  	if checkGIPconfig("EATWebFormAP") then
		sql = sql & "EATWebFormAP=" & pkStr(request("EATWebFormAP"),",")
  	end if
  	if checkGIPconfig("UserIPcheck") then
		sql = sql & "NetIP1=" & drn("htx_NetIP1")
		sql = sql & "NetIP2=" & drn("htx_NetIP2")
		sql = sql & "NetIP3=" & drn("htx_NetIP3")
		sql = sql & "NetIP4=" & drn("htx_NetIP4")
		sql = sql & "NetMask1=" & drn("htx_NetMask1")
		sql = sql & "NetMask2=" & drn("htx_NetMask2")
		sql = sql & "NetMask3=" & drn("htx_NetMask3")
		sql = sql & "NetMask4=" & drn("htx_NetMask4")
  	end if
	sql = left(sql,len(sql)-1) & " WHERE userId=" & pkStr(request("pfx_userId"),"")
	conn.execute(SQL)
	    if request("S")="A" then
	%>
		<script language=VBScript>
		  alert("編修完成！")		  
		  window.navigate "../APGroup/ListUser.asp?ugrpId=<%=request("ugrpId")%>&nowpage=1"
		</script>		    
	<%    else
	%>
		<script language=VBScript>
		  alert("編修完成！")		  
		  window.navigate "UserList.asp?nowpage=1"
		</script>	
<%     	    end if	
	response.end
elseif request("task")="刪除" then
	SQL="Delete from InfoUser where userId=N'" & request("pfx_userId") & "'"
	conn.execute(SQL)%>
		<script language=VBScript>
		  alert("刪除完成！")
		  window.navigate "UserList.asp?nowpage=1"
		</script>	
<%     		response.end	
else
	SQL="Select * from InfoUser where userId=N'" & request.querystring("userId") & "'"
	SET RS=conn.execute(SQL)
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>使用者管理</title>
</head>
<body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
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
    <td class="Formtext" width="60%">【編輯使用者】標有<font color=red>※</font>符號欄位為必填欄位</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right">
      <% if (HTProgRight and 8)=8 then %>      <a href="CtUserSet.asp?userId=<%=request("userId")%>">上稿權限</a><% End IF %>
      <% if (HTProgRight and 8)=8 then %>      <a href="CtUserSet2.asp?userId=<%=request("userId")%>">審稿權限</a><% End IF %>
      <% if (HTProgRight and 2)=2 then %>      <a href="UserQuery.asp">查詢</a><% End IF %>
  </tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2">
      <form method="POST" name="reg">
       <input type=hidden name="CanlendarTarget" value=""> 
       <input type=hidden name=tfx_ugrpName>        
       <input type=hidden name=task>       
       <input type=hidden name=ugrpId value=<%=request.querystring("ugrpId")%>>             
       <input type=hidden name=S value=<%=request.querystring("S")%>>                       
          <center>    
          <table border="0" width="100%" cellspacing="1" cellpadding="3" class="bluetable">    
            <tr>    
              <td width="20%" class="lightbluetable" align="right">使用帳號<font color=red>※</font></td>    
              <td class="whitetablebg"><input type="text" name="pfx_userId" size="15" class=sedit readonly></td>    
              <td width="20%" class="lightbluetable" align="right">負責人員<font color=red>※</font></td>
              <td width="30%" class="whitetablebg"><input type=text name="tfx_userName" size="15"></td>
    	    </tr>
    	    <tr>
              <td width="20%" class="lightbluetable" align="right">密碼<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg"><input type=password name="tfx_xpassword" size="10"></td>        	    
              <td width="20%" class="lightbluetable" align="right">密碼確認<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg"><input type=password name="xpassword2" size="10"></td>    
            </tr>                                                                                   
    	    <tr>
              <td width="20%" class="lightbluetable" align="right"><% if not checkGIPconfig("UserColumnUseless") then %>分 網<% end if %></td>
              <td width="30%" class="whitetablebg"><% if not checkGIPconfig("UserColumnUseless") then %>
              <Select name="sfx_tdataCat" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where codeMetaId='topDataCat' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select><% end if %>
              </td>        	    
              <td width="20%" class="lightbluetable" align="right">單 位<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg">
              <Select name="sfx_deptId" size=1>
<%
	sqlCom ="SELECT D.deptId,D.deptName,D.parent,len(D.deptId)-1,D.seq," & _
		"(Select Count(*) from Dept WHERE parent=D.deptId AND nodeKind='D') " & _
		"  FROM Dept AS D Where D.nodeKind='D' " _
		& " AND D.deptId LIKE N'" & session("deptId") & "%'" _
		& " ORDER BY len(D.deptId), D.parent, D.seq"	
'response.write SqlCom
	set RSS = conn.execute(sqlCom)
	if not RSS.EOF then
		ARYDept = RSS.getrows(300)
		glastmsglevel = 0
		genlist 0, 0, 1, 0
	        expandfrom ARYDept(cid, 0), 0, 0
	end if
%>
</select>
              </td>        	    
            </tr>                                                                                   
    	    <tr>
              <td width="20%" class="lightbluetable" align="right">電子信箱<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg"><input type=text name="tfx_email" size="30"></td>    
               <td width="20%" class="lightbluetable" align="right">聯絡電話</td>    
              <td width="30%" class="whitetablebg"><input type=text name="tfx_Telephone" size="30"></td>    
            </tr>                                                                                   
            <tr>  
              <td width="20%" class="lightbluetable" align="right" rowspan=2>權限群組<font color=red>※</font><BR>(按住CTRL鍵可複選)</td>    
              <td width="30%" class="whitetablebg" rowspan=2>
<%
	SQL1 = "Select count(*) From Ugrp WHERE 1 = 1 "
	if (HTProgRight and 64) = 0 then sql1 = sql1 & " AND isPublic = 'Y'"
	if Instr(session("ugrpId") & ",", "HTSD,") = 0 then  sql1 = sql1 & " AND ugrpId <> 'HTSD'"
	if instr(session("ugrpId") & ",", "EATWFSuper,") > 0 and Instr(session("ugrpId") & ",", "HTSD,") = 0 then
		sql1 = sql1 & " AND ugrpId in ('EATWFSuper', 'EATWebForm', 'EATWebFrm2')"
        end if
	set ts1 = conn.Execute(SQL1)
%>
	        <select size="<%= ts1(0) + 1 %>" name="sfx_ugrpId" multiple>
<%
	SQL1 = "Select * From Ugrp WHERE 1 = 1 "
	if (HTProgRight and 64) = 0 then	sql1 = sql1 & " AND isPublic = 'Y'"
	if Instr(session("ugrpId") & ",", "HTSD,") = 0 then	sql1 = sql1 & " AND ugrpId <> 'HTSD'"
	if instr(session("ugrpId") & ",", "EATWFSuper,") > 0 and Instr(session("ugrpId") & ",", "HTSD,") = 0 then
		sql1 = sql1 & " AND ugrpId in ('EATWFSuper', 'EATWebForm', 'EATWebFrm2')"
        end if
	set rs1 = conn.Execute(SQL1)
	If rs1.EOF Then
%>
	          <option value="" style="color:red">目前無群組</option>
<%
	Else
%>
	          <option value="" style="color:blue">請選擇...</option> 
<%
		Do While Not rs1.EOF
%> 
	          <option value="<%=rs1("ugrpId")%>"><%=rs1("ugrpName")%></option>
<%
			rs1.MoveNext
		Loop
	End If
%>
                </select>
              </td>
              <td width="20%" class="lightbluetable" align="right">職 稱</td>
              <td width="30%" class="whitetablebg"><input type=text name="tfx_JobName" size="15"></td>
            </tr>
            <tr>  
              <td width="20%" class="lightbluetable" align="right"><% if not checkGIPconfig("UserColumnUseless") then %>檔案上傳路徑<% end if %></td>
              <td width="30%" class="whitetablebg"><% if not checkGIPconfig("UserColumnUseless") then %>/public/data/<BR/><input type=text name="tfx_uploadPath" size="50"><% end if %></td>
            </tr>
            
<%
  	if checkGIPconfig("EATWebFormAP") then
		SQL1 = "SELECT count(*) FROM CodeMain WHERE (codeMetaId = 'WebFormAPCode')"
		set ts1 = conn.Execute(SQL1)
%>                                                                                   
            <tr>  
              <td width="20%" class="lightbluetable" align="right">EATWebForm權限<font color=red>※</font><BR>(按住CTRL鍵可複選)</td>    
              <td width="30%" class="whitetablebg">
	        <select size="<%= ts1(0) + 1 %>" name="EATWebFormAP" multiple>
		  <option value="" style="color:blue">請選擇...</option>   
<%
		SQL1 = "SELECT mcode, mvalue FROM CodeMain WHERE (codeMetaId = 'WebFormAPCode') ORDER BY  msortValue"
		set rs1 = conn.Execute(SQL1)
		While Not rs1.EOF
%> 
        	  <option value="<%=rs1("mcode")%>"><%=rs1("mvalue")%></option> 
<%
			rs1.MoveNext 
	        wend
%>
        	</select>
	      </td>
              <td width="20%" class="lightbluetable" align="right"></td>
              <td width="30%" class="whitetablebg"></td>
            </tr>

<%
	end if
%>
            
<%
  	if checkGIPconfig("UserIPcheck") then
%>
<TR><TD class="lightbluetable" align="right">IP 位址：</TD>
<TD class="whitetablebg"><input name="htx_NetIP1" size="3">
.<input name="htx_NetIP2" size="3">
.<input name="htx_NetIP3" size="3">
.<input name="htx_NetIP4" size="3">
</TD>
<TD class="lightbluetable" align="right">子網路遮罩：</TD>
<TD class="whitetablebg"><input name="htx_NetMask1" size="3">
.<input name="htx_NetMask2" size="3">
.<input name="htx_NetMask3" size="3">
.<input name="htx_NetMask4" size="3">
</TD>
</TR>
<% end if %>
          </table>
          </center>
            <p align="center">  
            <img src="../images/management/titlehr.gif" width="570" height="1" vspace="3"><br>
            <% if (HTProgRight and 8)=8 then %><input type="button" value="編修存檔" class="cbutton" onclick="VBS: formEdit"><% End If %>
            &nbsp;
            <% if (HTProgRight AND 16)<>0 then %><input type=button value ="刪　除" class="cbutton" onClick="VBS: DeleteForm"><%end If %>           
            &nbsp;
			<input type="button" value="清除重填" class="cbutton" onclick="VBS: formReset">         
      </form>          
    </td>
  </tr>  
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      </td>                                         
  </tr> 
</table>
</body>
</html> 
<script language="vbscript">
sub window_onLoad
    InitForm
end sub

sub formReset
    reg.reset
    InitForm
end sub

sub InitForm
	pxChoice = split("<%=RS("ugrpId")%>",",")	
	for ix=0 to UBound(pxChoice)
		for i=0 to reg.sfx_ugrpId.length-1
			if trim(pxChoice(ix))=reg.sfx_ugrpId.options(i).value then
				reg.sfx_ugrpId.options(i).selected=true
			end if
		next
	next
    reg.pfx_userId.value="<%=RS("userId")%>"
    reg.tfx_xpassword.value="<%=RS("password")%>"   
    reg.xpassword2.value="<%=RS("password")%>"         
    reg.tfx_userName.value="<%=RS("userName")%>"
    reg.tfx_ugrpName.value="<%=RS("ugrpName")%>"   
    reg.sfx_deptId.value="<%=RS("deptId")%>"
<% if not checkGIPconfig("UserColumnUseless") then %>
    reg.sfx_tdataCat.value="<%=RS("tdataCat")%>"
<% end if %>
    reg.tfx_jobName.value="<%=RS("jobName")%>"   
    reg.tfx_email.value="<%=RS("email")%>"   
    reg.tfx_telephone.value="<%=RS("telephone")%>"
<% if not checkGIPconfig("UserColumnUseless") then %>
    reg.tfx_uploadPath.value="<%=RS("uploadPath")%>"
<% end if %>
<%
  	if checkGIPconfig("EATWebFormAP") then
%>
	pxChoice = split("<%=RS("EATWebFormAP")%>",",")	
	for ix=0 to UBound(pxChoice)
		for i=0 to reg.EATWebFormAP.length-1
			if trim(pxChoice(ix))=reg.EATWebFormAP.options(i).value then
				reg.EATWebFormAP.options(i).selected=true
			end if
		next
	next
<% 
        end if 
%>    
<%
  	if checkGIPconfig("UserIPcheck") then
%>
    reg.htx_NetIP1.value="<%=RS("NetIP1")%>"
    reg.htx_NetIP2.value="<%=RS("NetIP2")%>"
    reg.htx_NetIP3.value="<%=RS("NetIP3")%>"
    reg.htx_NetIP4.value="<%=RS("NetIP4")%>"
    reg.htx_NetMask1.value="<%=RS("NetMask1")%>"
    reg.htx_NetMask2.value="<%=RS("NetMask2")%>"
    reg.htx_NetMask3.value="<%=RS("NetMask3")%>"
    reg.htx_NetMask4.value="<%=RS("NetMask4")%>"
<% 
        end if 
%>            
end sub

Sub formEdit
    if reg.tfx_userName.value="" then
    	alert "請輸入使用者姓名"
    	reg.tfx_userName.focus    	
    	exit sub
    end if
    if reg.tfx_xpassword.value="" then
    	alert "請輸入密碼"
    	exit sub
    end if 
    if len(reg.tfx_xpassword.value)>10 then
    	alert "密碼不能超過10個字元"
    	exit sub
    end if          
    if reg.xpassword2.value="" then
    	alert "請輸入確認密碼"
    	exit sub
    end if 
    if reg.tfx_xpassword.value<>"" then
    	if reg.tfx_xpassword.value <> reg.xpassword2.value then
    		alert "密碼與確認密碼不同, 請再次確認!"
    		exit sub
    	end if
    end if
    if reg.sfx_deptId.value="" then
    	alert "請選單位"
    	exit sub
    end if 
    reg.Task.value = "編修存檔"   
    reg.Submit
End Sub

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

sub DeleteForm
    chky=msgbox("注意！"& vbcrlf & vbcrlf &"　你確定刪除【 " & "<%=RS("userId")%>" & " 】這筆資料嗎？"& vbcrlf , 48+1, "請注意！！")
    if chky=vbok then
	reg.Task.value = "刪除"
	reg.Submit
    End If
end sub
</script> 
