<%@ CodePage = 65001 %>
<% 
Response.Expires = 0 
HTProgCode = "HT002"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
if request("task")="新增存檔" then
        if checkGIPconfig("SchoolAccountUB") then
		sqlCom = "SELECT mvalue FROM CodeMain WHERE codeMetaId = 'SchoolAccountUB'"
		Set RS = Conn.execute(sqlCom)
		if not RS.eof then
		        SchoolAccountUB = RS(0)
		end if
		if not isNumeric(SchoolAccountUB) then
		        SchoolAccountUB = 10
		end if
		sqlCom = "select count(*) from InfoUser where deptId = '" & replace(trim(request("sfx_deptId")), "'", "''") & "'"
		set RS = Conn.execute(sqlCom)
		if int(RS(0)) >= int(SchoolAccountUB) then
%>
	<script language=VBScript>
	  alert("單位帳號已達上限，無法繼續新增！")
	  window.history.back
	</script>
<%
			response.end
		end if
        end if
	sqlCom = "SELECT * FROM InfoUser WHERE userId = '" & request("pfx_userId") & "'"
	Set RS = Conn.execute(sqlCom)
	if not RS.EOF then %>
	<script language=VBScript>
	  alert("此使用者帳號已建立，無法再次建立,請另取帳號！")
	  window.history.back
	</script>	
<%	else
	SQL="Insert Into InfoUser(userId,password,userName,userType,ugrpId,deptId,jobName,email," _
		& "telephone,tdataCat,uploadPath,visitCount,"
	sqlValue = ") Values (" _
		& drs("pfx_userId") & drs("tfx_xpassword") & drs("tfx_userName") & "'P'," _
		& drs("sfx_ugrpId") & drs("sfx_deptId") & drs("tfx_jobName") _
		& drs("tfx_email") & drs("tfx_telephone") & drs("sfx_tdataCat") & drs("tfx_uploadPath")& " 0,"
	IF request("htx_NetIP1") <> "" Then
		sql = sql & "NetIP1" & ","
		sqlValue = sqlValue & pkStr(request("htx_NetIP1"),",")
	END IF
	IF request("htx_NetIP2") <> "" Then
		sql = sql & "NetIP2" & ","
		sqlValue = sqlValue & pkStr(request("htx_NetIP2"),",")
	END IF
	IF request("htx_NetIP3") <> "" Then
		sql = sql & "NetIP3" & ","
		sqlValue = sqlValue & pkStr(request("htx_NetIP3"),",")
	END IF
	IF request("htx_NetIP4") <> "" Then
		sql = sql & "NetIP4" & ","
		sqlValue = sqlValue & pkStr(request("htx_NetIP4"),",")
	END IF
	IF request("htx_NetMask1") <> "" Then
		sql = sql & "NetMask1" & ","
		sqlValue = sqlValue & pkStr(request("htx_NetMask1"),",")
	END IF
	IF request("htx_NetMask2") <> "" Then
		sql = sql & "NetMask2" & ","
		sqlValue = sqlValue & pkStr(request("htx_NetMask2"),",")
	END IF
	IF request("htx_NetMask3") <> "" Then
		sql = sql & "NetMask3" & ","
		sqlValue = sqlValue & pkStr(request("htx_NetMask3"),",")
	END IF
	IF request("htx_NetMask4") <> "" Then
		sql = sql & "NetMask4" & ","
		sqlValue = sqlValue & pkStr(request("htx_NetMask4"),",")
	END IF
	
	if checkGIPconfig("EATWebFormAP") then
		sql = sql & "EATWebFormAP" & ","
		sqlValue = sqlValue & pkStr(request("EATWebFormAP"),",")
	end if

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
	conn.execute(SQL)

%>
		<script language=VBScript>
		  alert("新增完成！")
		  window.navigate "CtUserSet.asp?userId=<%=request("pfx_userId")%>"
		</script>	
<%     		response.end
	end if
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
    <td class="Formtext" width="60%">【新增使用者】標有<font color=red>※</font>符號欄位為必填欄位</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right"><% if (HTProgRight and 2)=2 then %>
      <a href="UserQuery.asp">查詢</a><% End IF %>&nbsp;</td>
  </tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="95%%" colspan="2">
      <form method="POST" name="reg">
       <input type=hidden name="CanlendarTarget" value=""> 
       <input type=hidden name="task">              
          <center>    
          <table border="0" width="100%" cellspacing="1" cellpadding="3" class="bluetable">    
            <tr>    
              <td width="20%" class="lightbluetable" align="right">使用帳號<font color=red>※</font></td>    
              <td class="whitetablebg"><input type="text" name="pfx_userId" size="15"></td>    
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
<%
	SQL = "Select mcode,mvalue from CodeMain where codeMetaId='topDataCat' Order by msortValue"
	SET RSS=conn.execute(SQL)
	While not RSS.EOF
%>
                  <option value="<%=RSS(0)%>"><%=RSS(1)%></option>
<%
		RSS.movenext
	wend
%>
                </select><% end if %>
              </td>
              <td width="20%" class="lightbluetable" align="right">單 位<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg">
                <Select name="sfx_deptId" size=1>
                  <option value="">請選擇</option>
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
	SQL1 = "Select count(*) From Ugrp WHERE 1=1 "
        if (HTProgRight and 64)=0 then sql1 = sql1 & " AND isPublic='Y'"
	if Instr(session("ugrpId")&",", "HTSD,") = 0 then  sql1 = sql1 & " AND ugrpId<>'HTSD'"
	if instr(session("ugrpId") & ",", "EATWFSuper,") > 0 and Instr(session("ugrpId") & ",", "HTSD,") = 0 then
		sql1 = sql1 & " AND ugrpId in ('EATWFSuper', 'EATWebForm', 'EATWebFrm2')"
        end if
	set ts1 = conn.Execute(SQL1)
%>
	        <select size="<%= ts1(0) + 1 %>" name="sfx_ugrpId" multiple>
<%
	SQL1 = "Select * From Ugrp WHERE 1=1 "
        if (HTProgRight and 64)=0 then sql1 = sql1 & " AND isPublic='Y'"
	if Instr(session("ugrpId")&",", "HTSD,") = 0 then  sql1 = sql1 & " AND ugrpId<>'HTSD'"
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
        	  <option value="<%=rs1("ugrpId")%>"><%=rs1("UgrpName")%></option> 
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
            <% if (HTProgRight and 4)=4 then %><input type="button" value="新增存檔" class="cbutton" onclick="VBS: formAdd"><% End If %>
            					<input type="reset" value="清除重填" class="cbutton">         
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
Sub formAdd
    if reg.pfx_userId.value="" then
    	alert "請輸入使用者帳號"
    	reg.pfx_userId.focus
    	exit sub
    end if
    if len(reg.pfx_userId.value)>10 then
    	alert "使用者帳號不得超過10個字元"
    	reg.pfx_userId.focus
    	exit sub
    end if    
    if reg.tfx_xpassword.value="" then
    	alert "請輸入密碼"
    	reg.tfx_xpassword.focus
    	exit sub
    end if
    'if len(reg.tfx_xpassword.value)<7 then
    '	alert "密碼請超過7個字元"
    '	exit sub
    'end if 
    if reg.xpassword2.value="" then
    	alert "請輸入密碼確認"
    	reg.xpassword2.focus
    	exit sub
    end if
    if reg.tfx_xpassword.value<>reg.xpassword2.value then
    	alert "密碼與密碼確認不同, 請重新輸入"
    	reg.tfx_xpassword.focus
    	exit sub
    end if       
    if reg.sfx_ugrpId.value="" then
    	alert "請點選權限群組"
    	exit sub
    end if 
    if reg.sfx_deptId.value="" then
    	alert "請選單位"
    	exit sub
    end if 
    if reg.tfx_userName.value="" then
    	alert "請輸入使用者姓名"
    	reg.tfx_userName.focus    	
    	exit sub
    end if     
    reg.Task.value = "新增存檔"    
    reg.Submit    
End Sub
</script> 
