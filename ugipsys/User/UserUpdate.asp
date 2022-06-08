<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "HTP02"
   %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!-- #include virtual = "/inc/checkGIPconfig.inc" -->
<%
if request("task")="編修存檔" then
	SQL="Update InfoUser Set password=" & drs("tfx_xpassword") _
		& " userName=" & drs("tfx_userName") _
		& " tdataCat=" & drs("sfx_tdataCat") _
		& " email=" & drs("tfx_email") _
		& " telephone=" & drs("tfx_telephone") _
		& " jobName=" & pkStr(request("tfx_jobName"),"") _
		& " where userId=" & pkStr(request("pfx_userId"),"")
	conn.execute(SQL)
	
	if checkGIPconfig("CheckLoginPassword") then
	         session("password") = trim(request("tfx_xpassword"))
	end if
%>
		<script>
		  alert("編修完成！")
'		  window.navigate "/default.asp"
		</script>		    
<%
	response.end
else
	SQL="Select i.*, d.deptName from InfoUser as i LEFT JOIN Dept AS d on i.deptId=d.deptId " _
		& " where userId='" & session("userId") & "'"
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
    <td class="Formtext" width="60%">【帳號基本資料】標有<font color=red>※</font>符號欄位為必填欄位</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right">
  </tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2">
      <form method="POST" name="reg">
       <input type=hidden name="CanlendarTarget" value=""> 
       <input type=hidden name=task>       
          <center>    
          <table border="0" width="95%" cellspacing="1" cellpadding="3" class="bluetable">    
            <tr>    
              <td width="20%" class="lightbluetable" align="right">使用帳號<font color=red>※</font></td>    
              <td class="whitetablebg"><input type="hidden" name="pfx_userId" size="15" class=sedit readonly><%=RS("userId")%></td>    
              <td width="20%" class="lightbluetable" align="right">單 位</td>    
              <td width="30%" class="whitetablebg"><%=RS("deptName")%>
    	    </tr>
    	    <tr>
              </td>        	    
              <td width="20%" class="lightbluetable" align="right">負責人員<font color=red>※</font></td>
              <td width="30%" class="whitetablebg"><input type=text name="tfx_userName" size="15"></td>
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
            </tr>                                                                                   
    	    <tr>
              <td width="20%" class="lightbluetable" align="right">密碼<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg"><input type=password name="tfx_xpassword" size="10"></td>        	    
              <td width="20%" class="lightbluetable" align="right">密碼確認<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg"><input type=password name="xpassword2" size="10"></td>    
            </tr>                                                                                   
    	    <tr>
              <td width="20%" class="lightbluetable" align="right">電子信箱<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg"><input type=text name="tfx_email" size="30"></td>    
               <td width="20%" class="lightbluetable" align="right">聯絡電話</td>    
              <td width="30%" class="whitetablebg"><input type=text name="tfx_Telephone" size="30"></td>    
            </tr>                                                                                   
            <tr>  
              <td width="20%" class="lightbluetable" align="right"></td>    
              <td width="30%" class="whitetablebg"></td>
              <td width="20%" class="lightbluetable" align="right">職 稱</td>
              <td width="30%" class="whitetablebg"><input type=text name="tfx_JobName" size="15"></td>
            </tr>
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
    reg.pfx_userId.value="<%=RS("userId")%>"
    reg.tfx_xpassword.value="<%=RS("password")%>"   
    reg.xpassword2.value="<%=RS("password")%>"         
    reg.tfx_userName.value="<%=RS("userName")%>"
<% if not checkGIPconfig("UserColumnUseless") then %>
    reg.sfx_tdataCat.value="<%=RS("tdataCat")%>"
<% end if %>
    reg.tfx_jobName.value="<%=RS("jobName")%>"   
    reg.tfx_email.value="<%=RS("email")%>"   
    reg.tfx_telephone.value="<%=RS("telephone")%>"   
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
    reg.Task.value = "編修存檔"   
    reg.Submit
End Sub

</script> 
