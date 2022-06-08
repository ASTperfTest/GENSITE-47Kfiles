<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "HTP02"
   %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
if request("task")="編修存檔" then

	SQL="Update InfoUser Set Password=" & drs("tfx_xPassword") _
		& " UserName=" & drs("tfx_UserName") _
		& " tDataCat=" & drs("sfx_tDataCat") _
		& " email=" & drs("tfx_email") _
		& " telephone=" & drs("tfx_telephone") _
		& " jobName=" & pkStr(request("tfx_jobName"),"") & ","
		
	'判斷是否有記錄密碼更新日期-------------------------------------------------------------------Start
	
	xsql="sp_columns @table_name = 'InfoUser' , @column_name ='ModifyPassword'"
	set xRS= conn.execute(xsql)
	
	if not xRS.eof then 	
			SQL = SQL & " ModifyPassword=" & pkStr(Year(Now)& "/" & Month(Now) & "/" & Day(Now),"")
	end if
	'判斷是否有記錄密碼更新日期---------------------------------------------------------------------End
	
		SQL = SQL & " where UserID=" & pkStr(request("pfx_UserID"),"")
		'response.Write SQL
		'response.end
	conn.execute(SQL)
%>
		<script language=VBScript>
		  alert("編修完成！")		  
'		  window.navigate "/default.asp"
		</script>		    
<%
	response.end
else
	SQL="Select i.*, d.deptName from InfoUser as i LEFT JOIN dept AS d on i.deptID=d.deptID " _
		& " where UserID=N'" & session("UserID") & "'"
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
              <td class="whitetablebg"><input type="hidden" name="pfx_UserID" size="15" class=sedit readonly><%=RS("UserID")%></td>    
              <td width="20%" class="lightbluetable" align="right">單 位</td>    
              <td width="30%" class="whitetablebg"><%=RS("deptName")%>
    	    </tr>
    	    <tr>
              </td>        	    
              <td width="20%" class="lightbluetable" align="right">負責人員<font color=red>※</font></td>
              <td width="30%" class="whitetablebg"><input type=text name="tfx_UserName" size="15"></td>
              <td width="20%" class="lightbluetable" align="right">分 網</td>    
              <td width="30%" class="whitetablebg">
              <Select name="sfx_tDataCat" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from CodeMain where codeMetaID='topDataCat' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
              </td>        	    
            </tr>                                                                                   
    	    <tr>
              <td width="20%" class="lightbluetable" align="right">密碼<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg"><input type=password name="tfx_xPassword" size="10"></td>        	    
              <td width="20%" class="lightbluetable" align="right">密碼確認<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg"><input type=password name="xPassword2" size="10"></td>    
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
            <tr>
            <input type=hidden name="OriginalPassword" size="10">
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
    reg.pfx_UserID.value="<%=RS("UserID")%>"
    reg.tfx_xPassword.value="<%=RS("Password")%>"  
    reg.OriginalPassword.value="<%=RS("Password")%>"
    reg.xPassword2.value="<%=RS("Password")%>"
    reg.tfx_UserName.value="<%=RS("UserName")%>"
    reg.sfx_tDataCat.value="<%=RS("tDataCat")%>"   
    reg.tfx_jobName.value="<%=RS("jobName")%>"   
    reg.tfx_email.value="<%=RS("email")%>"   
    reg.tfx_telephone.value="<%=RS("telephone")%>"   
end sub

Sub formEdit
    if reg.tfx_UserName.value="" then
    	alert "請輸入使用者姓名"
    	reg.tfx_UserName.focus    	
    	exit sub
    end if
    if reg.tfx_xPassword.value="" then
    	alert "請輸入密碼"
    	exit sub
    end if 
    if reg.xPassword2.value="" then
    	alert "請輸入確認密碼"
    	exit sub
    end if
    if len(reg.tfx_xPassword.value)<7 then
    	alert "密碼請超過7個字元"
    	exit sub
    end if 
    if len(reg.tfx_xPassword.value)>10 then
    	alert "密碼不能超過10個字元"
    	exit sub
    end if              
    if reg.tfx_xPassword.value<>"" then
    	if reg.tfx_xPassword.value <> reg.xPassword2.value then
    		alert "密碼與確認密碼不同, 請再次確認!"
    		exit sub
    	end if
    end if    
  	if reg.tfx_xPassword.value = reg.OriginalPassword.value then
  		alert "密碼與原密碼相同, 請輸入不同密碼!"
  		exit sub
  	end if
    reg.Task.value = "編修存檔"   
    reg.Submit
End Sub

</script> 
