<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "HT002"
   %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
function xmlImportUser(op, userid, fn, email, pw )	' op: 1-新增，2-更新，3-刪除
	xs = "<?xml version=""1.0"" encoding=""utf-8""?>"
	xs = ""
	xs = xs & "<person recstatus=""" & op & """><userid>" &  userid & "</userid><name><fn> " _
		& fn & " </fn></name><email>" & email & "</email>" _
		& "<extension><userOrgID>443</userOrgID><Account_Password>" & pw & "</Account_Password></extension>" _
		& "</person>"
	response.write xs & "<HR>" & vbCRLF
	for i=1 to len(fn)
		response.write mid(fn,i,1) & asc(mid(fn,i,1)) & "=>" & hex(asc(mid(fn,i,1))) & "<BR>"
	next
	response.write server.urlencode(fn) & "<HR>"
	xmlImportUser = xs
end function

if request("task")="新增存檔" then
	sqlCom = "SELECT * FROM InfoUser WHERE UserID = N'" & request("pfx_UserID") & "'"
	Set RS = Conn.execute(sqlCom)
	if not RS.EOF then %>
	<script language=VBScript>
	  alert("此使用者帳號已建立，無法再次建立,請另取帳號！")
	  window.history.back
	</script>	
<%	else
	SQL="Insert Into InfoUser(UserID,Password,UserName,UserType,ugrpID,deptId,jobName,email,telephone,tDataCat,uploadPath,VisitCount) Values (" _
		& drs("pfx_UserID") & drs("tfx_xPassword") & drs("tfx_UserName") & "'P'," _
		& drs("sfx_ugrpID") & drs("sfx_deptID") & drs("tfx_jobName") _
		& drs("tfx_email") & drs("tfx_telephone") & drs("sfx_tDataCat") & drs("tfx_uploadPath")& " 0)"
'	conn.execute(SQL)

'http://10.10.5.73:7777/webservices/IlaWebServices?invoke=importUser&param0=loginusername%3Ddemo_admin&
'param1=loginpassword%3Dwelcome&param2=loginsiteshortname%3DDemo&param3=xmlcontents%3D%3Cperson+recstatus%3D%221%22%3E%3Cuserid%3ExEd%3C%2Fuserid%3E%3Cname%3E%3Cfn%3EEdwin+xSmith%3C%2Ffn%3E%3C%2Fname%3E%3Cemail%3ExEdwin.Smith@oracle.com%3C%2Femail%3E%3C%2Fperson%3E&param4=

'		wsURL = "http://mofsys.hyweb.com.tw/ws/ga.asp?invoke=importUser" _
		wsURL = "http://10.10.5.73:7777/webservices/IlaWebServices?invoke=importUser" _
			& "&param0=" & server.URLEncode("loginusername=demo_admin") _
			& "&param1=" & server.URLEncode("loginpassword=welcome") _
			& "&param2=" & server.URLEncode("loginsiteshortname=Demo") _
			& "&param3=" & server.URLEncode("locale=zh_cn") _
			& "&param4=" & server.URLEncode("xmlcontents=" _
				& xmlImportUser("1", request("pfx_UserID"), request("tfx_UserName"), _
					request("tfx_email"), request("tfx_xPassword")) )


	set oXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM")
	oXmlDoc.async = false
	oXMLDoc.setProperty("ServerHTTPRequest") = true	
	
	response.write vbCRLF & vbCRLF & wsURL & "<HR>" & vbCRLF
	rv = oXmlDoc.load(wsURL)
	response.write rv & "<HR>" & vbCRLF
'	response.write server.mapPath("\ms") & "\testfile.txt"
'	response.end
	response.write oXmlDoc.documentElement.xml


		response.end




%>
		<script language=VBScript>
		  alert("新增完成！")
		  window.navigate "CtUserSet.asp?userID=<%=request("pfx_UserID")%>"
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
          <table border="0" width="95%" cellspacing="1" cellpadding="3" class="bluetable">    
            <tr>    
              <td width="20%" class="lightbluetable" align="right">使用帳號<font color=red>※</font></td>    
              <td class="whitetablebg"><input type="text" name="pfx_UserID" size="15"></td>    
              <td width="20%" class="lightbluetable" align="right">負責人員<font color=red>※</font></td>
              <td width="30%" class="whitetablebg"><input type=text name="tfx_UserName" size="15"></td>
    	    </tr>
    	    <tr>
              <td width="20%" class="lightbluetable" align="right">密碼<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg"><input type=password name="tfx_xPassword" size="10"></td>        	    
              <td width="20%" class="lightbluetable" align="right">密碼確認<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg"><input type=password name="xPassword2" size="10"></td>    
            </tr>                                                                                   
    	    <tr>
              <td width="20%" class="lightbluetable" align="right">單 位<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg">
              <Select name="sfx_deptID" size=1>
<option value="">請選擇</option>
			<%SQL="Select deptID,deptName from Dept where inUse='Y' Order by kind"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
              </td>        	    
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
              <td width="20%" class="lightbluetable" align="right">電子信箱<font color=red>※</font></td>    
              <td width="30%" class="whitetablebg"><input type=text name="tfx_email" size="30"></td>    
               <td width="20%" class="lightbluetable" align="right">聯絡電話</td>    
              <td width="30%" class="whitetablebg"><input type=text name="tfx_Telephone" size="30"></td>    
            </tr>                                                                                   
            <tr>  
              <td width="20%" class="lightbluetable" align="right" rowspan=2>權限群組<font color=red>※</font><BR>(按住CTRL鍵可複選)</td>    
              <td width="30%" class="whitetablebg" rowspan=2><select size="6" name="sfx_ugrpID" multiple>    
	          <% SQL1 = "Select * From ugrp"                                                                                                                                
	           set rs1 = conn.Execute(SQL1)                                                                                                                                
	            If rs1.EOF Then%>                                                                                                                                
	          <option value="" style="color:red">目前無群組</option>
    	      <% Else %>
	          <option value="" style="color:blue">請選擇...</option> 
    	      <% Do While Not rs1.EOF %> 
        	  <option value="<%=rs1("ugrpID")%>"><%=rs1("ugrpName")%></option> 
          	  <% rs1.MoveNext 
            	 Loop 
            	End If %>
        	  </select></td>
              <td width="20%" class="lightbluetable" align="right">職 稱</td>
              <td width="30%" class="whitetablebg"><input type=text name="tfx_JobName" size="15"></td>
            </tr>
            <tr>  
              <td width="20%" class="lightbluetable" align="right">檔案上傳路徑</td>
              <td width="30%" class="whitetablebg">/public/data/<BR/><input type=text name="tfx_uploadPath" size="50"></td>
            </tr>
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
	<!--#include virtual = "/inc/Footer.inc"--> 
      </td>                                         
  </tr> 
</table>
</body>
</html> 
<script language="vbscript">
Sub formAdd
    if reg.pfx_UserID.value="" then
    	alert "請輸入使用者帳號"
    	reg.pfx_UserID.focus
    	exit sub
    end if
    if len(reg.pfx_UserID.value)>10 then
    	alert "使用者帳號不得超過10個字元"
    	reg.pfx_UserID.focus
    	exit sub
    end if    
    if reg.tfx_xPassword.value="" then
    	alert "請輸入密碼"
    	reg.tfx_xPassword.focus
    	exit sub
    end if
    if reg.xPassword2.value="" then
    	alert "請輸入密碼確認"
    	reg.xPassword2.focus
    	exit sub
    end if
    if reg.tfx_xPassword.value<>reg.xPassword2.value then
    	alert "密碼與密碼確認不同, 請重新輸入"
    	reg.tfx_xPassword.focus
    	exit sub
    end if       
    if reg.sfx_ugrpID.value="" then
    	alert "請點選權限群組"
    	exit sub
    end if 
    if reg.sfx_deptID.value="" then
    	alert "請選單位"
    	exit sub
    end if 
    if reg.tfx_UserName.value="" then
    	alert "請輸入使用者姓名"
    	reg.tfx_UserName.focus    	
    	exit sub
    end if     
    reg.Task.value = "新增存檔"    
    reg.Submit    
End Sub
</script> 