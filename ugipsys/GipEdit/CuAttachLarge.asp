<%@ CodePage = 65001 %>
<% Response.Expires = 0
Server.ScriptTimeOut = 3000
HTProgCap="資料附件"
HTProgFunc="新增"
HTUploadPath=session("Public")+"Attachment/"
HTProgCode="GC1AP1"
HTProgPrefix="CuAttach" %>
<!--#include virtual = "/inc/server.inc" -->
<%
   SQLP = "Select mValue from CodeMain where codeMetaID='AttachmentLarge' and mCode='1'"
   Set RSP = conn.execute(SQLP)
   xFileitemCount = 0
   if RSP.eof then%>
	<script language=vbs>
		alert "請於代碼維護設定大型物件存放路徑!"
		window.close
	</script>"  
<%	response.end	   	
   else
   	folderspec = RSP(0)
   end if	
   Set fso = CreateObject("Scripting.FileSystemObject")
   if not fso.FolderExists(folderspec) then%>
	<script language=vbs>
		alert "大型物件存放路徑不存在!"
		window.close
	</script>"  
<%	response.end	   	
   else
   	Set f = fso.GetFolder(folderspec)
   end if
   Set fc = f.Files
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>附件上傳 - 大型物件清單</title>
<link href="css/popup.css" rel="stylesheet" type="text/css">
<link href="css/editor.css" rel="stylesheet" type="text/css">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="PopFormName">附件上傳 - 大型物件清單</div>
<form name="form1" method="" action="">
<table width="100%" cellspacing="0" id="PopList">

  <tr>
	<th scope="col">&nbsp;</td>
	<th scope="col">檔案名稱</td>
	<th scope="col">檔案大小(MB)</td>
  </tr>
<%For Each f1 in fc
	xFileitemCount = xFileitemCount + 1
%>
  <tr>
    <td class="Center">  
    		<input type="radio" value="<%=f1.name%>" name="giCuItem">
    </td>
    <td class="Left"><%=f1.name%></td>
    <td class="right"><%=formatNumber((f1.size/1024/1024),2)%></td>
  </tr>
<%Next%>  
</table>
  <input type="button" class="InputButton" value="確定" onClick="closeForm()">
</form>
</body>
</html>
<script language=vbs>
sub closeForm()
	xFileitemCount = <%=xFileitemCount%>
        pickRadio = null
	if xFileitemCount>1 then
	for i=0 to form1.giCuItem.length-1                           
	    if form1.giCuItem(i).checked then 
		set pickRadio = form1.giCuItem(i)                       
		exit for                           
	    end if                           
	next  
	else
		if form1.giCuItem.checked then	
		set pickRadio =  form1.giCuItem                      
		end if
	end if
  	if not isNull(pickRadio) then 
		window.opener.doAttachLarge pickRadio.value
		window.close
	else
		alert "請選取物件!"
		exit sub
	end if
end sub
</script>