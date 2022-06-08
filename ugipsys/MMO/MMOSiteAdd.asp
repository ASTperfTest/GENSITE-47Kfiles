<%@ CodePage = 65001 %>
<%
'================ 文 件 變 更 履 歷 DOCUMENT AMENDMENT ============== $Id$ =========================
' 版號	修改日期		修訂者	變更原因			變更說明
'---------------------------------------------------------------------------------------------------------------------------------------
' 002		2008/05/30	Ahui		資安修改

HTProgCap="多媒體物件目錄管理"
HTProgCode="GC1AP3"
HTProgPrefix="MMO"
%>
<% response.expires = 0 %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<%	
dim FTPErrorMSG
dim deptCheck
deptCheck=true

if request("submittask")="新增存檔" then	
	fldrPath=session("MMOPublic")+request("htx_MMOSiteID")	
	Set fso = CreateObject("Scripting.FileSystemObject")
	if fso.FolderExists(server.MapPath(fldrPath)) then
%>
	<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	</head>
	<script language=VBScript>
	  alert("站台名稱已建立，無法再次建立！")
	  window.history.back
	</script>	
	</html>
<% 		response.end
	else	
		'----先檢查FTP同步站台是否OK,若不OK,則end
	  	if request("htx_UpLoadSiteFTPIP")<>"" and request("htx_UpLoadSiteFTPPort")<>"" and request("htx_UpLoadSiteFTPID")<>"" and request("htx_UpLoadSiteFTPPWD")<>"" then
			FTPErrorMSG=""
			fileAction=""
			FTPfilePath="public"
			ftpDo request("htx_UpLoadSiteFTPIP"),request("htx_UpLoadSiteFTPPort"),request("htx_UpLoadSiteFTPID"),request("htx_UpLoadSiteFTPPWD"),fileAction,FTPfilePath,"","","" 			
			if FTPErrorMSG<>"" then
%>
				<script language=VBScript>
				  alert("FTP同步站台無法連接上, 未能建立站台！")
				  window.history.back
				</script>	
<% 				response.end
			end if
   	  	end if  
		'----FTP同步機制,若無法產生目錄,則end
	  	if request("htx_UpLoadSiteFTPIP")<>"" and request("htx_UpLoadSiteFTPPort")<>"" and request("htx_UpLoadSiteFTPID")<>"" and request("htx_UpLoadSiteFTPPWD")<>"" then
			FTPErrorMSG=""
			fileAction="CreateDir"
			FTPfilePath="public"
			ftpDo request("htx_UpLoadSiteFTPIP"),request("htx_UpLoadSiteFTPPort"),request("htx_UpLoadSiteFTPID"),request("htx_UpLoadSiteFTPPWD"),fileAction,FTPfilePath,request("htx_MMOSiteID"),"","" 			
			if FTPErrorMSG<>"" then
%>
				<html>
				<head>
				<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
				</head>
				<script language=VBScript>
				  alert("<%=FTPErrorMSG%>未能建立站台！")
				  window.history.back
				</script>	
				</html>
<% 				response.end
			end if
   	  	end if  	   			
		'----資料庫存檔
		SQLSeq="Select max(MMOSiteSeq2) from MMOSite"
		Set RSSeq=conn.execute(SQLSeq)
		if not isNull(RSSeq(0)) then
			xMMOSiteSeq2=cInt(RSSeq(0))+1
		else
			xMMOSiteSeq2=1
		end if
		if request("htx_deptID")<>"" then
			xdeptID="'"&request("htx_deptID")&"'"
		else
			xdeptID="null"
		end if
		SQL="Insert Into MMOSite(MMOSiteID,MMOSiteName,MMOSiteDesc,MMOSiteSeq2, " & _
			"UpLoadSiteFTPIP,UpLoadSiteFTPPort,UpLoadSiteFTPID,UpLoadSiteFTPPWD,UpLoadSiteHTTP,deptID) Values(" _
			& pkStr(request("htx_MMOSiteID"),",") & pkStr(request("htx_MMOSiteName"),",") _
			& pkStr(request("htx_MMOSiteDesc"),",") & pkStr(xMMOSiteSeq2,",") _
			& pkStr(request("htx_UpLoadSiteFTPIP"),",") & pkStr(request("htx_UpLoadSiteFTPPort"),",") _
			& pkStr(request("htx_UpLoadSiteFTPID"),",") & pkStr(request("htx_UpLoadSiteFTPPWD"),",") _
			& pkStr(request("htx_UpLoadSiteHTTP"),",") & xdeptID & ")"
		conn.execute(SQL)
		'----存根目錄資料至MMOFolder表
		SQLI = "Insert Into MMOFolder values('/'," & _
			pkstr(request("htx_MMOSiteName"),"") & ",null," & _
			pkstr(request("htx_MMOSiteID"),"") & ",null,'zzz',null,"&xdeptID&")"
		conn.execute(SQLI)		
		
		'----本機目錄產生
		Set f = fso.CreateFolder(server.MapPath(fldrPath))		

	response.write "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/></head>"	
	response.write "<script language='javascript'>"
	response.write "alert('新增完成！');"
	response.write "window.parent.navigate('index.asp');"
	response.write "</script>"
	response.write  "</html>"			
%>

<%		response.end
   	end if
else
	showHTMLHead()
	formFunction = "add"
	showForm()
	initForm()
	showHTMLTail()
end if	
%>

<% sub initForm() %>
		
<% end sub '---- initForm() ----%>

<% sub showForm() %>

<!--#include file="MMOSiteForm.inc"-->

<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Query.asp">
<title></title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td class="FormName"><%=Title%><font size=2>【新增多媒體物件存放站台】</td>
    <td class="FormLink" valign="top">
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>  
  <tr>
    <td width="80%" height=230 valign=top colspan="2">
<% end sub '--- showHTMLHead() ------%>

<% sub ShowHTMLTail() %>  
    </td>
  </tr>  
</table> 
</body>
</html>
<script language="vbscript">
sub window_onLoad
	clientInitForm
end sub

sub clientInitForm
	reg.htx_deptID.value = "<%=session("deptID")%>"
end sub
Sub formAdd
msg3 = "請務必輸入「站台ID」，不得為空白！"
msg4 = "請務必輸入「站台名稱」，不得為空白！"
  If reg.htx_MMOSiteID.value = Empty Then
     MsgBox msg3, 64, "Sorry!"
     reg.htx_MMOSiteID.focus
     Exit Sub
  End if
  If reg.htx_MMOSiteName.value = Empty Then
     MsgBox msg4, 64, "Sorry!"
     reg.htx_MMOSiteName.focus
     Exit Sub
  End if

 	reg.submitTask.value = "新增存檔"
  	reg.Submit        	
  	     
End Sub

sub formReset
	reg.reset
end sub
</script>
<% end sub '--- showHTMLTail() ------%>
