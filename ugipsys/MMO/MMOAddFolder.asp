<%@ CodePage = 65001 %>
<%
'================ 文 件 變 更 履 歷 DOCUMENT AMENDMENT ============== $Id$ =========================
' 版號	修改日期		修訂者	變更原因			變更說明
'---------------------------------------------------------------------------------------------------------------------------------------
' 002		2008/05/30	Ahui		資安修改

HTProgCap="多媒體物件上稿"
if request.querystring("S")="11" then
	HTProgCode="GE1T11"
elseif request.querystring("S")="P1" then
	HTProgCode="GC1AP1"
else
	HTProgCode="GC1AP3"
end if
HTProgPrefix="MMO"
%>
<% response.expires = 0 %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

taskLable="新增" & HTProgCap
if request("submittask")="新增存檔" then
	xPath = request("xxPath")
'	upPath = "site/"& session("mysiteid") & "/public/" & session("MMOPublic") & request("MMOSiteID") & "/" & xPath
	upPath = session("MMOPublic") & request("MMOSiteID") & "/" & xPath
	if right(upPath,1)<>"/" then	upPath = upPath & "/"
	
	fldrPath=upPath+request("tfx_MMOFolderName")

	Set fso = CreateObject("Scripting.FileSystemObject")
	if fso.FolderExists(server.MapPath(fldrPath)) then

%>
	<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	</head>
	<script language=VBScript>
	  alert("目錄名稱已建立，無法再次建立！")
	  window.history.back
	</script>	
	</html>
<% 	else
		'----ftp參數處理
		FTPErrorMSG = ""
		FTPfilePath="public/"&request("MMOSiteID")&"/"&xPath
		SQLU="Select * from MMOSite where MMOSiteID="&pKstr(request("MMOSiteID"),"")
		SET RSreg=conn.execute(SQLU)
		if not RSreg.eof then
	   		xFTPIP = RSreg("UpLoadSiteFTPIP")
	   		xFTPPort = RSreg("UpLoadSiteFTPPort")
	   		xFTPID = RSreg("UpLoadSiteFTPID")
	   		xFTPPWD = RSreg("UpLoadSiteFTPPWD")
	   	end if
		'----先檢查FTP同步站台是否OK,若不OK,則end
	  	if xFTPIP<>"" and xFTPPort<>"" and xFTPID<>"" and xFTPPWD<>"" then
			fileAction=""
			ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,"","","" 			
			if FTPErrorMSG<>"" then
%>
	<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	</head>				<script language=VBScript>
				  alert("FTP同步站台無法連接上, 未能建立目錄！")
				  window.history.back
				</script>	
	</html>
<% 				response.end
			end if
   	  	end if  
		'----FTP同步機制,若無法產生目錄,則end
	  	if xFTPIP<>"" and xFTPPort<>"" and xFTPID<>"" and xFTPPWD<>"" then
			fileAction="CreateDir"
			ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,request("tfx_MMOFolderName"),"","" 	
			if FTPErrorMSG<>"" then
%>
	<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	</head>				<script language=VBScript>
				  alert("<%=FTPErrorMSG%>未能建立目錄！")
				  window.history.back
				</script>	
	</html>
<% 				response.end
			end if
   	  	end if  
   	  	'----新增資料庫資料	   	
		xMMOFolderName=""
		if request("htx_deptID")<>"" then
			xdeptID="'"&request("htx_deptID")&"'"
		else
			xdeptID="null"
		end if		
		if request("xxPath")<>"" then
			xMMOFolderName="/"+request("xxPath")+"/"+request("tfx_MMOFolderName")
		else
			xMMOFolderName="/"+request("tfx_MMOFolderName")
		end if
		xMMOFolderParent="/"+request("xxPath")
		SQL="Insert Into MMOFolder Values('"&xMMOFolderName&"',"&pKstr(request("tfx_MMOFolderDesc"),"")&",null,"&pKstr(request("MMOSiteID"),"")&",null,'"&xMMOFolderParent&"',"&pKstr(request("tfx_MMOFolderNameShow"),"")&","&xdeptID&")"
		conn.execute(SQL)
		'----新增本機目錄

		Set f = fso.CreateFolder(server.MapPath(fldrPath))

		response.write "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/></head>"	
		response.write "<script language='javascript'>"
		response.write "alert('新增完成！');"
		if request.querystring("MMOFolderID")<>"" then	
			response.write "window.opener.MMOFolderAddReload();"
			response.write "window.close;"
		else
			response.write "window.parent.navigate('index.asp');"
		end if	
		response.write "</script>"
		response.write  "</html>"	
%>

<%		response.end
   	end if
else
	if request.querystring("MMOFolderID")<>"" then
		SQLM="Select MMOFolderName,MMOSiteID from MMOFolder where MMOFolderID="&pKstr(request.querystring("MMOFolderID"),"")
		set RSM=conn.execute(SQLM)
		xMMOSiteID=RSM("MMOSiteID")		
		xPath = mid(RSM("MMOFolderName"),2)
		xxPath = mid(RSM("MMOFolderName"),2)
	else
		xMMOSiteID=request.querystring("MMOSiteID")		
		xPath = request("xPath")
		xxPath = request("xPath")
	end if
	upPath = session("MMOPublic") & xMMOSiteID & "/" & xPath
	if right(upPath,1)<>"/" then	upPath = upPath & "/"
	if right(xPath,1)<>"/" then	xPath = xPath & "/"
	if left(xPath,1)="/" then	xPath = mid(xPath,2)
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

<!--#include file="MMOFolderForm.inc"-->

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
    <td width="100%" class="FormName" colspan="2"><%=Title%><font size=2>【新增目錄】</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" width="60%">目錄路徑：  <%=upPath%></td>
    <td class="FormLink" valign="top" width="40%">
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
msg4 = "請務必輸入「目錄名稱(實體路徑用)」，不得為空白！"
msg5 = "請務必輸入「目錄顯示名稱」，不得為空白！"

  If reg.tfx_MMOFolderName.value = Empty Then
     MsgBox msg4, 64, "Sorry!"
     reg.tfx_MMOFolderName.focus
     Exit Sub
  End if
  If reg.tfx_MMOFolderNameShow.value = Empty Then
     MsgBox msg5, 64, "Sorry!"
     reg.tfx_MMOFolderNameShow.focus
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
