<%@ CodePage = 65001 %>
<%
HTProgCap="多媒體物件目錄管理"
HTProgCode="GC1AP3"
HTProgPrefix="MMO"
%>
<% 
response.expires = 0 
response.charset="utf-8"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<%
'================ 文 件 變 更 履 歷 DOCUMENT AMENDMENT ============== $Id$ =========================
' 版號	修改日期		修訂者	變更原因			變更說明
'---------------------------------------------------------------------------------------------------------------------------------------
' 002		2008/05/30	Ahui		資安修改

dim RSreg
taskLable="編修" & HTProgCap
dim formFunction
dim FTPErrorMSG
FTPErrorMSG=""
dim deptCheck
if request("submittask")="編修存檔" then
	if request("htx_deptID")<>"" then
		xdeptID="'"&request("htx_deptID")&"'"
	else
		xdeptID="null"
	end if
	SQL="Update MMOSite Set MMOSiteName=" & pkStr(request("htx_MMOSiteName"),",") _
		& "MMOSiteDesc=" & pkStr(request("htx_MMOSiteDesc"),",") _
		& "UpLoadSiteFTPIP=" & pkStr(request("htx_UpLoadSiteFTPIP"),",") _
		& "UpLoadSiteFTPPort=" & pkStr(request("htx_UpLoadSiteFTPPort"),",") _
		& "UpLoadSiteFTPID=" & pkStr(request("htx_UpLoadSiteFTPID"),",") _
		& "UpLoadSiteFTPPWD=" & pkStr(request("htx_UpLoadSiteFTPPWD"),",") _
		& "UpLoadSiteHTTP=" & pkStr(request("htx_UpLoadSiteHttp"),",") _
		& "deptID=" & xdeptID _
		& " where MMOSiteID='" & request("htx_MMOSiteID") & "'"
'	response.write SQL
'	response.end
	conn.execute(SQL)
	
	
	response.write "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/></head>"	
	response.write "<script language='javascript'>"
	response.write "alert('編修完成');"
	response.write "window.parent.navigate('index.asp');"
	response.write "</script>"
	response.write  "</html>"	
%>

<%	response.end
elseif request("submittask")="FTPTest" then
  	if request("htx_UpLoadSiteFTPIP")<>"" and request("htx_UpLoadSiteFTPPort")<>"" and request("htx_UpLoadSiteFTPID")<>"" and request("htx_UpLoadSiteFTPPWD")<>"" then
		FTPErrorMSG=""
		fileAction=""
		FTPfilePath="public/"&request("MMOSiteID")
		ftpDo request("htx_UpLoadSiteFTPIP"),request("htx_UpLoadSiteFTPPort"),request("htx_UpLoadSiteFTPID"),request("htx_UpLoadSiteFTPPWD"),fileAction,FTPfilePath,"","","" 			
		if FTPErrorMSG<>"" then
%>
			<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
			</head>
			<script language=VBScript>
			  alert("FTP同步站台無法正常連接！")
			  window.history.back
			</script>	
			</html>
<% 				response.end
		else
%>
			<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
			</head>
			<script language=VBScript>
			  alert("FTP同步站台可正常連接！")
			  window.history.back
			</script>	
			</html>
<% 				response.end
		end if
  	end if  
elseif request("submittask")="刪除" then
	'----先檢查FTP同步站台是否OK,若不OK,則end
  	if request("htx_UpLoadSiteFTPIP")<>"" and request("htx_UpLoadSiteFTPPort")<>"" and request("htx_UpLoadSiteFTPID")<>"" and request("htx_UpLoadSiteFTPPWD")<>"" then
		FTPErrorMSG=""
		fileAction=""
		FTPfilePath="public/"&request("MMOSiteID")
		ftpDo request("htx_UpLoadSiteFTPIP"),request("htx_UpLoadSiteFTPPort"),request("htx_UpLoadSiteFTPID"),request("htx_UpLoadSiteFTPPWD"),fileAction,FTPfilePath,"","","" 			
		if FTPErrorMSG<>"" then
%>
			<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
			</head>
			<script language=VBScript>
			  alert("FTP同步站台無法連接上, 未能刪除站台！")
			  window.history.back
			</script>	
			</html>
<% 				response.end
		end if
  	end if  
	'----FTP處理,若刪除FTP同步站台不OK,則end
  	if request("htx_UpLoadSiteFTPIP")<>"" and request("htx_UpLoadSiteFTPPort")<>"" and request("htx_UpLoadSiteFTPID")<>"" and request("htx_UpLoadSiteFTPPWD")<>"" then
		FTPErrorMSG=""
		fileAction="DeleteDir"
		FTPfilePath="public/"&request("MMOSiteID")
		ftpDo request("htx_UpLoadSiteFTPIP"),request("htx_UpLoadSiteFTPPort"),request("htx_UpLoadSiteFTPID"),request("htx_UpLoadSiteFTPPWD"),fileAction,"",FTPfilePath,"","" 			
  	end if
	if FTPErrorMSG<>"" then
%>
		<html>
		<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
		</head>
		<script language=VBScript>
		  alert("<%=FTPErrorMSG%>未能刪除站台！")
		  window.history.back
		</script>	
		</html>
<% 				response.end
	end if
	   	
	'----刪除資料庫資料
	SQLD="Delete MMOFolder where MMOSiteID=" & pKstr(request("htx_MMOSiteID"),"") & ";Delete MMOSite where MMOSiteID=" & pKstr(request("htx_MMOSiteID"),"")
	conn.execute(SQLD)
	'----刪除本機目錄
	fldrPath=session("MMOPublic")+request("htx_MMOSiteID")	
	Set fso = CreateObject("Scripting.FileSystemObject")
	if fso.FolderExists(server.MapPath(fldrPath)) then
		fso.DeleteFolder(server.MapPath(fldrPath))
	end if

	response.write "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/></head>"	
	response.write "<script language='javascript'>"
	response.write "alert('刪除完成！');"
	response.write "window.parent.navigate('index.asp');"
	response.write "</script>"
	response.write  "</html>"		
%>

<%	response.end
else
	deptCheck=false
	SQL="Select M.* " & _ 
		",(Select count(*) from Mmofolder where mmositeId=M.mmositeId) MMOSiteChildCount,D.deptName " & _
		"from Mmosite M " & _
		"	LEft Join Dept D ON M.deptId=D.deptId " & _
		"where M.mmositeId="&pKstr(request.querystring("MMOSiteID"),"")
	SET RSreg=conn.execute(SQL)
	if isnull(RSreg("deptID")) or (Len(RSreg("deptID")) >= Len(session("deptID")) and left(RSreg("deptID"),Len(session("deptID")))=session("deptID")) then deptCheck=true
	showHTMLHead()
	formFunction = "edit"
	showForm()
	initForm()
	showHTMLTail()
end if	
function qqRS(fldName)
	if request("submitTask")="" then
		xValue = RSreg(fldname)
	else
		xValue = ""
		if request("htx_"&fldName) <> "" then
			xValue = request("htx_"&fldName)
		end if
	end if
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
	end if
end function
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
    <td class="FormName"><%=Title%><font size=2>【編修多媒體物件存放站台】</td>
    <td class="FormLink" valign="top">
    	<%if (HTProgRight and 4)=4 then%>
	       <p align="right"><a href="MMOAddFolder.asp?MMOSiteID=<%=RSreg("MMOSiteID")%>&xPath=">新增子目錄</a>&nbsp;	    
	<%end if%>    
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
sub formReset
	reg.reset()
	clientInitForm
end sub

sub window_onLoad
	clientInitForm
end sub

sub clientInitForm
	reg.htx_MMOSiteID.value = "<%=qqRS("MMOSiteID")%>"
	reg.htx_MMOSiteName.value = "<%=qqRS("MMOSiteName")%>"
	reg.htx_MMOSiteDesc.value = "<%=qqRS("MMOSiteDesc")%>"
	reg.htx_UpLoadSiteFTPIP.value = "<%=qqRS("UpLoadSiteFTPIP")%>"
	reg.htx_UpLoadSiteFTPPort.value = "<%=qqRS("UpLoadSiteFTPPort")%>"
	reg.htx_UpLoadSiteFTPID.value = "<%=qqRS("UpLoadSiteFTPID")%>"
	reg.htx_UpLoadSiteFTPPWD.value = "<%=qqRS("UpLoadSiteFTPPWD")%>"
	reg.htx_UpLoadSiteHttp.value = "<%=qqRS("UpLoadSiteHttp")%>"
	reg.htx_deptID.value = "<%=qqRS("deptID")%>"
end sub

Sub formEdit
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

 	reg.submitTask.value = "編修存檔"
  	reg.Submit        	
  	     
End Sub

sub formDelSubmit()
   	deleteStr = "　你確定刪除嗎？"
	chky=msgbox("注意！"& vbcrlf & vbcrlf &deleteStr& vbcrlf , 48+1, "請注意！！")
        if chky=vbok then
		reg.submitTask.value = "刪除"
	      	reg.Submit
       	end If
end sub

sub FTPTest()
	if reg.htx_UpLoadSiteFTPIP.value = Empty or reg.htx_UpLoadSiteFTPPort.value = Empty or reg.htx_UpLoadSiteFTPID.value = Empty or reg.htx_UpLoadSiteFTPPWD.value = Empty then
		alert "FTP同步站台IP/Port/登入帳號/登入密碼等四個欄位需全部填寫, "+vbcrlf+"才能測試FTP同步站台是否正常連接!"
		exit sub
	end if
 	reg.submitTask.value = "FTPTest"
  	reg.Submit        	
end sub
</script>
<% end sub '--- showHTMLTail() ------%>
