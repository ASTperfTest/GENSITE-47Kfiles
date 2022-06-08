<%@ CodePage = 65001 %>
<%
'================ 文 件 變 更 履 歷 DOCUMENT AMENDMENT ============== $Id$ =========================
' 版號	修改日期		修訂者	變更原因			變更說明
'---------------------------------------------------------------------------------------------------------------------------------------
' 002		2008/05/30	Ahui		資安修改

HTProgCap="多媒體物件上稿"
HTProgCode="GC1AP3"
HTProgPrefix="MMO"
%>
<% response.expires = 0 %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<%
dim RSreg
taskLable="編修" & HTProgCap
dim formFunction
dim MMODataCount
dim deptCheck

if request("submittask")="編修存檔" then
	xPath = request("xxPath")
	if request("htx_deptID")<>"" then
		xdeptID="'"&request("htx_deptID")&"'"
	else
		xdeptID="null"
	end if		
	SQL="Update MMOFolder Set deptID="&xdeptID&",MMOFolderNameShow=" & pkStr(request("tfx_MMOFolderNameShow"),",") & "MMOFolderDesc=" & pkStr(request("tfx_MMOFolderDesc"),"") _
		& " where MMOFolderID=" & pkStr(request("pfx_MMOFolderID"),"")
	conn.execute(SQL)

	response.write "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/></head>"	
	response.write "<script language='javascript'>"
	response.write "alert('編修完成！');"
	response.write "window.parent.navigate('index.asp');"
	response.write "</script>"
	response.write  "</html>"	
%>

<%	response.end
elseif request("submittask")="刪除" then
	xPath = request("xxPath")
	'----ftp參數處理
	FTPErrorMSG = ""
	FTPfilePath="public/"&request("MMOSiteID")&"/"&xPath
	SQLU="Select * from MMOSite where MMOSiteID="&pkStr(request("MMOSiteID"),"")
	SET RSreg=conn.execute(SQLU)
	if not RSreg.eof then
   		xFTPIP = RSreg("UpLoadSiteFTPIP")
   		xFTPPort = RSreg("UpLoadSiteFTPPort")
   		xFTPID = RSreg("UpLoadSiteFTPID")
   		xFTPPWD = RSreg("UpLoadSiteFTPPWD")
   	end if
	'----先檢查FTP同步站台是否OK,若不OK,則end
  	if xFTPIP<>"" and xFTPPort<>"" and xFTPID<>"" and xFTPPWD<>"" then
		FTPErrorMSG = ""
		fileAction=""
		ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,"","","" 			
		if FTPErrorMSG<>"" then
%>
	<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	</head>
			<script language=VBScript>
			  alert("FTP同步站台無法連接上, 未能刪除目錄！")
			  window.history.back
			</script>	
	</html>
<% 			response.end
		end if
  	end if 
  	'----FTP處理,若刪除FTP同步站台目錄不OK,則end
	upPath = session("MMOPublic") & request("MMOSiteID") & "/" & xPath
	Set fso = CreateObject("Scripting.FileSystemObject")
	if fso.FolderExists(server.MapPath(upPath)) then
	  	if xFTPIP<>"" and xFTPPort<>"" and xFTPID<>"" and xFTPPWD<>"" then
			FTPErrorMSG = ""
			fileAction="DeleteDir"
			ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,"",FTPfilePath,"","" 	
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
   	  	end if  		
		fso.DeleteFolder(server.MapPath(upPath))
	end if 
	SQL="Delete MMOFolder where MMOSiteID="&pkStr(request("MMOSiteID"),"")&" and SUBSTRING(MMOFolderName, 2, LEN(MMOFolderName) - 1)='"&xPath&"'"   
	conn.execute(SQL)
	
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
	MMODataCount=0
	xPath = request("xPath")
	xxPath = request("xPath")
	SQL="Select * " & _
		",(Select count(*) from MMOFolder Where MMOFolderParent=M.MMOFolderName) MMOFolderChildCount,D.deptName " & _
		"from MMOFolder M " & _
		"	LEft Join Dept D ON M.deptID=D.deptID " & _
		"where MMOSiteID="&pkStr(request("MMOSiteID"),"")&" and SUBSTRING(MMOFolderName, 2, LEN(MMOFolderName) - 1)="&pkStr(xPath,"")
	SET RSreg=conn.execute(SQL)
	if isnull(RSreg("deptID")) or (Len(RSreg("deptID")) >= Len(session("deptID")) and left(RSreg("deptID"),Len(session("deptID")))=session("deptID")) then deptCheck=true	
	SQLB="Select sBaseTableName from BaseDSD where rDSDCat='MMO'"
	Set RSB=conn.execute(SQLB)
	if not RSB.eof then
	    while not RSB.eof
		SQLCheckMMODataCount="Select count(*) from "&RSB(0)&" where MMOFolderID="&RSreg("MMOFolderID")
		Set RSCheck=conn.execute(SQLCheckMMODataCount)
		if RSCheck(0) > 0 then
			MMODataCount=RSCheck(0)
		end if
		RSB.movenext
	    wend
	end if
	upPath = session("MMOPublic") & request("MMOSiteID") & "/" & xPath
	if right(upPath,1)<>"/" then	upPath = upPath & "/"
	if right(xPath,1)<>"/" then	xPath = xPath & "/"
	if left(xPath,1)="/" then	xPath = mid(xPath,2)
	
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
    <td class="FormName"><%=Title%><font size=2>【編修目錄】</td>
    <td class="FormLink" valign="top">
    	<%if (HTProgRight and 4)=4 then%>
	       <p align="right"><a href="MMOAddFolder.asp?MMOSiteID=<%=RSreg("MMOSiteID")%>&xPath=<%=request.querystring("xPath")%>">新增子目錄</a>&nbsp;	    
	<%end if%>   
    </td>
  </tr>
  <tr>
    <td colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext">目錄路徑：  <%=upPath%></td>
    <td class="FormLink" valign="top">
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
	reg.pfx_MMOFolderID.value = "<%=qqRS("MMOFolderID")%>"
	reg.tfx_MMOFolderName.value = "<%=qqRS("MMOFolderName")%>"
     	reg.tfx_MMOFolderDesc.value = "<%=qqRS("MMOFolderDesc")%>"
     	reg.tfx_MMOFolderNameShow.value = "<%=qqRS("MMOFolderNameShow")%>"
     	reg.htx_deptID.value = "<%=qqRS("deptID")%>"
end sub

Sub formEdit
msg4 = "請務必輸入「目錄名稱」，不得為空白！"

  If reg.tfx_MMOFolderName.value = Empty Then
     MsgBox msg4, 64, "Sorry!"
     reg.tfx_MMOFolderName.focus
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
</script>
<% end sub '--- showHTMLTail() ------%>
