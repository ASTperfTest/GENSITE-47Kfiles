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

dim xNewIdentity
'----ftp參數處理
	FTPErrorMSG=""
	FTPfilePath="public/Attachment"
	SQLP = "Select * from UpLoadSite where UpLoadSiteID='file'"
	Set RSP = conn.execute(SQLP)
	if not RSP.EOF  then
   		xFTPIP = RSP("UpLoadSiteFTPIP")
   		xFTPPort = RSP("UpLoadSiteFTPPort")
   		xFTPID = RSP("UpLoadSiteFTPID")
   		xFTPPWD = RSP("UpLoadSiteFTPPWD")
   	end if
'----ftp參數處理end 
 apath=server.mappath(HTUploadPath) & "\"
 nfname=request("nfname")
  if request("action")="del" then
 
	  if xFTPIP<>"" and xFTPID<>"" and xFTPPWD<>"" then
		fileAction="DELE"
		fileTarget=nfname
		fileSource=apath + nfname
			
		ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,"",fileTarget,fileSource	
   	  end if  
  end if
  if request("action")="add" then
          if xFTPIP<>"" and xFTPID<>"" and xFTPPWD<>"" then
		fileAction="MoveFile"
		fileTarget=nfname
		fileSource=apath + nfname
			
		ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,"",fileTarget,fileSource	
   	  end if  
  end if 	
  
   if request("action")="mod" then
          if xFTPIP<>"" and xFTPID<>"" and xFTPPWD<>"" then
                fileAction="DELE"
		fileTarget=request("ofname")
		fileSource=apath + request("ofname")
			
		ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,"",fileTarget,fileSource	
			
		fileAction="MoveFile"
		fileTarget=nfname
		fileSource=apath + nfname
			
		ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,"",fileTarget,fileSource	
   	  end if  
  end if 	  
   	  
   	  response.redirect "CuAttachList.asp?iCUItem=" & request("iCUItem")

function ftpDo(FTPIP,FTPPort,FTPID,FTPPWD,fileAction,FTPfilePath,fileDir,fileTarget,fileSource)	'----FTP機制  2004/7/7

	Set oFtp = CreateObject("FtpCom.FTP")
	oFtp.Connect FTPIP,FTPPort,FTPID,FTPPWD
	
	if left(oFtp.GetMsg,1) = "2" then 
		if FTPfilePath <> "" then oFtp.Execute("CWD " + FTPfilePath)
		if fileAction="CreateDir" then oFtp.CreateDir(fileDir)
		if fileAction="DeleteDir" then oFtp.DeleteDir(fileDir)
		if fileAction="MoveFile" then oFtp.MoveFile fileTarget,fileSource,1,0
		if fileAction="DELE" then oFtp.Execute("DELE " + fileTarget)
		if left(oFtp.GetMsg,1) <> "2" then FTPErrorMSG="  FTP機制出現錯誤,FTP未成功!"
		oFtp.LogOffServer 
		set oFtp = nothing
	else
		FTPErrorMSG="  FTP機制出現錯誤,FTP未成功!"
		oFtp.LogOffServer 
		set oFtp = nothing
	end if	
end function

 %>
