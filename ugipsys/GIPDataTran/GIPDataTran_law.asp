<%@ CodePage = 65001 %>
<% Response.Expires = 0%>
<!--#include virtual = "/inc/dbutil.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
</head>
<%
'----941117行政院知識分類網[行政院法規公佈區]主題單元轉檔

Server.ScriptTimeout = 1200
Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


function nullText(xNode)
  	on error resume next
  	xstr = ""
  	xstr = xNode.text
  	nullText = xStr
  	err.number=0
end function
function xStdTime2(dt)    
   if Len(dt)=0 or isnull(dt) then
     	xStdTime2=""
   else
        xyear = cstr(year(dt))     '補零
        xmonth = right("00"+ cstr(month(dt)),2)
        xday = right("00"+ cstr(day(dt)),2)
        xhour = right("00" + cstr(hour(dt)),2)
        xminute = right("00" + cstr(minute(dt)),2)
        xsecond = right("00" + cstr(second(dt)),2)
        xStdTime2 = xyear & "/" & xmonth & "/" & xday & " " & xhour & ":" & xminute & ":" & xsecond
   end if
end function

	HTUploadPath=session("Public")+"data/"
	targetPath=server.mappath(HTUploadPath) & "\"						'----主圖與檔案下載式檔案實體路徑
	HTUploadPath2=session("Public")+"Attachment/"
	targetPath2=server.mappath(HTUploadPath2) & "\"						'----附件檔案實體路徑
	sourcePath="D:\project\uWebey\upload\"			'----轉入檔案實體路徑

    if request("mySiteID") <> session("mySiteID") then
    	response.write "session不正確, 請重新登入後台!"
    	response.end
    end if

'----1.取得轉檔所需參數
	sql = "SELECT u.* FROM CtUnit AS u " _
		& " WHERE u.CtUnitID=" & request.querystring("CtUnitID")	
	set RS = Conn.execute(sql)
	if RS.Eof then
		response.write "無此主題單元ID!"
		response.end
	elseif isNull(RS("iBaseDSD")) then
		response.write "此主題單元尚未指定資料範本!"
		response.end
	else
		xCtUnitID = RS("CtUnitID")
		xctUnitName = RS("CtUnitName")
		xiBaseDSD = RS("iBaseDSD")
	end if
	'----開檔案
	outFile = server.MapPath("LOG/DataTranLog_"&xCtUnitID&".txt")
	Set fso = CreateObject("scripting.filesystemobject")
	set xfout = fso.createTextFile(outFile)	
	xfout.writeline "主題單元"&xCtUnitID&"資料轉檔("&date()&")"
	xfout.writeline "----------------------------------------------------------------------------"
'----2.Load DSD xml
	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true		
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(xCtUnitID) & ".xml")
	if fso.FileExists(filePath) then
		LoadXML = filePath
	else
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & cStr(xiBaseDSD) & ".xml")
	end if   
	xv = htPageDom.load(LoadXML)
  	if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reason: " &  htPageDom.parseError.reason)
    		Response.End()
  	end if
	xtableName = htPageDom.selectSingleNode("DataSchemaDef/dsTable/tableName").text
'----3.先刪除前次轉檔相同主題單元資料(注意不能刪到非轉檔而來上稿資料)
	'----刪除附件檔案
	SQLDA = "Select C.NFileName,C.xiCuItem from GIPDataTran G " & _
			"    Left Join GIPDataTranDetail GD ON G.TID=GD.TID " & _
			"    Left Join CuDTAttach C ON GD.iCuItem=C.xiCuItem " & _
			"where C.xiCuItem is not null and XMLCtUnitID=" & xCtUnitID & " AND XMLmySiteID='"&session("mySiteID")&"'"
	Set RSI = conn.execute(SQLDA)
	if not RSI.eof then
		while not RSI.eof
			if not isNull(RSI("NFileName")) then 
				if fso.FileExists(targetPath2+RSI("NFileName")) then 
					fso.DeleteFile(targetPath2+RSI("NFileName"))
					conn.execute("Delete I from CuDTAttach C Left Join ImageFile I ON C.NFileName=I.NewFileName where C.xiCuItem=" & RSI("xiCuItem") & ";Delete CuDTAttach where xiCuItem=" & RSI("xiCuItem")&";")
				end if
			end if
			RSI.movenext
		wend
	end if
	'----刪除Detail表
	SQLDL = "Delete C from GIPDataTran G " & _
			"    Left Join GIPDataTranDetail GD ON G.TID=GD.TID " & _
			"    Left Join "&xtableName&" C ON GD.iCuItem=C.giCuItem " & _
			"where C.giCuItem is not null and XMLCtUnitID=" & xCtUnitID & " AND XMLmySiteID='"&session("mySiteID")&"'"
	conn.execute(SQLDL)
	'----刪除CuDTGeneric表
	SQLDM = "Delete C from GIPDataTran G " & _
			"    Left Join GIPDataTranDetail GD ON G.TID=GD.TID " & _
			"    Left Join CuDTGeneric C ON GD.iCuItem=C.iCuItem " & _
			"where C.iCuItem is not null and XMLCtUnitID=" & xCtUnitID & " AND XMLmySiteID='"&session("mySiteID")&"'"
	conn.execute(SQLDM)		
	'----刪除GIPDataTranDetail表
	SQLDGL = "Delete GD from GIPDataTran G " & _
			"    Left Join GIPDataTranDetail GD ON G.TID=GD.TID " & _
			"where XMLCtUnitID=" & xCtUnitID & " AND XMLmySiteID='"&session("mySiteID")&"'"
	conn.execute(SQLDGL)
	'----刪除GIPDataImport表
	SQLDGM = "Delete GIPDataTran where XMLCtUnitID=" & xCtUnitID & " AND XMLmySiteID='"&session("mySiteID")&"'"
	conn.execute(SQLDGM)
	response.write "清除前次轉檔資料完成!<hr>"
'	response.end	
'----4.Loop 來源recordsets,新增資料
	'----寫入GIPDataTran表		
	SQLG = "set nocount on;Insert Into GIPDataTran values(null,'"&session("UserName")&"','"&date()&"','"&session("mySiteID")&"',"&xCtUnitID&",null,null) select @@IDENTITY as NewID;"
	SET RSG = conn.execute(SQLG)
	xTID = RSG(0)
	recordCount=0
	TranSuccess=0
	TranFail=0
	SQLL="Select * from [Webey].dbo.law where 1=1 order by seq desc"
	Set RSL=conn.execute(SQLL)
	if not RSL.eof then
		while not RSL.eof 
			recordCount=recordCount+1
			err.number=0
			on error resume next
	    	Set Conn2 = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'			Conn2.Open session("ODBCDSN")
'Set Conn2 = Server.CreateObject("HyWebDB3.dbExecute")
Conn2.ConnectionString = session("ODBCDSN")
Conn2.ConnectionTimeout=0
Conn2.CursorLocation = 3
Conn2.open
'----------HyWeb GIP DB CONNECTION PATCH----------

			conn2.begintrans
	    	'----新增CuDTGeneric處理
			sql = "INSERT INTO  "&targetDB&".dbo.CuDTGeneric(iBaseDSD,iCtUnit,fCTUPublic,xNewWindow,iDept,"
			sqlValue = ") VALUES("&xiBaseDSD&","&xCtUnitID&",'Y','N','0'," 
			if not isNull(RSL("seq")) and trim(RSL("seq"))<>"" then
				sql = sql & "xImportant,"
				sqlValue = sqlValue & pkstr(trim(RSL("seq")),",")
			end if
			if not isNull(RSL("title")) and trim(RSL("title"))<>"" then
				sql = sql & "sTitle,"
				sqlValue = sqlValue & pkstr(trim(RSL("title")),",")
			end if
			if not isNull(RSL("content")) and trim(RSL("content"))<>"" then
				sql = sql & "xBody,"
				sqlValue = sqlValue & pkstr(trim(RSL("content")),",")
			end if
			if not isNull(RSL("modifyuser")) and trim(RSL("modifyuser"))<>"" then
				sql = sql & "iEditor,"
				sqlValue = sqlValue & pkstr(trim(RSL("modifyuser")),",")
			end if
			if not isNull(RSL("postdate")) and trim(RSL("postdate"))<>"" then
				sql = sql & "xPostDate,"
				sqlValue = sqlValue & pkstr(xStdTime2(RSL("postdate")),",")
			end if
			if not isNull(RSL("createtime")) and trim(RSL("createtime"))<>"" then
				sql = sql & "createdDate,"
				sqlValue = sqlValue & pkstr(xStdTime2(RSL("createtime")),",")
			end if
			if not isNull(RSL("modifytime")) and trim(RSL("modifytime"))<>"" then
				sql = sql & "dEditDate,"
				sqlValue = sqlValue & pkstr(xStdTime2(RSL("modifytime")),",")
			end if
		    sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
			sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID;"
			response.write sql & "<br>"
			xfout.writeline sql
	    	set RSx = conn2.Execute(SQL)
	    	xNewIdentity = RSx(0)  
    		'----新增Slave table處理  
			sql = "INSERT INTO  "&targetDB&".dbo." & xtableName & "(giCuItem,"
			sqlValue = ") Values(" & xNewIdentity & ","	 
			sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"	
			response.write sql & "<br>"
			xfout.writeline sql		
			conn2.Execute(sql)
			'----寫入GIPDataTranDetail表		
			SQLGD = "Insert Into GIPDataTranDetail values("&xTID&","&xNewIdentity&")"
			conn2.execute(SQLGD)
			tranMessage="lawid="&cStr(RSL("lawid"))
			if err.number<>0 then
	    		conn2.rollbacktrans
	    		response.write "["+tranMessage+"] 資料發生轉檔錯誤!<br>發生原因:"+err.description+"<hr>"
	    		xfout.writeline "["+tranMessage+"] 資料發生轉檔錯誤!"
	    		xfout.writeline "發生原因:"+err.description
	   			TranFail = TranFail + 1
	    		err.number=0
	    	else
	    		conn2.committrans
	    		'----law圖檔附件處理
				Set Conn3 = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'				Conn3.Open session("ODBCDSN")	  
'Set Conn3 = Server.CreateObject("HyWebDB3.dbExecute")
Conn3.ConnectionString = session("ODBCDSN")
Conn3.ConnectionTimeout=0
Conn3.CursorLocation = 3
Conn3.open
'----------HyWeb GIP DB CONNECTION PATCH----------

				SQLAttach="Select title aTitle, title aDesc,docfile_realname OFileName,docfile NFileName, substring('00'+convert(varchar(3),seq),Len('00'+convert(varchar(3),seq))-2,3) seq from [Webey].dbo.law_detail where lawid=" & RSL("lawid") & _
					"union " & _
					"Select title aTitle, title aDesc,pdffile_realname OFileName,pdffile NFileName,substring('00'+convert(varchar(3),seq),Len('00'+convert(varchar(3),seq))-2,3) seq from [Webey].dbo.law_detail where lawid=" & RSL("lawid") & _
					"union " & _
					"Select title aTitle, title aDesc,docfile_realname OFileName,docfile NFileName, '00' seq from [Webey].dbo.law where lawid="&RSL("lawid")&" and docfile is not null and rtrim(docfile)<>'' " & _
					"union " & _
					"Select title aTitle, title aDesc,pdffile_realname OFileName,pdffile NFileName, '00' seq  from [Webey].dbo.law where lawid="&RSL("lawid")&" and pdffile is not null and rtrim(pdffile)<>'' " & _
					"order by seq desc"
				Set RSA = conn3.execute(SQLAttach)
				if not RSA.eof then
					while not RSA.eof
			    		SQLAInsert = ""
			    		NFileNameStr = "null"
			    		if not isNull(RSA("NFileName")) and trim(RSA("NFileName"))<>"" then NFileNameStr="'"&trim(RSA("NFileName"))&"'"
						SQLAInsert = "Insert Into CuDTAttach (xiCuItem,aTitle,aDesc,OFileName,NFileName,aEditor,aEditDate,bList,listSeq) " & _
								"values(" & xNewIdentity & "," & _
								"'" & trim(RSA("aTitle")) & "'," & _
								"'" & trim(RSA("aDesc")) & "',null," _
								& NFileNameStr & "," & _
								"'" & session("UserName") & "'," & _
								"'" & date() & "'," & _
								"'Y'," & _
								"'" & RSA("seq") & "');" 
						if not isNull(RSA("NFileName")) and trim(RSA("NFileName"))<>"" then
							SQLAInsert = SQLAInsert & "Insert Into imageFile(newFileName, oldFileName) VALUES(" & _
								"'" & trim(RSA("NFileName")) & "'," & _
								"'" & trim(RSA("OFileName")) & "');"   	
						end if	
						conn3.execute(SQLAInsert)		
			    		xNFileName = trim(RSA("NFileName"))
						if xNFileName <> "" and not isNull(xNFileName) then 
							if fso.FileExists(sourcePath+xNFileName) then _
								fso.CopyFile sourcePath+xNFileName,targetPath2+xNFileName
						end if	
						RSA.movenext
					wend
				end if
'				Conn3.close
				set Conn3 = Nothing
	    		response.write "["+tranMessage+"] 資料轉檔完成!<hr>"
	    		xfout.writeline "["+tranMessage+"] 資料轉檔完成!"
	    		TranSuccess = TranSuccess + 1
	    	end if	    
			xfout.writeline "----------------------------------------------------------------------------"
			Conn2.close
			set Conn2 = Nothing
			RSL.movenext
		wend
	end if
	SQLU = "Update GIPDataTran Set XMLSuccess="&TranSuccess&",XMLFail="&TranFail&" where TID="&xTID
	conn.execute(SQLU)
'	Conn.close
	set Conn = Nothing	
xfout.writeline "==============================================================================="
response.write "<hr>轉檔 "&recordCount&" 筆, 共成功 "&TranSuccess&" 筆!"
xfout.writeline "轉檔 "&recordCount&" 筆, 共成功 "&TranSuccess&" 筆!"
response.end
%>
</html>
