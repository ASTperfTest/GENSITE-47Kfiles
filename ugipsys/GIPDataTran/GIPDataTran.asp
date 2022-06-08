<%@ CodePage = 65001 %>
<% Response.Expires = 0%>
<!--#include virtual = "/inc/dbutil.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
</head>
<%
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
        xyear = cstr(year(dt))
        xmonth = right("00"+ cstr(month(dt)),2)
        xday = right("00"+ cstr(day(dt)),2)
        xhour = right("00" + cstr(hour(dt)),2)
        xminute = right("00" + cstr(minute(dt)),2)
        xsecond = right("00" + cstr(second(dt)),2)
        xStdTime2 = xyear & "/" & xmonth & "/" & xday & " " & xhour & ":" & xminute & ":" & xsecond
   end if
end function

function xStdTime(dt)    
   if Len(dt)=0 or isnull(dt) then
     	xStdTime=""
   else
        xyear = cstr(year(dt))
        xmonth = right("00"+ cstr(month(dt)),2)
        xday = right("00"+ cstr(day(dt)),2)
        xStdTime = xyear & "/" & xmonth & "/" & xday
   end if
end function


	HTUploadPath=session("Public")+"data/"
	targetPath=server.mappath(HTUploadPath) & "\"		'----主圖與檔案下載式檔案實體路徑
	HTUploadPath2=session("Public")+"Attachment/"
	targetPath2=server.mappath(HTUploadPath2) & "\"		'----附件檔案實體路徑
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
	'----取得轉檔參數xml
	set tranDom = Server.CreateObject("MICROSOFT.XMLDOM")
	tranDom.async = false
	tranDom.setProperty("ServerHTTPRequest") = true	
	LoadXML = server.MapPath("XML/DataTran"&xCtUnitID&".xml")
	xv = tranDom.load(LoadXML)
	if tranDom.parseError.reason <> "" then 
		Response.Write("tranDom parseError on line " &  tranDom.parseError.line)
		Response.Write("<BR>Reason: " &  tranDom.parseError.reason)
		Response.End()
	end if
	Set CtUintTranListNode = tranDom.selectSingleNode("DataTran/CtUintTranList")
	Set updateFieldListNode = tranDom.selectSingleNode("DataTran/updateFieldList")
	targetDB = nullText(tranDom.selectSingleNode("DataTran/targetDB"))	
	sourceDB = nullText(tranDom.selectSingleNode("DataTran/sourceDB"))
	sourceTable = nullText(tranDom.selectSingleNode("DataTran/sourceTable"))	
	'----串來源SQL字串取得來源recordsets
	SQLS = "Select * from ["&sourceDB&"].dbo."&sourceTable&" "
	if nullText(tranDom.selectSingleNode("DataTran/whereCondition"))<> "" then _
		SQLS = SQLS & " " & nullText(tranDom.selectSingleNode("DataTran/whereCondition"))
	if nullText(tranDom.selectSingleNode("DataTran/orderby"))<> "" then _
		SQLS = SQLS & " " & nullText(tranDom.selectSingleNode("DataTran/orderby"))
	SET RSS = conn.execute(SQLS)
	if RSS.Eof then
		response.write "來源table無資料!"
		response.end
	end if
	'----開檔案
	outFile = server.MapPath("LOG/DataTranLog_"&xCtUnitID&".txt")
	Set fso = CreateObject("scripting.filesystemobject")
	set xfout = fso.createTextFile(outFile, true)	
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
	'----刪除主圖/檔案下載式檔案
    if request.querystring("DYN")="" then
	SQLDI = "Select C.xImgFile,C.fileDownLoad,C.iCuItem " & _
			"from GIPDataTran G " & _
			"    Left Join GIPDataTranDetail GD ON G.TID=GD.TID " & _
			"    Left Join CuDTGeneric C ON GD.iCuItem=C.iCuItem " & _
			"where C.iCuItem is not null and XMLCtUnitID=" & xCtUnitID & " AND XMLmySiteID='"&session("mySiteID")&"'"
	Set RSI = conn.execute(SQLDI)
	if not RSI.eof then
		while not RSI.eof
			if not isNull(RSI("xImgFile")) then 
				if fso.FileExists(targetPath+RSI("xImgFile")) then fso.DeleteFile(targetPath+RSI("xImgFile"))
			end if
			if not isNull(RSI("fileDownLoad")) then 
				if fso.FileExists(targetPath+RSI("fileDownLoad")) then 
					fso.DeleteFile(targetPath+RSI("fileDownLoad"))
					conn.execute("Delete I from CuDTGeneric C Left Join ImageFile I ON C.fileDownLoad=I.NewFileName where C.iCuItem=" & RSI("iCuItem"))
				end if
			end if
			RSI.movenext
		wend
	end if
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
	'----刪除GIPDataTran表
	SQLDGM = "Delete GIPDataTran where XMLCtUnitID=" & xCtUnitID & " AND XMLmySiteID='"&session("mySiteID")&"'"
	conn.execute(SQLDGM)

	response.write "清除前次轉檔資料完成!<hr>"
    end if
'	response.end	
'----4.Loop 來源recordsets,新增資料
	'----寫入GIPDataTran表		
	SQLG = "set nocount on;Insert Into GIPDataTran values(null,'"&session("UserName")&"','"&date()&"','"&session("mySiteID")&"',"&xCtUnitID&",null,null) select @@IDENTITY as NewID;"
	SET RSG = conn.execute(SQLG)
	xTID = RSG(0)
	recordCount=0
	TranSuccess=0
	TranFail=0
	while not RSS.Eof 
		recordCount=recordCount+1
		err.number=0
		on error resume next
		'----主圖名稱
		xImgFileValue = ""
	    	if nullText(CtUintTranListNode.selectSingleNode("CtUintTran[targetField='ximgFile']"))<>"" then
	    		xImgFile_SourceField = nullText(CtUintTranListNode.selectSingleNode("CtUintTran[targetField='ximgFile']/sourceField"))
			xImgFileValue = RSS(xImgFile_SourceField)
	    	end if
		'----檔案下載式檔案名稱
		xfileDownLoadValue = ""
		xfileDownLoadValue_Org = ""		
	    	if nullText(CtUintTranListNode.selectSingleNode("CtUintTran[targetField='fileDownLoad']"))<>"" then
	    		xfileDownLoad_SourceField = nullText(CtUintTranListNode.selectSingleNode("CtUintTran[targetField='fileDownLoad']/sourceField"))
			xfileDownLoadValue = RSS(xfileDownLoad_SourceField)
	    		xfileDownLoad_SourceField_Org = nullText(CtUintTranListNode.selectSingleNode("CtUintTran[targetField='fileDownLoad']/sourceField_Org"))
			xfileDownLoadValue_Org = RSS(xfileDownLoad_SourceField_Org)
	    	end if
	    	Set Conn2 = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'		Conn2.Open session("ODBCDSN")
Set Conn2 = Server.CreateObject("HyWebDB3.dbExecute")
Conn2.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------

		conn2.begintrans
	    '----新增CuDTGeneric處理
		sql = "INSERT INTO  "&targetDB&".dbo.CuDTGeneric(iBaseDSD,iCtUnit,fCTUPublic,xNewWindow,iDept,showType,"
		sqlValue = ") VALUES("&xiBaseDSD&","&xCtUnitID&",'Y','N','0','1'," 
		for each fieldNode in htPageDom.selectNodes("//DataSchemaDef/dsTable[tableName='CuDtGeneric']/fieldList/field") 
			if nullText(CtUintTranListNode.selectSingleNode("CtUintTran[targetField='"&fieldNode.selectSingleNode("fieldName").text&"']"))<>"" then
			    Set CtUintTranNode = CtUintTranListNode.selectSingleNode("CtUintTran[targetField='"&fieldNode.selectSingleNode("fieldName").text&"']")			    
				sql = sql & fieldNode.selectSingleNode("fieldName").text & ","
				if fieldNode.selectSingleNode("fieldName").text="createdDate" or fieldNode.selectSingleNode("fieldName").text="deditDate" or fieldNode.selectSingleNode("fieldName").text="avBegin" or fieldNode.selectSingleNode("fieldName").text="avEnd" then
				    if not isNull(RSS(CtUintTranNode.selectSingleNode("sourceField").text)) then
						sqlValue = sqlValue & pkstr(xStdTime2(RSS(CtUintTranNode.selectSingleNode("sourceField").text)),",")
				    else
				        sqlValue = sqlValue & "null,"
				    end if
				elseif fieldNode.selectSingleNode("fieldName").text="xpostDate" then
				    if not isNull(RSS(CtUintTranNode.selectSingleNode("sourceField").text)) then
						sqlValue = sqlValue & pkstr(xStdTime(RSS(CtUintTranNode.selectSingleNode("sourceField").text)),",")
				    else
				        sqlValue = sqlValue & "null,"
				    end if
				else
					sqlValue = sqlValue & pkstr(trim(RSS(CtUintTranNode.selectSingleNode("sourceField").text)),",")
				end if			    
		    end if
		next
	    	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
		sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID;"
		response.write sql & "<br>"
		xfout.writeline sql
    		set RSx = conn2.Execute(SQL)
    		xNewIdentity = RSx(0)  
    	'----新增Slave table處理  
		sql = "INSERT INTO  "&targetDB&".dbo." & xtableName & "(giCuItem,"
		sqlValue = ") Values(" & xNewIdentity & ","	 
		for each fieldNode in htPageDom.selectNodes("//DataSchemaDef/dsTable[tableName='"&xtableName&"']/fieldList/field") 
			if nullText(CtUintTranListNode.selectSingleNode("CtUintTran[targetField='"&fieldNode.selectSingleNode("fieldName").text&"']"))<>"" then
				Set CtUintTranNode = CtUintTranListNode.selectSingleNode("CtUintTran[targetField='"&fieldNode.selectSingleNode("fieldName").text&"']")
				sql = sql & fieldNode.selectSingleNode("fieldName").text & ","
				sqlValue = sqlValue & trim(pkstr(RSS(CtUintTranNode.selectSingleNode("sourceField").text),","))
		    end if
		next
		sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"	
		response.write sql & "<br>"
		xfout.writeline sql		
		conn2.Execute(sql)
		'----寫入GIPDataTranDetail表		
		SQLGD = "Insert Into GIPDataTranDetail values("&xTID&","&xNewIdentity&")"
		conn2.execute(SQLGD)
		tranMessage=""
		for each fieldNode in updateFieldListNode.selectNodes("updateField") 
			tranMessage= tranMessage & "htx." & fieldNode.text & " = '"&RSS(fieldNode.text)&"' and "
		next
		tranMessage = left(tranMessage,len(tranMessage)-5)
    	if err.number<>0 then
    		conn2.rollbacktrans
    		response.write "["+tranMessage+"] 資料發生轉檔錯誤!<br>發生原因:"+err.description+"<hr>"
    		xfout.writeline "["+tranMessage+"] 資料發生轉檔錯誤!"
    		xfout.writeline "發生原因:"+err.description
   			TranFail = TranFail + 1
    		err.number=0
    	else
    		conn2.committrans
    		'----主圖檔案處理	
    		if xImgFileValue <> "" and fso.FileExists(sourcePath+xImgFileValue) then _
    			fso.CopyFile sourcePath+xImgFileValue,targetPath+xImgFileValue
    		'----檔案下載式檔案處理	
    		if xfileDownLoadValue <> "" and fso.FileExists(sourcePath+xfileDownLoadValue) then 
    			fso.CopyFile sourcePath+xfileDownLoadValue,targetPath+xfileDownLoadValue
    			conn.execute("Insert Into ImageFile values('"&xfileDownLoadValue&"','"&xfileDownLoadValue_Org&"')")		
    		end if
    		'----附件檔案處理
			'----串附件SQL
			Set Conn3 = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'			Conn3.Open session("ODBCDSN")
Set Conn3 = Server.CreateObject("HyWebDB3.dbExecute")
Conn3.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------

			attachTable = nullText(tranDom.selectSingleNode("DataTran/attachment/attachmentTable"))	
			if attachTable <> "" then 
				FKNodeStr = ""
				Set attachTranNodeList = tranDom.selectNodes("DataTran/attachment/attachmentTran")
				for each attachTranNode in attachTranNodeList
				    SQLA = ""
					SQLA = "Select ghtx."&nullText(attachTranNode.selectSingleNode("aTitleField"))&" aTitle, ghtx." _
							& nullText(attachTranNode.selectSingleNode("aDescField"))&" aDesc, ghtx." _
							& nullText(attachTranNode.selectSingleNode("NFileNameField"))&" NFileName " & _
							" from ["&sourceDB&"].dbo."&sourceTable&" as htx " & _
							" Left Join ["&sourceDB&"].dbo."&attachTable&" ghtx ON "
					Set FKNodeList = tranDom.selectNodes("DataTran/attachment/attachmentKeyList/attachmentKey")
					for each FKNode in FKNodeList
						FKNodeStr = FKNodeStr & " AND htx."&nullText(FKNode.selectSingleNode("PKField"))&"=ghtx."&nullText(FKNode.selectSingleNode("FKField"))&" "
					next
					SQLA = SQLA & mid(FKNodeStr,6)
					SQLA = SQLA & " where " & tranMessage
					if nullText(tranDom.selectSingleNode("DataTran/attachment/orderby"))<> "" then _
						SQLA = SQLA & " " & nullText(tranDom.selectSingleNode("DataTran/attachment/orderby"))	
		    		if SQLA <> "" then
		    			Set RSA = conn3.execute(SQLA)
		    			if not RSA.eof then
		    				listSeq = 0
		    				while not RSA.eof 
		    				    SQLAInsert = ""
		    				    listSeq = listSeq + 1
		    				    listSeqStr = right("00" + cStr(listSeq),2)
		    				    if (not isNull(RSA("aTitle")) and trim(RSA("aTitle"))<>"") and (not isNull(RSA("NFileName")) and trim(RSA("NFileName"))<>"") then
									SQLAInsert = SQLAInsert & "Insert Into CuDTAttach (xiCuItem,aTitle,aDesc,OFileName,NFileName,aEditor,aEditDate,bList,listSeq) " & _
											"values(" & xNewIdentity & "," & _
											"'" & trim(RSA("aTitle")) & "'," & _
											"'" & trim(RSA("aDesc")) & "',null," & _
											"'" & trim(RSA("NFileName")) & "'," & _
											"'" & session("UserName") & "'," & _
											"'" & date() & "'," & _
											"'Y'," & _
											"'" & listSeqStr & "');" & _
											"Insert Into imageFile(newFileName, oldFileName) VALUES(" & _
											"'" & trim(RSA("NFileName")) & "'," & _
											"'" & trim(RSA("aTitle")) & "');"   
'					response.write SQLAInsert & "<br>"
									conn3.execute(SQLAInsert)		
						    		xNFileName = trim(RSA("NFileName"))
									if xNFileName <> "" and not isNull(xNFileName) then 
										if fso.FileExists(sourcePath+xNFileName) then _
											fso.CopyFile sourcePath+xNFileName,targetPath2+xNFileName
									end if	
								end if							 					
		    					RSA.movenext
		    				wend
		    			end if
					end if	    		
				next
			end if
'			response.write SQLA & "<br>"
'			response.end    		
'			Conn3.close
			set Conn3 = Nothing
    		response.write "["+tranMessage+"] 資料轉檔完成!<hr>"
    		xfout.writeline "["+tranMessage+"] 資料轉檔完成!"
    		TranSuccess = TranSuccess + 1
    	end if	    
		xfout.writeline "----------------------------------------------------------------------------"
	    RSS.movenext
		Conn2.close
		set Conn2 = Nothing
	wend
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
