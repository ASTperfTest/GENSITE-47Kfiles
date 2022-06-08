<%@ codepage=65001 %>
<% Response.Expires = 0
HTProgCode="GW1M95"
HTProgPrefix="GIPDataImport" %>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
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
end function

	HTUploadPath=session("Public")+"data/"
	targetPath=server.mappath(HTUploadPath) & "\"	
	HTUploadPath2=session("Public")+"Attachment/"
	targetPath2=server.mappath(HTUploadPath2) & "\"	
	sourcePath=server.mappath("GIPDataXML/Data") & "\"
	sourcePath2=server.mappath("GIPDataXML/Attachment") & "\"
	INXMLPath = session("public")+"GIPDataXML/INXML"
	LogPath = session("public")+"GIPDataXML/Log"
'response.write sourcePath & "<br>"
'response.write sourcePath2 & "<br>"
'response.end
	'----處理本站DTD或DSD XML
	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true		
    '----找出對應的DSD,若不存在則用DTD
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & request("htx_CtUnitID") & ".xml")
	if fso.FileExists(filePath) then
		LoadXML = filePath
	else
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & request("htx_iBaseDSD") & ".xml")
	end if   
	xv = htPageDom.load(LoadXML)	
  	if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reason: " &  htPageDom.parseError.reason)
    		Response.End()
  	end if
	set refModel = htPageDom.selectSingleNode("//dsTable")
	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")    
	sBaseTableName = nullText(refModel.selectSingleNode("tableName")) 	
'	response.write "<XMP>"+allModel.xml+"</XMP>"	

'----若匯入方式為overwrite,則先刪除
	if request("htx_ImportWay") = "overwrite" then			
		'----刪除主圖/檔案下載式檔案
		SQLDI = "Select C.xImgFile,C.fileDownLoad,C.iCuItem from GIPDataImport G " & _
				"    Left Join GIPDataImportDetail GD ON G.TID=GD.TID " & _
				"    Left Join CuDTGeneric C ON GD.iCuItem=C.iCuItem " & _
				"where XMLCtUnitID=" & request("htx_CtUnitID") & " AND XMLFile='"&request("htx_XMLFile")&"' and (C.xImgFile is not null or C.fileDownLoad is not null)"
		Set RSI = conn.execute(SQLDI)
		if not RSI.eof then
			while not RSI.eof
				if not isNull(RSI("xImgFile")) then
					if fso.FileExists(targetPath+RSI("xImgFile")) then 
						fso.DeleteFile(targetPath+RSI("xImgFile"))
					end if
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
		SQLDA = "Select C.NFileName,C.xiCuItem from GIPDataImport G " & _
				"    Left Join GIPDataImportDetail GD ON G.TID=GD.TID " & _
				"    Left Join CuDTAttach C ON GD.iCuItem=C.xiCuItem " & _
				"where XMLCtUnitID=" & request("htx_CtUnitID") & " AND XMLFile='"&request("htx_XMLFile")&"' and C.xiCuItem is not null"
		Set RSI = conn.execute(SQLDA)
		if not RSI.eof then
			while not RSI.eof
				if not isNull(RSI("NFileName")) then
					if fso.FileExists(targetPath2+RSI("NFileName")) then 
						fso.DeleteFile(targetPath2+RSI("NFileName"))
						conn.execute("Delete I from CuDTAttach C Left Join ImageFile I ON C.NFileName=I.NewFileName where C.xiCuItem=" & RSI("xiCuItem"))
					end if
				end if
				RSI.movenext
			wend
		end if
		
		'----刪除ImageFile
		SQLDIF = "Delete I from GIPDataImport G " & _
				"    Left Join GIPDataImportDetail GD ON G.TID=GD.TID " & _
				"    Left Join CuDTAttach C ON GD.iCuItem=C.xiCuItem " & _
				"    Left Join ImageFile I ON C.NFileName=I.NewFileName " & _				
				"where XMLCtUnitID=" & request("htx_CtUnitID") & " AND XMLFile='"&request("htx_XMLFile")&"'"
		conn.execute(SQLDIF)			
		'----刪除CuDTAttach
		SQLDAF = "Delete C from GIPDataImport G " & _
				"    Left Join GIPDataImportDetail GD ON G.TID=GD.TID " & _
				"    Left Join CuDTAttach C ON GD.iCuItem=C.xiCuItem " & _
				"where XMLCtUnitID=" & request("htx_CtUnitID") & " AND XMLFile='"&request("htx_XMLFile")&"'"
		conn.execute(SQLDAF)		
		'----刪除Detail表
		SQLDL = "Delete C from GIPDataImport G " & _
				"    Left Join GIPDataImportDetail GD ON G.TID=GD.TID " & _
				"    Left Join "&sBaseTableName&" C ON GD.iCuItem=C.giCuItem " & _
				"where XMLCtUnitID=" & request("htx_CtUnitID") & " AND XMLFile='"&request("htx_XMLFile")&"'"
		conn.execute(SQLDL)
		'----刪除CuDTGeneric表
		SQLDM = "Delete C from GIPDataImport G " & _
				"    Left Join GIPDataImportDetail GD ON G.TID=GD.TID " & _
				"    Left Join CuDTGeneric C ON GD.iCuItem=C.iCuItem " & _
				"where XMLCtUnitID=" & request("htx_CtUnitID") & " AND XMLFile='"&request("htx_XMLFile")&"'"
		conn.execute(SQLDM)		
		'----刪除GIPDataImportDetail表
		SQLDGL = "Delete GD from GIPDataImport G " & _
				"    Left Join GIPDataImportDetail GD ON G.TID=GD.TID " & _
				"where XMLCtUnitID=" & request("htx_CtUnitID") & " AND XMLFile='"&request("htx_XMLFile")&"'"
		conn.execute(SQLDGL)
		'----刪除GIPDataImport表
		SQLDGM = "Delete GIPDataImport where XMLCtUnitID=" & request("htx_CtUnitID") & " AND XMLFile='"&request("htx_XMLFile")&"'"
		conn.execute(SQLDGM)
	end if
'response.end
'----新增匯入資料
	Set xfout = fso.CreateTextFile(server.mappath(LogPath) & "\" & request("htx_XMLFile") & ".log")
	xfout.writeline "匯入記錄檔" + cStr(now())
	xfout.writeline "---------------------------------------------------------"	
    XMLCount = 0 : XMLSuccess = 0 : XMLFail = 0
	'----寫入GIPDataImport表		
	SQLG = "set nocount on;Insert Into GIPDataImport values('"&request("htx_XMLFile")&"','"&session("UserName")&"','"&date()&"',"&request("htx_CtUnitID")&",null,null) select @@IDENTITY as NewID;"
	SET RSG = conn.execute(SQLG)
	xTID = RSG(0)
	'----Load匯入XML檔
	set GIPXMLDom = Server.CreateObject("MICROSOFT.XMLDOM")
	GIPXMLDom.async = false
	GIPXMLDom.setProperty("ServerHTTPRequest") = true
	xv = GIPXMLDom.load(server.mappath(INXMLPath) & "\" & request("htx_XMLFile"))	
  	if GIPXMLDom.parseError.reason <> "" then 
		Response.Write("GIPXMLDom parseError on line " &  GIPXMLDom.parseError.line)
		Response.Write("<BR>Reason: " &  GIPXMLDom.parseError.reason)
		Response.End()
  	end if
	Set GIPXMLNode = GIPXMLDom.selectNodes("GIPDataXML/GIPData/fieldList")	
	for each XMLNode in GIPXMLNode
		XMLCount = XMLCount + 1
		err.number=0
		tranMessage = ""
		on error resume next
		'----新增CuDTGeneric表
		conn.begintrans
		sql = ""
		sql = "INSERT INTO CuDTGeneric(iBaseDSD,iCtUnit,"
		sqlValue = ") VALUES("&request("htx_iBaseDSD")&","&request("htx_CtUnitID")&","
		for each param in allModel.selectNodes("dsTable[tableName='CuDtGeneric']/fieldList/field[fieldName!='icuitem' and fieldName!='ibaseDsd' and fieldName!='ictunit']") 
			if nullText(XMLNode.selectSingleNode("field[fieldName='"&nullText(param.selectSingleNode("fieldName"))&"']/fieldValue")) <> "" then
				sql = sql & nullText(param.selectSingleNode("fieldName")) & ","
				sqlValue = sqlValue & "'" & nullText(XMLNode.selectSingleNode("field[fieldName='"&nullText(param.selectSingleNode("fieldName"))&"']/fieldValue")) & "',"		
			end if
			if nullText(param.selectSingleNode("fieldName")) = "stitle" then 
				tranMessage = nullText(XMLNode.selectSingleNode("field[fieldName='"&nullText(param.selectSingleNode("fieldName"))&"']/fieldValue"))
			end if
		next  
		sqlm = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
		sqlm = "set nocount on;"&sqlm&"; select @@IDENTITY as NewID"	
		set RSx = conn.Execute(sqlm)
		xNewIdentity = RSx(0)	
		'----新增Slave表
		sql = "INSERT INTO  " & sBaseTableName & "(giCuItem,"
		sqlValue = ") Values(" & xNewIdentity & ","
		for each param in refModel.selectNodes("fieldList/field") 
			if nullText(XMLNode.selectSingleNode("field[fieldName='"&nullText(param.selectSingleNode("fieldName"))&"']/fieldValue"))<>"" then
				sql = sql & nullText(param.selectSingleNode("fieldName")) & ","
				sqlValue = sqlValue & "'" & nullText(XMLNode.selectSingleNode("field[fieldName='"&nullText(param.selectSingleNode("fieldName"))&"']/fieldValue")) & "',"		
			end if
		next
		sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"		
		conn.Execute(sql)			
		'----寫入GIPDataImportDetail表		
		SQLGD = "Insert Into GIPDataImportDetail values("&xTID&","&xNewIdentity&")"
		conn.execute(SQLGD)

    	if err.number<>0 then
    		conn.rollbacktrans
    		response.write "[ "+tranMessage+" ] 此筆資料匯入錯誤!<br>錯誤原因:"+err.description+"<br>"
    		response.write sqlm & "<br>"
    		response.write sql & "<hr>"
    		xfout.writeline "[ "+tranMessage+" ] 此筆資料匯入錯誤!"
    		xfout.writeline "錯誤原因:"+err.description
     		xfout.writeline sqlm
    		xfout.writeline sql
    		xfout.writeline "---------------------------------------------------------"
   			XMLFail = XMLFail + 1
    		err.number=0
    	else
    		conn.committrans
    		'----主圖與檔案下載式檔案處理
    		xImgFileValue = nullText(XMLNode.selectSingleNode("field[fieldName='ximgFile']/fieldValue"))
			if xImgFileValue <> "" and fso.FileExists(sourcePath+xImgFileValue) then
				fso.CopyFile sourcePath+xImgFileValue,targetPath+xImgFileValue
			end if
    		fileDownLoadValue = nullText(XMLNode.selectSingleNode("field[fieldName='fileDownLoad']/fieldValue"))
			if fileDownLoadValue <> "" and fso.FileExists(sourcePath+fileDownLoadValue) then
				fso.CopyFile sourcePath+fileDownLoadValue,targetPath+fileDownLoadValue	
				conn.execute("Insert Into ImageFile values('"&fileDownLoadValue&"','"&fileDownLoadValue&"')")		
			end if    		
    		'----附件檔案處理
			SQLAInsert = ""
			for each AttachNode in XMLNode.selectNodes("attachList/Attach")
				SQLAInsert = SQLAInsert & "Insert Into CuDTAttach (xiCuItem,aTitle,aDesc,OFileName,NFileName,aEditor,aEditDate,bList,listSeq) " & _
						"values(" & xNewIdentity & "," & _
						"'" & nullText(AttachNode.selectSingleNode("aTitle")) & "'," & _
						"'" & nullText(AttachNode.selectSingleNode("aDesc")) & "',null," & _
						"'" & nullText(AttachNode.selectSingleNode("NFileName")) & "'," & _
						"'" & session("UserName") & "'," & _
						"'" & date() & "'," & _
						"'" & nullText(AttachNode.selectSingleNode("bList")) & "'," & _
						"'" & nullText(AttachNode.selectSingleNode("listSeq")) & "');" & _
						"Insert Into imageFile(newFileName, oldFileName) VALUES(" & _
						"'" & nullText(AttachNode.selectSingleNode("NFileName")) & "'," & _
						"'" & nullText(AttachNode.selectSingleNode("OldFileName")) & "');"
						
			    		xNFileName = nullText(AttachNode.selectSingleNode("NFileName"))
						if xNFileName <> "" and fso.FileExists(sourcePath2+xNFileName) then
							fso.CopyFile sourcePath2+xNFileName,targetPath2+xNFileName
						end if
			next
			if SQLAInsert<>"" then conn.execute(SQLAInsert)
    		
    		response.write "[ "+tranMessage+" ] 此筆資料匯入OK!<hr>"
    		xfout.writeline "[ "+tranMessage+" ] 此筆資料匯入OK!"
    		xfout.writeline "---------------------------------------------------------"
    		XMLSuccess = XMLSuccess + 1
    	end if  		
	next
	SQLU = "Update GIPDataImport Set XMLSuccess="&XMLSuccess&",XMLFail="&XMLFail&" where TID="&xTID
	conn.execute(SQLU)	
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
共完成 <%=XMLCount%> 筆資料匯入! 成功 <%=XMLSuccess%> 筆, 失敗 <%=XMLFail%> 筆!<br/>
<A href="GIPDataImport.asp">回新增匯入</A>
</body>
</html>  	
