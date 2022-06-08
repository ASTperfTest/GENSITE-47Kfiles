<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GC1AP1" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
</head>
<%
'----950214繁轉簡自動產生機制
	Set oConvert = CreateObject("Convertsim.Convert")
	'----Load sysPara設定檔
	set convertsimDom = Server.CreateObject("MICROSOFT.XMLDOM")
	convertsimDom.async = false
	convertsimDom.setProperty("ServerHTTPRequest") = true		
	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysPara.xml"
	xv = convertsimDom.load(LoadXML)
	if convertsimDom.parseError.reason <> "" then 
		Response.Write("convertsimDom parseError on line " &  convertsimDom.parseError.line)
		Response.Write("<BR>Reason: " &  convertsimDom.parseError.reason)
		Response.End()
	end if		
	'----Load 自動轉碼對象unit
	Set convertsimNodes = convertsimDom.selectNodes("//convertsimList/convertsimUnit[convertsimFrom='"&session("CtUnitID")&"']/convertsimTo")
	if convertsimNodes.length > 0 then
		'----取得from unit的xml	
		set htPageDom = session("codeXMLSpec")
		set refModel = htPageDom.selectSingleNode("//dsTable")
		set allModel = htPageDom.selectSingleNode("//DataSchemaDef")		
		'----取得轉碼資料
		SQLCDG = "Select CDG.*,u.checkYn from CuDtGeneric CDG " & _
			"Left Join CtUnit AS u ON CDG.ictunit=u.ctUnitId " & _
			"where CDG.icuitem=" & pkStr(session("convertsim_icuitem"),"")
		set RSreg = conn.execute(SQLCDG)
		xSlaveTable = refModel.selectSingleNode("tableName").text
		SQLSlave = "Select * from "&xSlaveTable&" where gicuitem = " & session("convertsim_icuitem")
		set RSSlave = conn.execute(SQLSlave)
		'---Loop處理各轉碼對象unit
		for each convertsimNode in convertsimNodes
			CtUnitIDTo=nullText(convertsimNode.selectSingleNode("."))
			sql = "SELECT u.*,b.sbaseTableName FROM CtUnit AS u " _
				& " Left Join BaseDsd As b ON u.ibaseDsd=b.ibaseDsd" _
				& " WHERE u.CtUnitID=" & pkStr(CtUnitIDTo,"")	
			set RS = Conn.execute(sql)	
			if not RS.eof then
				xctUnitId = RS("ctUnitId")
				xibaseDsd = RS("ibaseDsd")
				xcheckYn = RS("checkYn")
				if isNull(RS("sbaseTableName")) then
					xsbaseTableName = "CuDTx" & xibaseDsd
				else
					xsbaseTableName = RS("sbaseTableName")
				end if
				if xcheckYn=RSreg("checkYn") then	'----主題單元同為須審核或同為不須審核
					xfCTUPublic=RSreg("fCTUPublic")
				elseif RSreg("checkYn")="Y" and xcheckYn="N" then	'----Parent須審核但child不須審核
					xfCTUPublic="Y"
				elseif RSreg("checkYn")="N" and xcheckYn="Y" then	'----Parent不須審核但child須審核
					xfCTUPublic="P"
				end if						
				'----Master table處理
				sql = "INSERT INTO  CuDtGeneric(fctupublic,ibaseDsd,ictunit,showType,refId,deditDate,createdDate,"
				sqlValue = ") VALUES('"&xfCTUPublic&"',"&xibaseDsd&","&xctUnitId&",'4',"&session("convertsim_icuitem")&",getdate(),getdate()," 
				for each param in allModel.selectNodes("dsTable[tableName='CuDtGeneric']/fieldList/field[fieldName!='icuitem' and fieldName!='ibaseDsd' and fieldName!='ictunit' and fieldName!='fctupublic' and fieldName!='showType' and fieldName!='refID' and fieldName!='deditDate' and fieldName!='createdDate']") 
					xdataType = nullText(param.selectSingleNode("dataType"))
					if not isNull(RSreg(param.selectSingleNode("fieldName").text)) then
						sql = sql & param.selectSingleNode("fieldName").text & ","
						if xdataType = "char" or xdataType = "varchar" or xdataType = "text" or xdataType = "nchar" or xdataType = "nvarchar" or xdataType = "ntext" then
							sqlValue = sqlValue & "N" & pkstr(RSreg(param.selectSingleNode("fieldName").text),",")		
						else
							sqlValue = sqlValue & pkstr(RSreg(param.selectSingleNode("fieldName").text),",")		
						end if
					end if		
				next  				
				sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
				sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"
				sql = oConvert.com2sim(sql)
				set RSx = conn.execute(sql)
				xNewIdentity = RSx(0)
				'----Slave table處理
				if xSlaveTable = xsbaseTableName then
					sql = "INSERT INTO  " & xsbaseTableName & "(gicuitem,"
					sqlValue = ") Values(" & xNewIdentity & ","
					for each param in refModel.selectNodes("fieldList/field") 
					    xdataType = nullText(param.selectSingleNode("dataType"))
					    if not isNull(RSSlave(param.selectSingleNode("fieldName").text)) then
						sql = sql & param.selectSingleNode("fieldName").text & ","
						if xdataType = "char" or xdataType = "varchar" or xdataType = "text" or xdataType = "nchar" or xdataType = "nvarchar" or xdataType = "ntext" then
							sqlValue = sqlValue & "N" & pkstr(RSSlave(param.selectSingleNode("fieldName").text),",")		
						else
							sqlValue = sqlValue & pkstr(RSSlave(param.selectSingleNode("fieldName").text),",")		
						end if
					    end if
					next
					sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"	
				else
					sql = "INSERT INTO  " & xsbaseTableName & "(gicuitem) Values(" & xNewIdentity & ")"
				end if
				sql = oConvert.com2sim(sql)
				conn.Execute(SQL)			
				'----關鍵字詞處理
			    SQLK = "Select xkeyword,weight from CuDtkeyword where icuitem=" & session("convertsim_icuitem")
			    Set RSK = conn.execute(SQLK)
			    if not RSK.eof then
				keywordStr = ""
				while not RSK.eof
				    keywordStr = keywordStr & "Insert Into CuDtkeyword Values(" & xNewIdentity & ",N'"&RSK("xkeyword")&"',"&RSK("weight")&");"
				    RSK.movenext
				wend
				if keywordStr <> "" then conn.execute(oConvert.com2sim(keywordStr))
		  	    end if
			end if
		next
	end if

set oConvert = nothing
%>
