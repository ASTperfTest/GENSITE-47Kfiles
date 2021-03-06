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
HTUploadPath=session("Public")+"Attachment/"
apath=server.mappath(HTUploadPath) & "\"

'----941207繁轉簡
if checkGIPconfig("convertsim") then
	Set oConvert = CreateObject("Convertsim.Convert")
end if

set htPageDom = session("codeXMLSpec")

set refModel = htPageDom.selectSingleNode("//dsTable")
set allModel = htPageDom.selectSingleNode("//DataSchemaDef")
SQLCDG = "Select CDG.*,u.checkYn from CuDtGeneric CDG " & _
	"Left Join CtUnit AS u ON CDG.ictunit=u.ctUnitId " & _
	"where CDG.icuitem=" & request.querystring("icuitem")
set RSreg = conn.execute(SQLCDG)
'hying edit
old_showtype=RSreg("showtype")
if old_showtype="2" or old_showtype="3" then
   new_showtype=old_showtype
else
   new_showtype=request("showType")
end if
xSlaveTable = refModel.selectSingleNode("tableName").text
SQLSlave = "Select * from "&xSlaveTable&" where gicuitem = " & request.querystring("icuitem")
set RSSlave = conn.execute(SQLSlave)
	
ctNodeId = request("ctNodeId")

MyArray = Split(ctNodeId,",")
for i=0 to UBound(MyArray)
	sql = "SELECT u.*,b.sbaseTableName FROM CatTreeNode AS n LEFT JOIN CtUnit AS u ON u.ctUnitId=n.ctUnitId" _
		& " Left Join BaseDsd As b ON u.ibaseDsd=b.ibaseDsd" _
		& " WHERE n.ctNodeId=" & MyArray(i)
	set RS = Conn.execute(sql)
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
	sql = "INSERT INTO  CuDtGeneric(fctupublic,ibaseDsd,ictunit,showType,refId,deditDate,created_Date,"
	sqlValue = ") VALUES('"&xfCTUPublic&"',"&xibaseDsd&","&xctUnitId&",'"+ new_showtype +"',"&request.querystring("icuitem")&",getdate(),getdate()," 
	for each param in allModel.selectNodes("dsTable[tableName='CuDtGeneric']/fieldList/field[fieldName!='icuitem' and fieldName!='ibaseDsd' and fieldName!='ictunit' and fieldName!='fctupublic' and fieldName!='showType' and fieldName!='refID' and fieldName!='deditDate' and fieldName!='created_Date']") 
		xdataType = nullText(param.selectSingleNode("dataType"))
		if not isNull(RSreg(param.selectSingleNode("fieldName").text)) then
			sql = sql & param.selectSingleNode("fieldName").text & ","
			'if xdataType = "char" or xdataType = "varchar" or xdataType = "text" or xdataType = "nchar" or 'xdataType = "nvarchar" or xdataType = "ntext" then
				'sqlValue = sqlValue & "N" & pkstr(RSreg(param.selectSingleNode("fieldName").text),",")		
			'else
				sqlValue = sqlValue & pkstr(RSreg(param.selectSingleNode("fieldName").text),",")		
			'end if
		end if		
	next  
	
	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
	sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"

	if checkGIPconfig("convertsim") and request("sim")="Y" then
		sql = oConvert.com2sim(sql)
	end if	
'response.write sql
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
	if checkGIPconfig("convertsim") and request("sim")="Y" then
		sql = oConvert.com2sim(sql)
	end if	
	conn.Execute(SQL)			
	'----關鍵字詞處理
	if checkGIPconfig("convertsim") and request("sim")="Y" then
	    SQLK = "Select xkeyword,weight from CuDtkeyword where icuitem=" & request.querystring("icuitem")
	    Set RSK = conn.execute(SQLK)
	    if not RSK.eof then
		keywordStr = ""
		while not RSK.eof
		    keywordStr = keywordStr & "Insert Into CuDtkeyword Values(" & xNewIdentity & ",N'"&RSK("xkeyword")&"',"&RSK("weight")&");"
		    RSK.movenext
		wend
		if keywordStr <> "" then conn.execute(oConvert.com2sim(keywordStr))
  	    end if
	else
	    sql = "Insert Into CuDtkeyword Select "&xNewIdentity&",xkeyword,weight from CuDtkeyword where icuitem=" & request.querystring("icuitem")
	    conn.Execute(SQL)
	end if		
next
'response.end
if checkGIPconfig("convertsim") then
	set oConvert = nothing
end if
        '----940627多向出版(複製)時,附件一併複製
	'if request("showType")="4" or request("showType")="5"  then
		SQLAttach="Select ixCuAttach,NFileName,I.* from CuDTAttach C Left Join ImageFile I ON C.NFileName=I.NewFileName " & _
			"where xiCuItem=" & request.querystring("iCuItem")

		set RSAttach=conn.execute(SQLAttach)
		if not RSAttach.eof then
		    SQLAttach=""
		    while not RSAttach.eof
		    	Set fso = CreateObject("Scripting.FileSystemObject")
		    	if fso.fileExists(apath & RSAttach("NFileName")) then
		    		'----檔名編碼
			    	randomize
				ofname = RSAttach("OldFileName")
				fnExt = ""
				if instrRev(ofname, ".")>0 then	fnext=mid(ofname, instrRev(ofname, "."))
				tstr = now()
				nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & Int((1000 - 1 + 1) * Rnd + 1) & fnext
		    		'----複製檔案
				Set SourceFile = fso.GetFile(apath & RSAttach("NFileName"))
				SourceFile.Copy apath + nfname		    		
				'----新增資料至CuDTAttach
				SQLAttach=SQLAttach&"Insert Into CuDTAttach Select "&xNewIdentity&",aTitle,aDesc,OFileName,'"& nfname &"',aEditor,aEditDate,bList,listSeq " & _
					"from CuDTAttach where ixCuAttach=" & RSAttach("ixCuAttach") & ";Insert Into ImageFile(newFileName,oldFileName) values('"& nfname & "','"& RSAttach("OldFileName") &"');"	
                             
                        else
                           
		    	end if
			RSAttach.movenext
		    wend
		    if SQLAttach<>"" then conn.execute(SQLAttach)
		end if
	'end if

%>
<script language=VBS>
	alert "多向出版完成!"
	'window.navigate "DsdXMLList.asp"
</script>
</html>
