<%@ CodePage = 65001 %>
<% Response.Expires = 0%>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->

<%

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

'----1.取得所需參數
set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
htPageDom.async = false
htPageDom.setProperty("ServerHTTPRequest") = true	
LoadXML = server.MapPath("/GipDSD/CtUnitCrossSiteCopy.xml")
xv = htPageDom.load(LoadXML)
if htPageDom.parseError.reason <> "" then 
	Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
	Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
	Response.End()
end if
'----Source DB Name/Target DB Name
sourceDB = nullText(htPageDom.selectSingleNode("CtUnitCrossSiteCopy/sourceDB"))
targetDB = nullText(htPageDom.selectSingleNode("CtUnitCrossSiteCopy/targetDB"))


'----2.Loop CtUintCopyList
for each CtUintCopy in htPageDom.selectNodes("CtUnitCrossSiteCopy/CtUintCopyList/CtUintCopy")   
	on error resume next
	sourceCtUnitID = nullText(CtUintCopy.selectSingleNode("sourceCtUnitID"))
	targetCtUnitID = nullText(CtUintCopy.selectSingleNode("targetCtUnitID"))
	'----來源SQL
	sqlSource = "SELECT u.* FROM "&sourceDB&".dbo.CtUnit AS u " _
		& " Left Join "&sourceDB&".dbo.BaseDSD As b ON u.iBaseDSD=b.iBaseDSD" _
		& " WHERE u.CtUnitID=" & sourceCtUnitID
	set RSSource = Conn.execute(sqlSource)
	sCtUnitID = RSSource("CtUnitID")
	siBaseDSD = RSSource("iBaseDSD")
	'----目標SQL
	sqlTarget = "SELECT u.* FROM "&targetDB&".dbo.CtUnit AS u " _
		& " Left Join "&targetDB&".dbo.BaseDSD As b ON u.iBaseDSD=b.iBaseDSD" _
		& " WHERE u.CtUnitID=" & targetCtUnitID
	set RSTarget = Conn.execute(sqlTarget)
	tCtUnitID = RSTarget("CtUnitID")
	tiBaseDSD = RSTarget("iBaseDSD")
	if err.number<>0 then
		response.write "sourceCtUnitID["+sourceCtUnitID+"]/targetCtUnitID["+targetCtUnitID+"]錯誤發生!<br>發生原因:"+err.description+"<hr>"
	else
	    '----選取複製目標
	    SQLMaster = "Select CDTG.*,CDT.* from "&sourceDB&".dbo.CuDTGeneric CDTG " & _
	    	"Left Join "&sourceDB&".dbo.CuDTx" & siBaseDSD & " CDT ON CDTG.iCuItem=CDT.giCuItem " & _
	    	"WHERE iCtUnit=" & sCtUnitID
	    set RSMaster = conn.execute(SQLMaster)
	    '----先刪除處理
	    	'----刪除targetDB.dbo.CuDTGeneric資料
	    	SQLDelete1 = "Delete CDTG from "&targetDB&".dbo.CuDTGeneric CDTG " & _
	    		"Left Join GIPhy.dbo.CtUnitCopy CC ON CDTG.iCuItem=CC.iCuItem " & _
	    		"where CC.sourceDB=N'"&sourceDB&"' and CC.targetDB=N'"&targetDB&"' and CC.sCtUnitID="&sCtUnitID&" and CC.tCtUnitID="&tCtUnitID&";"
	    	'----刪除targetDB.dbo.CuDTx??資料
	    	SQLDelete2 = "Delete CDTG from "&targetDB&".dbo.CuDTx" & tiBaseDSD & " CDTG " & _
	    		"Left Join GIPhy.dbo.CtUnitCopy CC ON CDTG.giCuItem=CC.iCuItem " & _
	    		"where CC.sourceDB=N'"&sourceDB&"' and CC.targetDB=N'"&targetDB&"' and CC.sCtUnitID="&sCtUnitID&" and CC.tCtUnitID="&tCtUnitID&";"
	    	'----刪除GIPhy.dbo.CtUnitCopy資料
	    	SQLDelete3 = "Delete GIPhy.dbo.CtUnitCopy where sourceDB=N'"&sourceDB&"' and targetDB=N'"&targetDB&"' and sCtUnitID="&sCtUnitID&" and tCtUnitID="&tCtUnitID&";"
	    	SQLDelete = SQLDelete1+SQLDelete2+SQLDelete3
'response.write SQLDelete
'response.end	    	
	    	conn.execute(SQLDelete)
	    '----再新增copy處理
	    if not RSMaster.EOF then
	    	while not RSMaster.EOF 
	    	    '----Master Table處理
		    sql = "INSERT INTO  "&targetDB&".dbo.CuDTGeneric(iBaseDSD,iCtUnit,dEditDate,Created_Date,"
		    sqlValue = ") VALUES("&tiBaseDSD&","&tCtUnitID&",getdate(),getdate()," 
	    	    if not ISNULL(RSMaster("fCTUPublic")) then
			sql = sql & "fCTUPublic,"
			sqlValue = sqlValue & pkstr(RSMaster("fCTUPublic"),",")			    	    
	    	    end if
	    	    if not ISNULL(RSMaster("avBegin")) then
			sql = sql & "avBegin,"
			sqlValue = sqlValue & pkstr(RSMaster("avBegin"),",")			    	    
	    	    end if
	    	    if not ISNULL(RSMaster("avEnd")) then
			sql = sql & "avEnd,"
			sqlValue = sqlValue & pkstr(RSMaster("avEnd"),",")			    	    
	    	    end if
	    	    if not ISNULL(RSMaster("sTitle")) then
			sql = sql & "sTitle,"
			sqlValue = sqlValue & pkstr(RSMaster("sTitle"),",")			    	    
	    	    end if
	    	    if not ISNULL(RSMaster("iEditor")) then
			sql = sql & "iEditor,"
			sqlValue = sqlValue & pkstr(RSMaster("iEditor"),",")			    	    
	    	    end if
	    	    if not ISNULL(RSMaster("iDept")) then
			sql = sql & "iDept,"
			sqlValue = sqlValue & pkstr(RSMaster("iDept"),",")			    	    
	    	    end if
	    	    if not ISNULL(RSMaster("topCat")) then
			sql = sql & "topCat,"
			sqlValue = sqlValue & pkstr(RSMaster("topCat"),",")			    	    
	    	    end if
	    	    if not ISNULL(RSMaster("vGroup")) then
			sql = sql & "vGroup,"
			sqlValue = sqlValue & pkstr(RSMaster("vGroup"),",")			    	    
	    	    end if
	    	    if not ISNULL(RSMaster("xKeyword")) then
			sql = sql & "xKeyword,"
			sqlValue = sqlValue & pkstr(RSMaster("xKeyword"),",")			    	    
	    	    end if
	    	    if not ISNULL(RSMaster("xImportant")) then
			sql = sql & "xImportant,"
			sqlValue = sqlValue & RSMaster("xImportant") & ","		    	    
	    	    end if
	    	    if not ISNULL(RSMaster("xURL")) then
			sql = sql & "xURL,"
			sqlValue = sqlValue & pkstr(RSMaster("xURL"),",")			    	    
	    	    end if
	    	    if not ISNULL(RSMaster("xNewWindow")) then
			sql = sql & "xNewWindow,"
			sqlValue = sqlValue & pkstr(RSMaster("xNewWindow"),",")			    	    
	    	    end if
	    	    if not ISNULL(RSMaster("xPostDate")) then
			sql = sql & "xPostDate,"
			sqlValue = sqlValue & pkstr(RSMaster("xPostDate"),",")			    	    
	    	    end if
	    	    if not ISNULL(RSMaster("xBody")) then
			sql = sql & "xBody,"
			sqlValue = sqlValue & pkstr(RSMaster("xBody"),",")			    	    
	    	    end if
		    sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
		    sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"
	    	    set RSx = conn.Execute(SQL)
	    	    xNewIdentity = RSx(0)  
	    	    '----Slave table處理  
		    sql = "INSERT INTO  "&targetDB&".dbo.CuDTx" & tiBaseDSD & "(giCuItem,"
		    sqlValue = ") Values(" & xNewIdentity & ","	 
		    if tiBaseDSD = "2" then
	    	    	if not ISNULL(RSMaster("CuDTx2F5")) then
			    sql = sql & "CuDTx2F5,"
			    sqlValue = sqlValue & pkstr(RSMaster("CuDTx2F5"),",")			    	    
	    	    	end if		    
	    	    	if not ISNULL(RSMaster("CuDTx2F7")) then
			    sql = sql & "CuDTx2F7,"
			    sqlValue = sqlValue & pkstr(RSMaster("CuDTx2F7"),",")			    	    
	    	    	end if		    
	    	    	if not ISNULL(RSMaster("CuDTx2F9")) then
			    sql = sql & "CuDTx2F9,"
			    sqlValue = sqlValue & pkstr(RSMaster("CuDTx2F9"),",")			    	    
	    	    	end if		    
		    elseif tiBaseDSD = "4" then  	    	
	    	    	if not ISNULL(RSMaster("CuDTx4F14")) then
			    sql = sql & "CuDTx4F14,"
			    sqlValue = sqlValue & pkstr(RSMaster("CuDTx4F14"),",")			    	    
	    	    	end if		    
		    elseif tiBaseDSD = "7" then  	    	
	    	    	if not ISNULL(RSMaster("CuDTx7F8")) then
			    sql = sql & "CuDTx7F8,"
			    sqlValue = sqlValue & pkstr(RSMaster("CuDTx7F8"),",")			    	    
	    	    	end if		    
		    end if
		    sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"	
		    conn.Execute(sql)
		    '----
		    SQLLog = "Insert Into GIPhy.dbo.CtUnitCopy values('"&sourceDB&"','"&targetDB&"',"&sCtUnitID&","&tCtUnitID&","&xNewIdentity&",'"&siBaseDSD&"',"&tiBaseDSD&")"
		    conn.execute(SQLLog)
		    RSMaster.movenext
	        wend
	    end if
	    if err.number<>0 then
	    	response.write "sourceCtUnitID["+sourceCtUnitID+"]/targetCtUnitID["+targetCtUnitID+"]錯誤發生!<br>發生原因:"+err.description+"<hr>"
	    else
	    	response.write "sourceCtUnitID["+sourceCtUnitID+"]-->targetCtUnitID["+targetCtUnitID+"]完成!<hr>"
	    end if	    
'response.write sCtUnitID&"[]"&siBaseDSD&"[]"&xSlaveTable&"[]"&tCtUnitID&"[]"&tiBaseDSD&"<br>" 	
	end if
next
response.write "<hr>完成2!"
response.end

%>
