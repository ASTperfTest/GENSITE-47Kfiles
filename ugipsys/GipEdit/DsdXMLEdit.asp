<%@ CodePage = 65001 %>
<% CodePage=65001
Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="GC1AP1"
HTProgPrefix="DsdXML" 
Server.ScriptTimeout = 120000
%>

<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<!--#include virtual = "/inc/hyftdGIP.inc" -->
<!-- #INCLUDE FILE="KnowledgeKpiFunction.inc" -->
<%
Function Send_email (S_email,R_email,Re_Sbj,Re_Body)

	Set objNewMail = CreateObject("CDONTS.NewMail") 
	objNewMail.MailFormat = 0
	objNewMail.BodyFormat = 0 
	call objNewMail.Send(S_email,R_email,Re_Sbj,Re_Body)

	Set objNewMail = Nothing
End Function	
function MMOPathStr(MMOFolderID)	'----940407 MMO路徑字串
	sql = "SELECT MM.mmositeId,mmofolderParent,Case mmofolderParent when 'zzz' then mmositeName else mmofolderNameShow END MMOFolderNameShow " & _ 
		"FROM Mmofolder MM Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId where mmofolderID=" & MMOFolderID
	set RSN = conn.execute(sql)
	xParent = RSN("mmofolderParent")
	xPathStr = RSN("MMOFolderNameShow")
	xMMOSiteID = RSN("mmositeId")
	while xParent <> "zzz"
		sql = "SELECT MM.mmositeId,mmofolderParent,Case mmofolderParent when 'zzz' then mmositeName else mmofolderNameShow END MMOFolderNameShow " & _ 
			"FROM Mmofolder MM Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId where MM.mmositeId=" & pkstr(xMMOSiteID,"") & " and mmofolderName=" & pkstr(xParent,"")
		set RS = conn.execute(sql)
		xPathStr = RS("MMOFolderNameShow") & " / " & xPathStr
		xParent = RS("mmofolderParent")
		xMMOSiteID = RSN("mmositeId")
	wend
	MMOPathStr = xPathStr
end function
function FilmRelated(xfunc,xTable,xType,xicuitem,xFieldStr)
	Set Conn2 = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	Conn2.Open session("ODBCDSN")
'Set Conn2 = Server.CreateObject("HyWebDB3.dbExecute")
Conn2.ConnectionString = session("ODBCDSN")
Conn2.ConnectionTimeout=0
Conn2.CursorLocation = 3
Conn2.open
'----------HyWeb GIP DB CONNECTION PATCH----------

    	xxicuitem=xicuitem
    	if xTable="Corp" then 		'----Corp處理
	    if xfunc="edit" then 
	    	SQLD="delete from FilmCorpInfo where FilmNo="&xxicuitem&" and CompanyType=N'"&xType&"'"  
	    	conn2.execute(SQLD)
	    end if
	    xKeywordArray=split(xFieldStr,",")
	    for i=0 to ubound(xKeywordArray)
	        xStr=trim(xKeywordArray(i))
		SQL="Select gicuitem from CorpInformation AI Left Join CuDtGeneric CDT " & _
			" ON AI.gicuitem=CDT.icuitem where sTitle=N'"&xStr&"'"
		Set RSC=conn2.execute(SQL)
		if not RSC.eof then 
			SQLI=SQLI+"Insert Into FilmCorpInfo values("&xxicuitem&","&RSC(0)&",'"&xType&"',null,null,null);"
		end if
	    next     
    	elseif xTable="Actor" then 	'----People處理
	    if xfunc="edit" then 
	    	SQLD="delete from FilmPeopleInfo where FilmNo="&xxicuitem&" and RoleInfo=N'"&xType&"'"  
	    	conn2.execute(SQLD)
	    end if
	    xKeywordArray=split(xFieldStr,",")
	    for i=0 to ubound(xKeywordArray)
		'----取最後括號
		xPos=instrRev(xKeywordArray(i),"(")
		if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
			xStrPar="'"+mid(trim(xKeywordArray(i)),xPos)+"'"
		else
			xStr=trim(xKeywordArray(i))
			xStrPAr="null"
		end if
		SQL="Select gicuitem from ActorInformation AI Left Join CuDtGeneric CDT " & _
			" ON AI.gicuitem=CDT.icuitem where sTitle=N'"&xStr&"'"
		Set RSC=conn2.execute(SQL)
		if not RSC.eof then 
			SQLI=SQLI+"Insert Into FilmPeopleInfo values("&xxicuitem&","&RSC(0)&",'"&xType&"',"&xStrPAr&",null,null,null);"
		end if
	    next 
    	end if	
	if SQLI<>"" then conn2.execute(SQLI)  
'    	conn2.close
    	FilmRelated = ""
end function

Dim hyftdGIPStr
Dim pKey
Dim RSreg
Dim formFunction
Dim allModel2
Dim xshowTypeStr,xshowType
Dim xRef5Count,xref5YN
Dim orgInputType
Dim xMMOFolderID,MMOPath,xFTPIPMMO,xFTPPortMMO,xFTPIDMMO,xFTPPWDMMO,MMOFTPfilePath,MMOCount
MMOCount=0

mynowpage = request("nowpage")
mypagesize = request("pagesize")

taskLable="編輯" & HTProgCap
	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"
	
'----ftp參數處理
	FTPErrorMSG=""
	FTPfilePath="public/data"
	SQLP = "Select * from UpLoadSite where upLoadSiteId='file'"
	Set RSP = conn.execute(SQLP)
	if not RSP.EOF  then
   		xFTPIP = RSP("UpLoadSiteFTPIP")
   		xFTPPort = RSP("UpLoadSiteFTPPort")
   		xFTPID = RSP("UpLoadSiteFTPID")
   		xFTPPWD = RSP("UpLoadSiteFTPPWD")
   	end if
'----ftp參數處理end 
	
apath=server.mappath(HTUploadPath) & "\"
'response.write apath
'response.end

if request.querystring("phase")="edit" or session("BatchDphase") = "edit" then
	Set xup = Server.CreateObject("TABS.Upload")
else
	Set xup = Server.CreateObject("TABS.Upload")
	xup.codepage=65001
	xup.Start apath
end if

function xUpForm(xvar)
	xUpForm = xup.form(xvar)
end function


if request.querystring("S")="Approve" or request.querystring("S")="Ref" then	'----由GipApprove或編修參照來源而來,沒有session("codeXMLSpec")
	sql = "SELECT u.*,b.sbaseTableName FROM CuDtGeneric AS n LEFT JOIN CtUnit AS u ON u.ctUnitId=n.ictunit" _
		& " Left Join BaseDsd As b ON n.ibaseDsd=b.ibaseDsd" _
		& " WHERE n.icuitem=" & request.querystring("icuitem")
	'response.write "<HR>" & sql & "</HR>"
	set RS = Conn.execute(sql)
	session("cuItem") = request.querystring("icuitem")
	session("ctUnitId") = RS("ctUnitId")
	session("ctUnitName") = RS("CtUnitName")
	session("iBaseDSD") = RS("ibaseDsd")
	session("fCtUnitOnly") = RS("fCtUnitOnly")
	if isNull(RS("sbaseTableName")) then
		session("sBaseTableName") = "CuDTx" & session("iBaseDSD")
	else
		session("sBaseTableName") = RS("sbaseTableName")
	end if	
	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true		
    	'----找出對應的CtUnitX???? xmlSpec檔案(若找不到則抓default), 並依fieldSeq排序成物件存入session
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(session("ctUnitId")) & ".xml") 	
    	if fso.FileExists(filePath) then
    		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(session("ctUnitId")) & ".xml")
    	else
    		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & session("iBaseDSD") & ".xml")
    	end if 
'	response.write LoadXML & "<HR>"
'	response.end
	xv = htPageDom.load(LoadXML)
'	response.write xv & "<HR>"
  	if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  	end if 	
   	set root = htPageDom.selectSingleNode("DataSchemaDef")
    	'----Load XSL樣板
    	set oxsl = server.createObject("microsoft.XMLDOM")
   	oxsl.async = false
   	xv = oxsl.load(server.mappath("/GipDSD/xmlspec/CtUnitXOrder.xsl"))    		
    	'----複製Slave的dsTable,並依順序轉換
	set DSDNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='"&session("sBaseTableName")&"']").cloneNode(true)    
    	set DSDNodeXML = server.createObject("microsoft.XMLDOM")
   	DSDNodeXML.appendchild DSDNode
    	set nxml = server.createObject("microsoft.XMLDOM")
    	nxml.LoadXML(DSDNodeXML.transformNode(oxsl))
    	set nxmlnewNode = nxml.documentElement    
    	DSDNode.replaceChild nxmlnewNode,DSDNode.selectSingleNode("fieldList")
    	root.replaceChild DSDNode,root.selectSingleNode("dsTable[tableName='"&session("sBaseTableName")&"']")
    	'----複製CuDtGeneric的dsTable,並依順序轉換
    	set GenericNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDtGeneric']").cloneNode(true)    
    	set GenericNodeXML = server.createObject("microsoft.XMLDOM")
    	GenericNodeXML.appendchild GenericNode
   	set nxml2 = server.createObject("microsoft.XMLDOM")
    	nxml2.LoadXML(GenericNodeXML.transformNode(oxsl))
    	set nxmlnewNode2 = nxml2.documentElement    
    	GenericNode.replaceChild nxmlnewNode2,GenericNode.selectSingleNode("fieldList")
    	root.replaceChild GenericNode,root.selectSingleNode("dsTable[tableName='CuDtGeneric']")       	


  	set session("codeXMLSpec") = htPageDom
  	'----混合field順序
	set nxml0 = server.createObject("microsoft.XMLDOM")
	nxml0.LoadXML(htPageDom.transformNode(oxsl))
	set session("codeXMLSpec2") = nxml0	
'	response.write nxm10.xml
'	response.end
'response.write "Hello2"
'response.end
end if

  	set htPageDom = session("codeXMLSpec")

  	set refModel = htPageDom.selectSingleNode("//dsTable")
  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")
  	'----940215Film關聯處理session
	session("FilmRelated_CorpActor")=nullText(allModel.selectSingleNode("FilmRelated_CorpActor"))  	
  	
    set ideptNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='idept']")
    ideptNode.selectSingleNode("inputType").text = "refSelect"  
    ideptNode.selectSingleNode("refLookup").text = "refDept"  
    
	if xUpForm("submitTask") = "UPDATE" then
	
		errMsg = ""
		checkDBValid()
		if errMsg <> "" then
			EditInBothCase()
		else
			doUpdateDB()
			showDoneBox("資料更新成功！")
		end if
		'======	2006.5.24 by Gary
		
	elseif xUpForm("submitTask") = "DELETE" or session("BatchDsubmitTask")="DELETE" then
		
		'Dim ActFlag
		'ActFlag = false
		'---知識家的活動---
		if session("ibaseDsd") = "39" then				
			'Actsql = "SELECT * FROM Activity WHERE (GETDATE() BETWEEN ActivityStartTime AND ActivityEndTime)"
			'set ActRs = conn.execute(Actsql)
			'---it is in the act time---
			'if not ActRs.eof then
			'	ActFlag = true
			'end If				
			'---delete data in activityinfo---		
			'---刪除AP table中的資料---question---				
			
			if session("ctUnitId") = "932" then			
				qicuitem = xUpForm("htx_icuitem")	
				sql = "SELECT Status FROM KnowledgeForum WHERE gicuitem = " & qicuitem	
				set rs = conn.execute(sql)
				if not rs.eof then
					if rs("Status") = "D" then 
						response.write "<script language=""javascript"">alert('文章不可重複刪除!!');history.back();</script>"
						response.end
					end if
				end if				
				'---kpi---
				'------檢查是否有討論------
				sql = "SELECT CuDTGeneric.iCUItem FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
				sql = sql & "WHERE (KnowledgeForum.ParentIcuitem = " & qicuitem & ") AND (KnowledgeForum.Status = 'N') AND (CuDTGeneric.iCtUnit = 933)"
				
				set rs = conn.execute(sql)
				while not rs.eof 
					DeleteCommend rs("iCUItem") '---刪除評價---	
					DeleteOpinion rs("iCUItem") '---刪除意見---			
					DeleteDiscuss rs("iCUItem") '---刪除討論---		
					rs.movenext
				wend
				rs.close
				set rs = nothing
				'-----------------------------------------------------------------------------------		
				'---刪除發問---
				DeleteQuestion qicuitem	
				UpdateStatus qicuitem
				
				'---原先的delete--------------------------------------------------------------------------------	
				'qicuitem = xUpForm("htx_icuitem")		
				'qeditor = xUpForm("htx_ieditor")								
				'---選出此筆問題的討論---
				'sql = "SELECT icuitem, iEditor FROM CuDtGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.Icuitem = KnowledgeForum.gicuitem " & _
				'			"WHERE KnowledgeForum.ParentIcuitem = '" & qicuitem & "' "																
				'set drs1 = conn.execute(sql)
				'while not drs1.eof 						
				'	dicuitem = drs1("icuitem")		
				'	deditor = drs1("iEditor")						
				'	'---扣有勾選的討論分數---
				'	if xUpForm("htx_vgroup") = "A" then
				'		sql = "SELECT DiscussCheckGrade FROM ActivityMemberNew WHERE MemberId = '" & deditor & "' "
				'		set dcrs = conn.execute(sql)
				'		if not dcrs.eof then
				'			DiscussCheckGrade = dcrs("DiscussCheckGrade")						
				'			'---若 Activity1DiscussGrade = 0 , 保持為0, 若>0, 才要扣分---
				'			If DiscussCheckGrade <= 2 Then
				'				DiscussCheckGrade = 0
				'			Else
				'				DiscussCheckGrade = DiscussCheckGrade - 2
				'			End If					
				'			'---更新回TABLE---
				'			sql = "UPDATE ActivityMemberNew SET DiscussCheckGrade = " & DiscussCheckGrade & " WHERE MemberId = '" & deditor & "' "
				'			conn.execute(sql)
				'		end if					
				'		set dcrs = nothing
				'	end if					
				'	'---flow : 刪除此篇討論時, 此member的ActivityMember中的Activity1DiscussGrade 扣2分回來---
				'	sql = "SELECT DiscussGrade FROM ActivityMemberNew WHERE MemberId = '" & deditor & "' "
				'	set dcrs = conn.execute(sql)
				'	if not dcrs.eof then
				'		DiscussGrade = dcrs("DiscussGrade")
				'		'---若 Activity1DiscussGrade = 0 , 保持為0, 若>0, 才要扣分---
				'		If DiscussGrade <= 2 Then
				'			DiscussGrade = 0
				'		Else
				'			DiscussGrade = DiscussGrade - 2
				'		End If					
				'		'---更新回TABLE---
				'		sql = "UPDATE ActivityMemberNew SET DiscussGrade = " & DiscussGrade & " WHERE MemberId = '" & deditor & "' "
				'		conn.execute(sql)
				'	end if					
				'	set dcrs = nothing											
				'	drs1.movenext
				'wend
				'set drs1 = nothing
				
				'if xUpForm("htx_vgroup") = "A" then
					'---flow : 刪除此篇問題時, 此member的ActivityMember中的QuestionGrade 扣2分回來---
				'	sql = "SELECT QuestionCheckGrade FROM ActivityMemberNew WHERE MemberId = '" & qeditor & "' "
				'	set qcrs = conn.execute(sql)
				'	if not qcrs.eof then
				'		QuestionCheckGrade = qcrs("QuestionCheckGrade")
				'		'---若 Activity1DiscussGrade = 0 , 保持為0, 若>0, 才要扣分---
				'		If QuestionCheckGrade <= 2 Then
				'			QuestionCheckGrade = 0
				'		Else
				'			QuestionCheckGrade = QuestionCheckGrade - 2
				'		End If					
				'		'---更新回TABLE---
				'		sql = "UPDATE ActivityMemberNew SET QuestionCheckGrade = " & QuestionCheckGrade & " WHERE MemberId = '" & qeditor & "' "
				'		conn.execute(sql)	
				'	end if				
				'	set qcrs = nothing
				'end if
				
				'---flow : 刪除此篇問題時, 此member的ActivityMember中的QuestionGrade 扣2分回來---
				'sql = "SELECT QuestionGrade FROM ActivityMemberNew WHERE MemberId = '" & qeditor & "' "
				'set qcrs = conn.execute(sql)
				'if not qcrs.eof then
				'	QuestionGrade = qcrs("QuestionGrade")
				'	'---若 Activity1DiscussGrade = 0 , 保持為0, 若>0, 才要扣分---
				'	If QuestionGrade <= 2 Then
				'		QuestionGrade = 0
				'	Else
				'		QuestionGrade = QuestionGrade - 2
				'	End If					
				'	'---更新回TABLE---
				'	sql = "UPDATE ActivityMemberNew SET QuestionGrade = " & QuestionGrade & " WHERE MemberId = '" & qeditor & "' "
				'	conn.execute(sql)	
				'end if				
				'set qcrs = nothing
			'---刪除AP table中的資料---discuss---	
			'-----------------------------------------------------------------------------------	
			
			elseif session("ctUnitId") = "933" then			
				
				dicuitem = xUpForm("htx_icuitem")	
				
				sql = "SELECT Status FROM KnowledgeForum WHERE gicuitem = " & dicuitem	
				set rs = conn.execute(sql)
				if not rs.eof then
					if rs("Status") = "D" then 
						response.write "<script language=""javascript"">alert('article can not be deleted!!');window.location.href='KnowledgeForumlist.asp';</script>"
						response.end
					end if
				end if	
	
				DeleteCommend dicuitem '---刪除評價---	
				DeleteOpinion dicuitem '---刪除意見---				
				DeleteDiscuss dicuitem '---刪除討論---		
		
				'---更新parent的count---				
				sql = "SELECT * FROM KnowledgeForum WHERE gicuitem = " & dicuitem 
				set delrs = conn.execute(sql)
				while not delrs.eof 
					CommandCount = CInt(delrs("CommandCount"))
					GradeCount = CInt(delrs("GradeCount"))
					GradePersonCount = CInt(delrs("GradePersonCount"))
					ParentIcuitem = delrs("ParentIcuitem")					
					sql = "UPDATE KnowledgeForum SET DiscussCount = DiscussCount - 1, CommandCount = CommandCount - " & CommandCount & ", " & _
								"GradeCount = GradeCount - " & GradeCount & ", GradePersonCount = GradePersonCount - " & GradePersonCount & " " & _
								"WHERE gicuitem = " & ParentIcuitem			
					'response.write sql & "<hr />"
					conn.execute(sql)					
					delrs.movenext					
				wend				
				delrs.close
				set delrs = nothing
				
			elseif session("ctUnitId") = "935" then			
			
				dicuitem = xUpForm("htx_icuitem")	
				
				sql = "SELECT * FROM KnowledgeForum WHERE gicuitem = " & dicuitem	
				set rs = conn.execute(sql)
				if not rs.eof then
					if rs("Status") = "D" then 
						'response.write "<script language=""javascript"">alert('article can not be deleted!!');history.back();</script>"
						'response.end
					end if
					ParentIcuitem = rs("ParentIcuitem")			
				end if	
				rs.close
				set rs = nothing
				
				DeleteOpinion ParentIcuitem '---刪除意見---		
				
				ParentIcuitem = ""
				sql = "SELECT * FROM KnowledgeForum WHERE gicuitem = " & dicuitem 
				set delrs = conn.execute(sql)
				if not delrs.eof then
					ParentIcuitem = delrs("ParentIcuitem")					
					sql = "UPDATE KnowledgeForum SET CommandCount = CommandCount - 1 WHERE gicuitem = " & ParentIcuitem								
					'response.write sql
					conn.execute(sql)					
				end if											
				delrs.close
				set delrs = nothing
				
				sql = "SELECT * FROM KnowledgeForum WHERE gicuitem = " & ParentIcuitem 
				set delrs = conn.execute(sql)
				if not delrs.eof then
					ParentIcuitem = delrs("ParentIcuitem")					
					sql = "UPDATE KnowledgeForum SET CommandCount = CommandCount - 1 WHERE gicuitem = " & ParentIcuitem								
					'response.write sql
					conn.execute(sql)					
				end if											
				delrs.close
				set delrs = nothing				
				
				'dicuitem = xUpForm("htx_icuitem")		
				'deditor = xUpForm("htx_ieditor")			
				'	
				'sql = "SELECT vGroup FROM KnowledgeForum INNER JOIN CuDTGeneric "
				'sql = sql & " ON KnowledgeForum.ParentIcuitem = CuDTGeneric.iCUItem "
				'sql = sql & " WHERE KnowledgeForum.gicuitem = " & dicuitem
				'set dcrs = conn.execute(sql)
				'if not dcrs.eof then
				'	if dcrs("vGroup") = "A" Then	
				'		
				'		Dim ActStartTime, ActEndTime
				'		sql = "SELECT ActivityStartTime, ActivityEndTime "
				'		sql = sql & " FROM Activity WHERE ActivityId = '" & session("ActivityId") & "'"
				'		set actrs = conn.execute(sql)
				'		if not actrs.eof then
				'			ActStartTime = replace(actrs("ActivityStartTime"), "上午", "")
				'			ActEndTime = replace(actrs("ActivityEndTime"), "上午", "")
				'		end if
				'		set actrs = nothing
				'		'---若這個討論是在活動期間內才要扣分---
				'		sql = "SELECT * FROM CuDTGeneric WHERE iCuItem = " & dicuitem 
				'		sql = sql & " AND dEditDate BETWEEN '" & ActStartTime & "' AND '" & ActEndTime & "'"						
				'		set inactrs = conn.execute(sql)
				'		if not inactrs.eof then 								
				'			sql = "SELECT DiscussCheckGrade FROM ActivityMemberNew WHERE MemberId = '" & deditor & "' "
				'			set dcrs1 = conn.execute(sql)
				'			if not dcrs1.eof then
				'				DiscussCheckGrade = dcrs1("DiscussCheckGrade")						
				'				'---若 Activity1DiscussGrade = 0 , 保持為0, 若>0, 才要扣分---
				'				If DiscussCheckGrade <= 2 Then
				'					DiscussCheckGrade = 0
				'				Else
				'					DiscussCheckGrade = DiscussCheckGrade - 2
				'				End If					
				'				'---更新回TABLE---
				'				sql = "UPDATE ActivityMemberNew SET DiscussCheckGrade = " & DiscussCheckGrade & " WHERE MemberId = '" & deditor & "' "
				'				conn.execute(sql)
				'			end if					
				'			set dcrs1 = nothing
				'		end if
				'		set inactrs = nothing
				'	end if
				'end if
				'set dcrs = nothing
					
				'---flow : 刪除此篇討論時, 此member的ActivityMember中的Activity1DiscussGrade 扣2分回來---
				'sql = "SELECT DiscussGrade FROM ActivityMemberNew WHERE MemberId = '" & deditor & "' "
				'set dcrs = conn.execute(sql)
				'if not dcrs.eof then
				'	DiscussGrade = dcrs("DiscussGrade")
				'	'---若 Activity1DiscussGrade = 0 , 保持為0, 若>0, 才要扣分---
				'	If DiscussGrade <= 2 Then
				'		DiscussGrade = 0
				'	Else
				'		DiscussGrade = DiscussGrade - 2
				'	End If					
				'	'---更新回TABLE---
				'	sql = "UPDATE ActivityMemberNew SET DiscussGrade = " & DiscussGrade & " WHERE MemberId = '" & deditor & "' "
				'	conn.execute(sql)
				'end if					
				'set dcrs = nothing							
			end if				
		end if		
		'---end of 知識家的活動---
	
		if session("BatchDicuitem") <> "" then
			icuitem = session("BatchDicuitem")
		else
			icuitem = request.queryString("iCuItem")
		end if		
		'======	2006.5.8 by Gary ======
		if checkGIPconfig("RSSandQuery") then  	
			SQLRSS = "SELECT YNrss FROM catTreeNode WHERE ctNodeId=" & pkStr(session("ctNodeId"),"")			
			Set RSS = conn.execute(SQLRSS)
			if not RSS.eof and RSS("YNrss")="Y" then
				session("RSS_method") = "delete"
				'session("RSS_iCuItem") = request.queryString("iCuItem")	
				session("RSS_iCuItem") = icuitem
				postURL = "/ws/ws_RSSPool.asp"
				Server.Execute (postURL) 
			end if
		end if
		'======
		'----940407 MMO上傳路徑/ftp等參數處理
		SQLCheck="Select rdsdcat,sbaseTableName from CtUnit C Left Join BaseDsd B ON C.ibaseDsd=B.ibaseDsd " & _
			"where C.CtUnitId='"&session("CtUnitID")&"'"
		Set RSCheck=conn.execute(SQLCheck)	
		if checkGIPconfig("MMOFolder") and RSCheck("rDSDCat")="MMO" then
			'----FTP所需參數
			SQLM="Select MM.mmositeId+MM.mmofolderName as MMOFOlderID,MM.mmositeId,MM.mmofolderName,MS.upLoadSiteFtpip,MS.upLoadSiteFtpport,MS.upLoadSiteFtpid,MS.upLoadSiteFtppwd " & _
				" from "&RSCheck("sbaseTableName")&" CM Left Join Mmofolder MM ON CM.mmofolderId=MM.mmofolderId " & _
				" Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId " & _
				"where CM.gicuitem="&icuitem
				'"where CM.gicuitem="&request.querystring("icuitem")
			Set RSM=conn.execute(SQLM)
			if not RSM.eof then 
				xMMOFolderID=RSM("MMOFOlderID")
	   		xFTPIPMMO = RSM("upLoadSiteFtpip")
	   		xFTPPortMMO = RSM("upLoadSiteFtpport")
	   		xFTPIDMMO = RSM("upLoadSiteFtpid")
	   		xFTPPWDMMO = RSM("upLoadSiteFtppwd")
				MMOFTPfilePath="public/"&xMMOFolderID
			end if
			'----上傳路徑
			MMOPath = session("MMOPublic") & xMMOFolderID
			if right(MMOPath,1)<>"/" then	MMOPath = MMOPath & "/"
		end if
		'----940407 MMO上傳路徑/ftp等參數處理完成
		'------- 記錄 異動 log -------- start --------------------------------------------------------	
		if checkGIPconfig("UserLogFile") then
			sql = "INSERT INTO userActionLog(loginSID,xTarget,xAction,recordNumber,objTitle) VALUES(" _
				& dfn(session("loginLogSID")) & "'0A','3'," _
				& dfn(icuitem) _
				& pkstr(xUpForm("htx_sTitle"),")")
			conn.execute sql
		end if	
		'------- 記錄 異動 log -------- end --------------------------------------------------------	
		'----刪除圖檔
		'SQLDMCheck="Select * from CuDtGeneric WHERE icuitem=" & pkStr(request.queryString("icuitem"),"")
		SQLDMCheck="Select * from CuDtGeneric WHERE icuitem=" & pkStr(icuitem,"")
		Set RSMCheck=conn.execute(SQLDMCheck)
		for each param in allModel.selectNodes("dsTable[tableName='icuitem']/fieldList/field[formList!='']") 
	    if lcase(right(nullText(param.selectSingleNode("inputType")),4)) = "file" then 
	    	if not isNull(RSMCheck(nullText(param.selectSingleNode("fieldName")))) then
	   	    processUpdate param
	    	end if
	    end if
		next
		SQLDSCheck="Select * from " & nullText(refModel.selectSingleNode("tableName")) & "  WHERE gicuitem=" & pkStr(icuitem,"")
		'SQLDSCheck="Select * from " & nullText(refModel.selectSingleNode("tableName")) & "  WHERE gicuitem=" & pkStr(request.queryString("icuitem"),"")
		Set RSSCheck=conn.execute(SQLDSCheck)
		for each param in refModel.selectNodes("fieldList/field[formList!='']") 
	    if lcase(right(nullText(param.selectSingleNode("inputType")),4)) = "file" then 
	  	  if not isNull(RSSCheck(nullText(param.selectSingleNode("fieldName")))) then
	    	  processUpdate param
	    	end if
	    end if
		next
		sql = ""
		'----刪除主表---
		'---for 知識家---
		if session("ibaseDsd") = "39" then
			'---do nothing---
			sql = ""
		else
			sql = "DELETE FROM CuDtGeneric WHERE icuitem = " & pkStr(icuitem, "")
			'sql = "DELETE FROM CuDtGeneric WHERE icuitem=" & pkStr(request.queryString("icuitem"),"")			
			'response.write sql
			'response.end
			conn.execute sql
		end if		
		sql = ""
		'----刪除Slave表---
		'---for 知識家---
		if session("ibaseDsd") = "39" then
			sql = "UPDATE KnowledgeForum SET Status = 'D' WHERE gicuitem = " & pkStr(icuitem,"")					
			conn.Execute SQL			
		else
			sql = "DELETE FROM " & nullText(refModel.selectSingleNode("tableName")) & " WHERE gicuitem = " & pkStr(icuitem,"")
						'& " WHERE gicuitem=" & pkStr(request.queryString("icuitem"),"")
			conn.Execute SQL			
		end if		
							
		'---for 知識家的刪除-刪題目或是討論或是意見---
		'if session("ibaseDsd") = "39" then			
		'	if session("ctUnitId") = "932" then '---question---
		'		'---do nothing---因為題目本身已看不到---	
		'		sql = "SELECT gicuitem FROM KnowledgeForum WHERE ParentIcuitem = " & pkStr(icuitem,"")
		'		set qrs = conn.execute(sql)
		'		while not qrs.eof					
		'			sql = "UPDATE KnowledgeForum SET Status = 'D' WHERE ParentIcuitem = " & pkStr(qrs("gicuitem"), "")
		'			conn.execute(sql)					
		'			qrs.movenext
		'		wend				
		'		sql = "UPDATE KnowledgeForum SET Status = 'D' WHERE ParentIcuitem = " & pkStr(icuitem,"")
		'		conn.execute(sql)				
		'	elseif session("ctUnitId") = "933" then '---discuss---				
		'		'---更新此討論的意見為D---
		'		sql = "UPDATE KnowledgeForum SET Status = 'D' WHERE ParentIcuitem = " & pkStr(icuitem,"")
		'		conn.execute(sql)
		'		'---更新parent的count---				
		'		sql = "SELECT * FROM KnowledgeForum WHERE gicuitem = " & pkStr(icuitem,"")
		'		set delrs = conn.execute(sql)
		'		while not delrs.eof 
		'			CommandCount = CInt(delrs("CommandCount"))
		'			GradeCount = CInt(delrs("GradeCount"))
		'			GradePersonCount = CInt(delrs("GradePersonCount"))
		'			ParentIcuitem = delrs("ParentIcuitem")					
		'			sql = "UPDATE KnowledgeForum SET DiscussCount = DiscussCount - 1, CommandCount = CommandCount - " & CommandCount & ", " & _
		'						"GradeCount = GradeCount - " & GradeCount & ", GradePersonCount = GradePersonCount - " & GradePersonCount & " " & _
		'						"WHERE gicuitem = " & pkStr(ParentIcuitem, "")					
		'			conn.execute(sql)					
		'			delrs.movenext					
		'		wend				
		'	elseif session("ctUnitId") = "935" then '---opinion---				
		'		sql = "SELECT ParentIcuitem FROM KnowledgeForum WHERE gicuitem = " & pkStr(icuitem, "")
		'		set ors = conn.execute(sql)
		'		while not ors.eof
		'			parent = ors("ParentIcuitem")
		'			ors.movenext
		'		wend				
		'		sql = "SELECT ParentIcuitem FROM KnowledgeForum WHERE gicuitem = " & pkStr(parent, "")
		'		set gors = conn.execute(sql)
		'		while not gors.eof
		'			grandparent = gors("ParentIcuitem")
		'			gors.movenext
		'		wend				
		'		sql = "UPDATE KnowledgeForum SET CommandCount = CommandCount - 1 WHERE gicuitem = " & pkStr(parent, "")
		'		conn.execute(sql)				
		'		sql = "UPDATE KnowledgeForum SET CommandCount = CommandCount - 1 WHERE gicuitem = " & pkStr(grandparent, "")
		'		conn.execute(sql)				
		'	end if			
		'end if
		'---end of for 知識家的刪除---
		
		'=============	2006.6.7 by Gary
		if checkGIPconfig("Discuss") and nullText(refModel.selectSingleNode("tableName")) = "Discuss" then
			sql_item = "SELECT * FROM " & nullText(refModel.selectSingleNode("tableName")) _
				& " WHERE reply=" & pkStr(icuitem,"")
			Set RSitem = Conn.execute(sql_item)
			while not RSitem.EOF  '=====check reply
				sql = "DELETE FROM CuDtGeneric WHERE icuitem=" & RSitem("gicuitem")
				conn.execute sql
				sql = "DELETE FROM " & nullText(refModel.selectSingleNode("tableName")) _
					& " WHERE gicuitem=" & RSitem("gicuitem")
				conn.Execute SQL
				RSitem.movenext
			wend
		end if
		'=============
		'----刪除關鍵字詞
		sql = "DELETE FROM CuDtkeyword WHERE icuitem=" & pkStr(icuitem,"")
		'sql = "DELETE FROM CuDtkeyword WHERE icuitem=" & pkStr(request.queryString("icuitem"),"")
		conn.execute sql
		'----刪除參照資料
		userId = session("userId")
    SQLRef5List = "Select CDT.icuitem,B.sbaseTableName " _
			& " FROM CatTreeNode AS t " _
			& " LEFT JOIN CtUserSet AS u ON u.ctNodeId=t.ctNodeId " _
			& " Left Join CuDtGeneric CDT ON t.ctUnitId=CDT.ictunit AND CDT.showType='5' AND CDT.refId=" & icuitem _
			& " Left Join BaseDsd B ON CDT.ibaseDsd = B.ibaseDsd " _
			& " WHERE ctRootId = "& session("ItemID") & " AND u.userId=" & pkStr(userId,"") & " AND CDT.icuitem is not null"
		Set RSRef5List=conn.execute(SQLRef5List)
		if not RSRef5List.EOF then
	    while not RSRef5List.EOF
	    	sql="Delete from CuDtGeneric WHERE icuitem=" & RSRef5List("icuitem") & ";"
	    	sql=sql&"Delete from " & RSRef5List("sbaseTableName") & " where gicuitem=" & RSRef5List("icuitem")
	    	conn.execute(sql)
	    	RSRef5List.movenext
	    wend
		end if
		'----刪除hyftd文章索引
		if checkGIPconfig("hyftdGIP") then
			hyftdGIPStr=hyftdGIP("delete",icuitem)
			'hyftdGIPStr=hyftdGIP("delete",request.queryString("icuitem"))
		end if			
		'----940215刪除電影網關聯資料
		if session("FilmRelated_CorpActor")="Y" then
			conn.execute("delete from FilmCorpInfo where FilmNo="&icuitem)
			conn.execute("delete from FilmPeopleInfo where FilmNo="&icuitem)
		end if		
		'----940410 刪除MMO紀錄檔資料
		if checkGIPconfig("MMOFolder") then	
			conn.execute("delete from MMOReferened where MMOID=" & icuitem)
		end if		
		
		showDoneBox("資料刪除成功！")
		
	else
		EditInBothCase()
	end if
	'========	2006.5.25 by Gary
	session("BatchDicuitem") = ""
	session("BatchDphase") = ""
	session("BatchDsubmitTask") = ""
'========	
sub EditInBothCase
	xSelect = "SELECT htx.*, ghtx.*, xrefNFileName.oldFileName AS fxr_fileDownLoad,rdsdcat,sbaseTableName " 
	for each param in refModel.selectNodes("//field[xSQL!='']")
		xSelect = xSelect & ", " & nullText(param.selectSingleNode("xSQL")) & " AS " _
			& nullText(param.selectSingleNode("fieldName"))
	next
	sqlCom = xSelect _
		& " FROM " & nullText(refModel.selectSingleNode("tableName")) _
		& " AS htx JOIN CuDtGeneric AS ghtx ON ghtx.icuitem=htx.gicuitem "_
		& " LEFT JOIN ImageFile AS xrefNFileName ON xrefNFileName.newFileName = ghtx.fileDownLoad " _
		& " Left Join BaseDsd B ON ghtx.ibaseDsd=B.ibaseDsd " _
		& " WHERE ghtx.icuitem=" & pkStr(request.queryString("icuitem"),"")		
	Set RSreg = Conn.execute(sqlcom)
	if checkGIPconfig("MMOFolder") and RSreg("rdsdcat")="MMO" then
		SQLM="Select MM.mmositeId+MM.mmofolderName+'/' as MMOFOlderPath " & _
			",(Select count(*) from MMOReferened where MMOID="&pkStr(request.queryString("icuitem"),"") & ") MMOCount " & _
			"from "&RSreg("sbaseTableName")&" CM Left Join Mmofolder MM ON CM.mmofolderId=MM.mmofolderId " & _
			"where CM.gicuitem="&pkStr(request.queryString("icuitem"),"")				
		set RSMMO=conn.execute(SQLM)
		MMOPath=session("MMOPublic") & RSMMO("MMOFOlderPath")
		MMOCount=RSMMO("MMOCount")
	end if
	xFrom = nullText(refModel.selectSingleNode("tableName")) 
	if RSreg("showType")="5" then xref5YN="Y"
	pKey = "icuitem=" & RSreg("icuitem")
	'----被參照個數
	userId = session("userId")
    	SQLRef5Count = "Select count(CDT.icuitem) " _
		& " FROM CatTreeNode AS t " _
		& " LEFT JOIN CtUserSet AS u ON u.ctNodeId=t.ctNodeId " _
		& " Left Join CuDtGeneric CDT ON t.ctUnitId=CDT.ictunit AND CDT.showType='5' AND CDT.refId="&request.querystring("icuitem") _
		& " WHERE ctRootId = "& session("ItemID") & " AND u.userId=" & pkStr(userId,"") & " AND CDT.icuitem is not null"
'----- Chris Modified, 2006/08/12, 在試不指定 treeID 的方式，碰到這裡出問題。被參照只限於同顆樹才能參照？Why？adn Why 只能是同人的？---------------------------------------------
    	SQLRef5Count = "Select count(CDT.icuitem) " _
		& " FROM CatTreeNode AS t " _
		& " LEFT JOIN CtUserSet AS u ON u.ctNodeId=t.ctNodeId " _
		& " Left Join CuDtGeneric CDT ON t.ctUnitId=CDT.ictunit AND CDT.showType='5' AND CDT.refId="&request.querystring("icuitem") _
		& " WHERE u.userId=" & pkStr(userId,"") & " AND CDT.icuitem is not null"
'		& " WHERE ctRootId = "& session("ItemID") & " AND u.userId=" & pkStr(userId,"") & " AND CDT.icuitem is not null"
'----- Chris Modified, 2006/08/12 End --------------------------------------------------------------------------		
'	response.Write SQLRef5count & "<HR/>"
'	response.end
	Set RS5Count=conn.execute(SQLRef5Count)
	xRef5Count=RS5Count(0)
'	response.write "xx="&xRef5Count
'	response.end
    if (request.querystring("S")="Edit" and request.querystring("showType")="") or ((request.querystring("S")="" or request.querystring("S")="Approve" or request.querystring("S")="Ref") and RSreg("showType")="1") then	'----@
	set allModel2 = session("codeXMLSpec2").documentElement   
	for each param in allModel2.selectNodes("//fieldList/field[showTypeStr='4']") 
		set romoveNode=allModel2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		allModel2.removeChild romoveNode
	next
'response.write "<XMP>"+allModel2.xml+"</XMP>"
'response.end	
  	xshowTypeStr="一般資料式"
  	xshowType="1"
    elseif (request.querystring("S")="Edit" and request.querystring("showType")="2") or ((request.querystring("S")="" or request.querystring("S")="Approve" or request.querystring("S")="Ref") and RSreg("showType")="2") then	'----URLs
    	'----Load XSL樣板
	set allModel2 = session("codeXMLSpec2").documentElement
	for each param in allModel2.selectNodes("//fieldList/field[showTypeStr='4']")
		set romoveNode=allModel2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		allModel2.removeChild romoveNode
	next
  	xshowTypeStr="URL連結式"
  	xshowType="2"
    elseif (request.querystring("S")="Edit" and request.querystring("showType")="3") or ((request.querystring("S")="" or request.querystring("S")="Approve" or request.querystring("S")="Ref") and RSreg("showType")="3") then	'----U
    	'----Load XSL樣板
	set allModel2 = session("codeXMLSpec2").documentElement
	for each param in allModel2.selectNodes("//fieldList/field[showTypeStr='4']")
		set romoveNode=allModel2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		allModel2.removeChild romoveNode
	next
 	xshowTypeStr="檔案下載式"
  	xshowType="3"
    elseif ((request.querystring("S")="" or request.querystring("S")="Approve" or request.querystring("S")="Ref") and (RSreg("showType")="4" or RSreg("showType")="5")) then	'----
    	'----Load XSL樣板
  	set htPageDom = session("codeXMLSpec")
    	set oxsl2 = server.createObject("microsoft.XMLDOM")
   	oxsl2.async = false
   	xv = oxsl2.load(server.mappath("/GipDSD/xmlspec/CtUnitXOrder.xsl"))    	
	set nxml2 = server.createObject("microsoft.XMLDOM")
	nxml2.LoadXML(htPageDom.transformNode(oxsl2))
	set allModel2 = nxml2.documentElement  
	for each param in allModel2.selectNodes("//fieldList/field[showTypeStr='3']") 
		set romoveNode=allModel2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		allModel2.removeChild romoveNode
	next
	if RSreg("showType")="4" then
  		xshowTypeStr="引用資料式(複製)"
  	else
  		xshowTypeStr="引用資料式(參照)"  	
  	end if
  	xshowType=RSreg("showType")
    end if

'response.write "<XMP>"+allModel2.xml+"</XMP>"
'response.end
	'----940822 GIP多站間多向出版 圖檔連結修改
	if checkGIPconfig("CrossSitePublishReceive") then
		set CSPFromSiteDom = Server.CreateObject("MICROSOFT.XMLDOM")
		CSPFromSiteDom.async = false
		CSPFromSiteDom.setProperty("ServerHTTPRequest") = true

		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\CSPFromSet.xml"
		xv = CSPFromSiteDom.load(LoadXML)
		if CSPFromSiteDom.parseError.reason <> "" then
			Response.Write("CSPFromSiteDom parseError on line " &  CSPFromSiteDom.parseError.line)
			Response.Write("<BR>Reason: " &  CSPFromSiteDom.parseError.reason)
			Response.End()
		end if
		if not isNull(RSreg("fsiteID")) then _
			HTUploadPath = nullText(CSPFromSiteDom.selectSingleNode("CrossSitePublish/FromSiteList/FromSite[siteID='"&RSreg("fsiteID")&"']/ImageURL"))

	end if

	showHTMLHead()
	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
'	response.write sqlcom
	showForm()
	initForm()
	showHTMLTail()
end sub


function qqRS(fldName)
'on error resume next
	if request("submitTask")="" then
		xValue = RSreg(fldname)
	else
		xValue = ""
		if request("htx_"&fldName) <> "" then
			xValue = request("htx_"&fldName)
		end if
	end if
'	if err.number > 0 then	
'		response.write "***"&fldName&"***"&Err.Description&"***"
'		err.number = 0
'		Err.Clear
'	end if 
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
'		xp = instr(xqqRS,vbCRLF&vbCRLF)
'		while xp > 0
'			xqqRS = left(xqqRS,xp-1) & mid(xqqRS,xp+4)
'			xp = instr(xqqRS,vbCRLF&vbCRLF)
'		wend
		xqqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
		xqqRS = replace(xqqRS,chr(13),"")
		xqqRS = trim(xqqRS)
		qqRS = replace(xqqRS,chr(10),"")
	end if
end function %>

<%Sub initForm() %>
<script language=vbs>
cvbCRLF = vbCRLF
cTabchar = chr(9)

Dim CanTarget
Dim followCanTarget

<% if session("documentDomain") <> "" then %>                        
 document.domain = "<%=session("documentDomain")%>"
<% end if %>

<%
	for each xCode in allModel.selectNodes("scriptCode")
		response.write replace(replace(xCode.text,chr(10),chr(13)&chr(10)),"zzzzzzmySiteURL",session("mySiteMMOURL"))
	next
%>

sub popCalendar(dateName,followName)        
 	CanTarget=dateName
 	followCanTarget=followName
	xdate = document.all(CanTarget).value
	if not isDate(xdate) then	xdate = date()
	document.all.calendar1.setDate xdate
	
 	If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 	Else        
       ex=window.event.clientX + 100
'       ey=document.body.scrolltop+window.event.clientY+10
       ey = window.event.srcElement.parentElement.offsetTop
      if ex>520 then ex=520
       document.all.calendar1.style.pixelleft=ex-80
       document.all.calendar1.style.pixeltop=ey
       document.all.calendar1.style.visibility=""
 	End If              
end sub     

sub calendar1_onscriptletevent(n,o) 
    	document.all("CalendarTarget").value=n         
    	document.all.calendar1.style.visibility="hidden"    
    if n <> "Cancle" then
        document.all(CanTarget).value=document.all.CalendarTarget.value
        document.all("pcShow"&CanTarget).value=document.all.CalendarTarget.value
        if followCanTarget<>"" then
        	document.all(followCanTarget).value=document.all.CalendarTarget.value
        	document.all("pcShow"&followCanTarget).value=document.all.CalendarTarget.value
        end if
	end if
end sub   

function ncTabChar(n)
	if cTabchar = "" then
		ncTabChar = ""
	else
		ncTabChar = String(n,cTabchar)
	end if
end function

    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
    end sub

      sub resetForm 
	       reg.reset()
	       clientInitForm
      end sub

     sub clientInitForm
	reg.showType.value="<%=xshowType%>"     
<%
'	if xshowType <> "5" then
	    for each param in allModel2.selectNodes("//fieldList/field[formList!='']") 
	    	if not (nullText(param.selectSingleNode("fieldName"))="fctupublic" and session("CheckYN")<>"N") then
			EditProcessInit param
	    	end if
	    next
'	else
'	  for each param in allModel2.selectNodes("//fieldList/field[formList!='']") 
'	   if not (nullText(param.selectSingleNode("fieldName"))="fctupublic" and session("CheckYN")="Y") then
'	    if nullText(param.selectSingleNode("inputType"))<>"hidden" and nullText(param.selectSingleNode("fieldRefEditYN"))="N" then
'		orgInputType = param.selectSingleNode("inputType").text
'		param.selectSingleNode("inputType").text = "showType5"
'		EditProcessInit param
'		param.selectSingleNode("inputType").text = orgInputType			    	
'	    else
'		EditProcessInit param
'	    end if
'	   end if
'	  next
'	end if
%>
    end sub

    sub EditRefParent()
    	window.navigate "DsdXMLEdit.asp?S=Ref&icuitem=<%=RSreg("refId")%>"
    end sub

    sub window_onLoad
         clientInitForm
    end sub
    
    sub initRadio(xname,value)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    end sub

    sub initOtherRadio(xname,value, otherName)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    	if value="" then	exit sub
		reg.all(xname).item(reg.all(xname).length-1).checked = true
		reg.all(otherName).value = value
		reg.all(xname).item(reg.all(xname).length-1).value = value
    end sub

    sub initCheckbox(xname,ckValue)
		dim a
		a=Split(ckValue,",")

    	'value = ckValue & ","
    	for each xri in reg.all(xname)
    		'if instr(value, xri.value&",") > 0 then 
    		'	xri.checked = true
    		'end if
			For i=0 to UBound(a)
				if a(i) = xri.value then
					xri.checked = true
				end if
			Next
    	next
    end sub
    
    sub initOtherCheckbox(xname,ckValue,otherName)
    	valueArray = split(ckValue,", ")
    	valueCount = ubound(valueArray) + 1
    	value = ckValue & ","
    	ckCount = 0
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    			ckCount = ckCount + 1
    		end if
    	next
		if ckCount <> valueCount then
			reg.all(xname).item(reg.all(xname).length-1).checked = true
			reg.all(otherName).value = valueArray(ubound(valueArray))
			reg.all(xname).item(reg.all(xname).length-1).value = valueArray(ubound(valueArray))
		end if
    end sub

    sub initImgFile(xname, value)
		reg.all("htImgActCK_"&xname).value=""
		reg.all("htImg_"&xname).style.display="none"
		reg.all("hoImg_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
	   		document.all("logo_"&xname).src= "<%=HTUploadPath%>" & value
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
	end sub
    sub initMMOFile(xname, value)
		reg.all("htMMOActCK_"&xname).value=""
		reg.all("htMMO_"&xname).style.display="none"
		reg.all("htMMO_"&xname).value = value
		reg.all("hoMMO_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
			xValuePos=instr(value,".")
			if instr(value,".")<>0 then 
	   		    if ucase(mid(value,xValuePos+1))="JPG" or ucase(mid(value,xValuePos+1))="JPEG" or ucase(mid(value,xValuePos+1))="BMP" or ucase(mid(value,xValuePos+1))="GIF" then 
	   			document.all("logo_"&xname).src= "<%=MMOPath%>" & value
	   		    else
	   			document.all("logo_"&xname).style.display="none"
	   		    end if
	   		end if
	   		document.all("filename_"&xname).innerText= value
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
    end sub
Sub addLogo(xname)	'新增logo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="addLogo"
End sub

Sub chgLogo(xname)	'更換logo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="editLogo"
End sub

Sub orgLogo(xname)	'原圖
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value=""
End sub

Sub delLogo(xname)	'刪除logo
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("noLogo_"&xname).style.display=""
    reg.all("htImg_"&xname).value=""
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value="delLogo"
End sub

    sub initAttFile(xname, value, orgValue)
		reg.all("htFileActCK_"&xname).value=""
		reg.all("htFile_"&xname).style.display="none"
		reg.all("hoFile_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
	   		document.all("logo_"&xname).innerText= orgValue
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
	end sub

Sub addXFile(xname)	'新增logo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="addLogo"
End sub

Sub chgXFile(xname)	'更換logo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="editLogo"
End sub

Sub orgXFile(xname)	'原圖
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value=""
End sub

Sub delXFile(xname)	'刪除logo
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("noLogo_"&xname).style.display=""
    reg.all("htFile_"&xname).value=""
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value="delLogo"
End sub
Sub addMMO(xname)	'新增MMO
    reg.all("htMMO_"&xname).style.display=""
    reg.all("htMMOActCK_"&xname).value="addLogo"
End sub

Sub chgMMO(xname)	'更換MMO
    reg.all("htMMO_"&xname).style.display=""
    reg.all("htMMOActCK_"&xname).value="editLogo"
End sub

Sub orgMMO(xname)	'原MMO
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    if instr(document.all("logo_"&xname).src,".")<>0 then _
    	document.all("logo_"&xname).style.display=""
    document.all("filename_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htMMO_"&xname).style.display="none"
    reg.all("htMMOActCK_"&xname).value=""
End sub

Sub delMMO(xname)	'刪除MMO
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("filename_"&xname).style.display="none"    
    document.all("noLogo_"&xname).style.display=""
    reg.all("htMMO_"&xname).value=""
    reg.all("htMMO_"&xname).style.display="none"
    reg.all("htMMOActCK_"&xname).value="delLogo"
End sub
sub keywordComp(xBodyNodeStr)
	if reg.htx_xKeyword.value = "" then
		reg.htx_xKeyword.value = xBodyNodeStr
	else	
		i=0
		redim compArray(1,0)
		xKeywordArray=split(reg.htx_xKeyword.value,",")
		xBodyArray=split(xBodyNodeStr,",")
		'----將reg.htx_xKeyword.value拆成字串與權重存入合併陣列中
		for i=0 to ubound(xKeywordArray)
		    xPos=instr(xKeywordArray(i),"*")
		    if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
			xWeight=mid(xKeywordArray(i),xPos+1)
		    else
			xStr=trim(xKeywordArray(i))
			xWeight=""
		    end if	
		    redim preserve compArray(1,i)
		    compArray(0,i)=xStr
		    compArray(1,i)=xWeight
		next	
		'----將hysearch傳回之內文斷詞字串陣列拿來與reg.htx_xKeyword.value陣列逐一比較,若不存在則存入合併陣列中
		for j=0 to ubound(xBodyArray)
		    insertFlag=true
		    xPos2=instr(xBodyArray(j),"*")
		    if xPos2<>0 then
			xStr2=Left(trim(xBodyArray(j)),xPos2-1)
			xWeight2=mid(xBodyArray(j),xPos2+1)
		    else
			xStr2=trim(xBodyArray(j))
			xWeight2=""
		    end if			    
		    for k=0 to ubound(compArray,2)		    
			if xStr2=compArray(0,k) then
		    	    compArray(1,k)=xWeight2			    
			    insertFlag=false
			    exit for
			end if
		    next
		    if insertFlag then
			redim preserve compArray(1,i)
		    	compArray(0,i)=xStr2
		    	compArray(1,i)=xWeight2
			i=i+1
		    end if
					
		next
		'----將合併陣列串成最後字串
		xKeywordStr=""
		for w=0 to ubound(compArray,2)
		    if compArray(1,w)<>"" then
		    	xKeywordStr=xKeywordStr+compArray(0,w)+"*"+compArray(1,w)+","
		    else
		    	xKeywordStr=xKeywordStr+compArray(0,w)+","
		    end if
		next	
		reg.htx_xKeyword.value = Left(xKeywordStr,Len(xKeywordStr)-1)	
	end if	
end sub
 
sub keywordMake()
	if reg.htx_xBody.value="" then exit sub
	xStr=replace(replace(reg.htx_xBody.value,VBCRLF,""),"&nbsp;","")
	xmlStr = "<?xml version=""1.0"" encoding=""utf-8"" ?>" & cvbCRLF
	xmlStr = xmlStr & "<xBodyList>" & cvbCRLF
	xmlStr = xmlStr & ncTabchar(1) & "<xBody><![CDATA[" & xStr & "]]></xBody>" & cvbCRLF
	xmlStr = xmlStr & "</xBodyList>" & cvbCRLF
  	set oXmlReg = createObject("Microsoft.XMLDOM")
  	oXmlReg.async = false
  	oXmlReg.loadXML xmlStr
	if oXmlReg.parseError.reason <> "" then 
		alert "內文不符合字串比較格式!"
		exit sub
	end if  	
  	postURL = "<%=session("mySiteMMOURL")%>/ws/ws_xBody.asp"  	
	set xmlHTTP = CreateObject("MSXML2.XMLHTTP")
	set rXmlObj = CreateObject("MICROSOFT.XMLDOM")
	rXmlObj.async = false
	xmlHTTP.open "POST", postURL, false
	xmlHTTp.send oXmlReg.xml  
	rv = rXmlObj.load(xmlHTTP.responseXML)		
	if not rv then
		alert "關鍵字詞傳回出現錯誤!"
		exit sub
	end if
	set xBodyNode = rXmlObj.selectSingleNode("//xBody")
	if xBodyNode.text<>"" then keywordComp xBodyNode.text
end sub
   
<%  if checkGIPconfig("KeywordAutoGen") then %>
sub htx_xKeyword_onChange()
	'----檢查字詞需一個字元以上
	xKeywordArray=split(reg.htx_xKeyword.value,",")
	for i=0 to ubound(xKeywordArray)
		'----去除權重符號
		xPos=instr(xKeywordArray(i),"*")
		if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
		else
			xStr=trim(xKeywordArray(i))
		end if
		if blen(xStr)<=2 then
			alert "每一關鍵字詞長度至少二個字!"
			exit sub
		end if
	next	
	'----檢查完成
	set oXML = createObject("Microsoft.XMLDOM")
	oXML.async = false
	xURI = "<%=session("mySiteMMOURL")%>/ws/ws_keyword.asp?xKeyword=" & B5toUTF8(reg.htx_xKeyword.value)
	oXML.load(xURI)	
	set xKeywordNode = oXML.selectSingleNode("xKeywordList/xKeywordStr")
	if xKeywordNode.text<>"" then KeywordWTP xKeywordNode.text
end sub
<% end if %>

sub KeywordWTP(xKeyworWTPdStr)
	set oXML = createObject("Microsoft.XMLDOM")
	oXML.async = false
	xStr=""
	xReturnValue=""
	xKeywordWTPArray=split(xKeyworWTPdStr,";")
	for i=0 to ubound(xKeywordWTPArray)
            chky=msgbox("注意！"& vbcrlf & vbcrlf &"　[ "&xKeywordWTPArray(i)&" ]關鍵字詞不存在, 加入待審清單中嗎？"& vbcrlf , 48+1, "請注意！！")
            if chky=vbok then
		xURI = "<%=session("mySiteMMOURL")%>/ws/ws_keywordWTP.asp?xKeyword=" & B5toUTF8(xKeywordWTPArray(i))
		oXML.load(xURI)		          		
       	    end If
	next
end sub 


	clientInitForm
 </script>	
<script language=javascript>
function B5toUTF8(x){
	return encodeURI(x);
}
</script>
<%End sub '---- initForm() ----%>


<%Sub showForm()  	'===================== Client Side Validation Put HERE =========== %>
    <script language="vbscript">
      sub formModSubmit()
    
'----- 用戶端Submit前的檢查碼放在這裡 ，如下例 not valid 時 exit sub 離開------------
'msg1 = "請務必填寫「客戶編號」，不得為空白！"
'msg2 = "請務必填寫「客戶名稱」，不得為空白！"
'
'  If reg.tfx_ClientName.value = Empty Then
'     MsgBox msg2, 64, "Sorry!"
'     settab 0
'     reg.tfx_ClientName.focus
'     Exit Sub
'  End if
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
  
<%
	'for each param in allModel2.selectNodes("//fieldList/field[formList!='' and inputType!='hidden']") 
	    'if not (nullText(param.selectSingleNode("fieldName"))="fctupublic" and session("CheckYN")<>"N") then
		'processValid param
	    'end if
	'next
%>
  
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub

   xRef5Count="<%=xRef5Count%>"
   sub formDelSubmit()
   	deleteStr = ""
   	if xRef5Count > 0 then deleteStr = deleteStr & "   此筆資料有被其他資料參照, 刪除此筆資料將會一併刪除參照資料!" & vbcrlf & vbcrlf 
   	if xMMOCount > 0 then deleteStr = deleteStr & "   此筆資料為多媒體物件且已被引用, 刪除後將造成引用資料的內容不完整!" & vbcrlf & vbcrlf 
   	deleteStr = deleteStr & "　確定刪除資料嗎？"
	chky=msgbox("注意！"& vbcrlf & vbcrlf &deleteStr& vbcrlf , 48+1, "請注意！！")
        if chky=vbok then
		for i=0 to reg.elements.length-1
               	    set e=document.reg.elements(i)
               	    if left(e.name,11)="htImgActCK_" or left(e.name,12)="htFileActCK_" or left(e.name,11)="htMMOActCK_" then 
               	    	e.value="delLogo"
               	    end if
          	next          
		reg.submitTask.value = "DELETE"
	      	reg.Submit
       	end If
  end sub
  sub CuCheckPrint()
'	window.open "<%=session("myWWWSiteURL")%>/ContentOnly.asp?CuItem=<%=RSreg("icuitem")%>&ItemID=<%=session("ItemID")%>&userId=<%=session("userId")%>"
	window.open "<%=session("myWWWSiteURL")%>/fp.asp?CuItem=<%=RSreg("icuitem")%>&ItemID=<%=session("ItemID")%>&userId=<%=session("userId")%>"
  end sub
<%if checkGIPconfig("MMOFolder") then %>
sub MMOFolderAdd()
	window.open "/MMO/MMOAddFolder.asp?MMOFolderID="&reg.htx_MMOFolderID.value,"","height=400,width=550"
end sub
sub MMOFolderAddReload()
	document.location.reload()
end sub
<%end if%>
</script>

<!--#include file="DsdXMLFormE.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>

<body>
<div id="FuncName">
	<h1><%=Title%></h1><font size=2>【目錄樹節點: <%=session("catName")%>】</font>
	<div id="Nav">
<% if nullText(htPageDom.selectSingleNode("//xfLinkList/anchor/anchorURI"))<>"" then 
	for each xfLink in htPageDom.selectNodes("//xfLinkList/anchor")
		response.write "<A href=""" & nullText(xfLink.selectSingleNode("anchorURI")) _
			& "&xNode=" & session("ctNodeId") & "&" & pKey & """>"
		xfStr = nullText(xfLink.selectSingleNode("anchorLabel"))
		for each lableReplace in xfLink.selectNodes("lableReplace")
			rpStr = nullText(lableReplace.selectSingleNode("rpStr"))
			if rpStr<>"" then
				xfStr = replace(xfStr,rpStr,RSreg(nullText(lableReplace.selectSingleNode("rpValue"))))
			end if
		next
		response.write xfStr & "</A>"
	next %>
   </div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">
<%
	response.write nullText(htPageDom.selectSingleNode("//field[fieldName='stitle']/fieldLabel"))
	response.write "：" & RSreg("stitle") & "　>"
	response.write nullText(htPageDom.selectSingleNode("//xfLinkList/anchor/anchorLabel"))
%>
</div>
<% else %>
	   <%if xshowType <> "5" then%>
	       <%if (HTProgRight and 128)=128 then
	       	if checkGIPconfig("DsdXMLAddMulti") then
	       	
	       	else
	       %>	       
           <!--<a href="DsdXMLAdd_Multi.asp?<%=pKey%>&showType=4">多向出版(複製)</a>&nbsp;-->
	       <%
	       end if
	       end if
	       if checkGIPconfig("DsdXMLAddMulti") then	       	
	       else
		%>	    
	       <!--<a href="DsdXMLAdd_Multi.asp?<%=pKey%>&showType=5">多向出版(參照)</a>&nbsp;-->
	       <%
	       end if
	       if (HTProgRight and 8)=8 then%>
	       	    <%if RSreg("fctupublic")="P" then%>
	       		<a href="Javascript: CuCheckPrint();">列印待審稿</a>&nbsp;
	            <%end if%>
	       <a href="CuHtmlEdit.asp?<%=pKey%>">網頁</a>&nbsp;
	       <% 'if session("CodeID")="7" then %>
	       <a href="CuPageList.asp?<%=pKey%>">連結</a>&nbsp;
	       <%'end if%>
	       <a href="CuAttachList.asp?<%=pKey%>">附件</a>&nbsp;
	       <%end if%>
	   <%end if%>
	       
	       <%if (HTProgRight and 2)=2 then%>
	       <a href="<%=HTprogPrefix%>Query.asp">查詢</a>&nbsp;
	       <%end if%>	    
	       <%if (HTProgRight and 4)=4 then%>
	       <a href="<%=HTProgPrefix%>Add.asp?phase=add">新增</a>&nbsp;
	       <%end if%>
	       <a href="DsdXMLList.asp?iBaseDSD=<%=session("iBaseDSD")%>">回條列</a>
		   
		   <!-- -------------------------------------------------------------------->
		   <%
		   		sql = "SELECT n.* FROM CuDtGeneric AS n " _
				& " WHERE n.icuitem=" & request.querystring("icuitem")
	    		set RS = Conn.execute(sql)
				topCat = RS("topCat")
				session("ictunit") = RS("ictunit")
				cuitem = request.querystring("icuitem")
				sql = "SELECT dbo.GET_SYNC_LOCAL_UNIT("&RS("ictunit")&") AS XX"
	    		set RS = Conn.execute(sql)
		   %>
		   <% IF RS("XX") > "-1" THEN %>
		   		<a href="ChangeUnit.asp?ispostback=0&topCat=<%=topCat%>&icuitem=<% =cuitem %>&ictunit=<% =session("ictunit")%>&iBaseDSD=<%=session("iBaseDSD")%>" onClick="changeUnit_onClick()">變更主題單元</a>
		   <% END IF %>
		   <!----------------------------------------------------------------------->
		   
		   <%if session("ctNodeId")  = "1582" then%>
		   <a href="publishset.asp?questionId=<%=request.queryString("icuitem")%>">發佈至主題館</a>
		   <%end if%>
	       <% If Trim(refModel.selectSingleNode("tableDesc").text)="題庫試題範本" Then%>
	       <a href="../hakkalangExam/topic_edit.asp?<%=pKey%>">題庫</a>&nbsp;
	       <%end if%>
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">
	<% if xshowType <> "5" then %>
	       <%if (HTProgRight and 8)=8 then%>
	       <input type="button" value="一般資料式" onclick="window.location.href='<%=HTprogPrefix%>Edit.asp?S=Edit&icuitem=<%=request.queryString("icuitem")%>&phase=edit'"></input>
	       <input type="button" value="URL連結式" onclick="window.location.href='<%=HTprogPrefix%>Edit.asp?S=Edit&icuitem=<%=request.queryString("icuitem")%>&showType=2&phase=edit'"></input>
	       <input type="button" value="檔案下載式" onclick="window.location.href='<%=HTprogPrefix%>Edit.asp?S=Edit&icuitem=<%=request.queryString("icuitem")%>&showType=3&phase=edit'"></input>
	       <%end if%><br>
       	<%end If %>     	       	
	<%=HTProgCap%>&nbsp;
    <font size=2>【編輯(<font color=red><%=xshowTypeStr%></font>)--主題單元:<%=session("CtUnitName")%> / 單元資料:<%=nullText(htPageDom.selectSingleNode("//tableDesc"))%>】
</div>
<% end if %>
<%End sub '--- showHTMLHead() ------%>


<%Sub ShowHTMLTail() %>     
</body>
</html>
<%End sub '--- showHTMLTail() ------%>


<%Sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
'  			window.history.back
	</script>
<%
    End sub '---- showErrBox() ----

Sub checkDBValid()	'===================== Server Side Validation Put HERE =================

'---- 後端資料驗證檢查程式碼放在這裡	，如下例，有錯時設 errMsg="xxx" 並 exit sub ------
'	if Request("tfx_TaxID") <> "" Then
'	  SQL = "Select * From Client Where TaxID = '"& Request("tfx_TaxID") &"'"
'	  set RSreg = conn.Execute(SQL)
'	  if not RSreg.EOF then
'		if trim(RSreg("ClientId")) <> request("pfx_ClientID") Then
'			errMsg = "「統一編號」重複!!請重新輸入客戶統一編號!"
'			exit sub
'		end if
'	  end if
'	end if

End sub '---- checkDBValid() ----

function d6date(dt)     '轉成民國年  999/99/99 給資料型態為SmallDateTime 使用
	if Len(dt)=0 or isnull(dt) then
	     d6date=""
	else
                      xy=right("000"+ cstr((year(dt)-1911)),2)     '補零
                      xm=right("00"+ cstr(month(dt)),2)
                      xd=right("00"+ cstr(day(dt)),2)
                      d6date=xy &  xm &  xd
                  end if
end function

Sub doUpdateDB()
	'----940407 MMO上傳路徑/ftp等參數處理
	SQLCheck="Select rdsdcat,sbaseTableName from CtUnit C Left Join BaseDsd B ON C.ibaseDsd=B.ibaseDsd " & _
		"where C.CtUnitId='"&session("CtUnitID")&"'"
	Set RSCheck=conn.execute(SQLCheck)	
	if checkGIPconfig("MMOFolder") and RSCheck("rDSDCat")="MMO" then
		'----FTP所需參數
		SQLM="Select MM.mmositeId+MM.mmofolderName as MMOFOlderID,MM.mmositeId,MM.mmofolderName,MS.upLoadSiteFtpip,MS.upLoadSiteFtpport,MS.upLoadSiteFtpid,MS.upLoadSiteFtppwd " & _
			" from "&RSCheck("sbaseTableName")&" CM Left Join Mmofolder MM ON CM.mmofolderId=MM.mmofolderId " & _
			" Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId " & _
			"where CM.gicuitem="&request.querystring("icuitem")
		Set RSM=conn.execute(SQLM)
		if not RSM.eof then 
			xMMOFolderID=RSM("MMOFOlderID")
	   		xFTPIPMMO = RSM("upLoadSiteFtpip")
	   		xFTPPortMMO = RSM("upLoadSiteFtpport")
	   		xFTPIDMMO = RSM("upLoadSiteFtpid")
	   		xFTPPWDMMO = RSM("upLoadSiteFtppwd")
			MMOFTPfilePath="public/"&xMMOFolderID
		end if
		'----上傳路徑
		MMOPath = session("MMOPublic") & xMMOFolderID
		if right(MMOPath,1)<>"/" then	MMOPath = MMOPath & "/"
	end if
	'----940407 MMO上傳路徑/ftp等參數處理完成

		'----940215 Film關聯處理
		if session("FilmRelated_CorpActor")="Y" then
      xProductionCompanyA="":xDistributorsA="":xDirector="":xScriptwriter="":xProducer="":xActor="":xskill="":xArt="":xOthers=""
    	SQLF="Select * from FilmInformation where gicuitem=" & pkStr(request.queryString("icuitem"),"")
    	Set RSF=conn.execute(SQLF)
    	if not RSF.eof then
    	    if not isnull(RSF("ProductionCompanyA")) then xProductionCompanyA=RSF("ProductionCompanyA")
    	    if not isnull(RSF("DistributorsA")) then xDistributorsA=RSF("DistributorsA")
    	    if not isnull(RSF("Director")) then xDirector=RSF("Director")
    	    if not isnull(RSF("Scriptwriter")) then xScriptwriter=RSF("Scriptwriter")
    	    if not isnull(RSF("Producer")) then xProducer=RSF("Producer")
    	    if not isnull(RSF("Actor")) then xActor=RSF("Actor")
    	    if not isnull(RSF("skill")) then xskill=RSF("skill")
    	    if not isnull(RSF("Art")) then xArt=RSF("Art")
    	    if not isnull(RSF("Others")) then xOthers=RSF("Others")
			end if
    	if xUpForm("htx_ProductionCompanyA")<>xProductionCompanyA then FilmRelatedStr=FilmRelated("edit","Corp","s@~",request.queryString("icuitem"),xUpForm("htx_ProductionCompanyA"))
    	if xUpForm("htx_DistributorsA")<>xDistributorsA then FilmRelatedStr=FilmRelated("edit","Corp","o~",request.queryString("icuitem"),xUpForm("htx_DistributorsA"))
    	if xUpForm("htx_Director")<>xDirector then FilmRelatedStr=FilmRelated("edit","Actor","1",request.queryString("icuitem"),xUpForm("htx_Director"))
			if xUpForm("htx_Scriptwriter")<>xScriptwriter then FilmRelatedStr=FilmRelated("edit","Actor","2",request.queryString("icuitem"),xUpForm("htx_Scriptwriter"))
			if xUpForm("htx_Producer")<>xProducer then FilmRelatedStr=FilmRelated("edit","Actor","3",request.queryString("icuitem"),xUpForm("htx_Producer"))
			if xUpForm("htx_Actor")<>xActor then FilmRelatedStr=FilmRelated("edit","Actor","4",request.queryString("icuitem"),xUpForm("htx_Actor"))
			if xUpForm("htx_skill")<>xskill then FilmRelatedStr=FilmRelated("edit","Actor","6",request.queryString("icuitem"),xUpForm("htx_skill"))
			if xUpForm("htx_Art")<>xArt then FilmRelatedStr=FilmRelated("edit","Actor","7",request.queryString("icuitem"),xUpForm("htx_Art"))
			if xUpForm("htx_Others")<>xOthers then FilmRelatedStr=FilmRelated("edit","Actor","8",request.queryString("icuitem"),xUpForm("htx_Others"))
		end if
		'----940215 Film關聯處理結束
    
    '----只有一般資料/參照引用方式才要處理slave資料	
    if xUpForm("showType")="1" or xUpForm("showType")="4" or xUpForm("showType")="5" then		'----1@/45    
			xn = 0
			sql = ""
			for each param in refModel.selectNodes("fieldList/field[formList!='']") 
				processUpdate param
				xn = xn + 1
			next

			if sql<>"" then
				sql = "UPDATE " & nullText(refModel.selectSingleNode("tableName")) & " SET " & sql
				sql = left(sql,len(sql)-1) & " WHERE gicuitem=" & pkStr(request.queryString("icuitem"),"")
				'response.write SQL
				'response.End
				conn.Execute(SQL)
			end if
    end if

		'---檢查vGroup是否有值,若有值,代表已加過分---	
		Dim oldvGroup
		sql = "SELECT ISNULL(vGroup,'') AS vGroup FROM CuDTGeneric WHERE icuitem = " & pkStr(request.queryString("icuitem"),"")
		Set ActUpdateRs = conn.execute(sql)
		if not ActUpdateRs.Eof then
			oldvGroup = ActUpdateRs("vGroup")
		end if
		Set ActUpdateRs = nothing		

    '----CuDtGeneric處理
		mailFlag=false	
		sql = "UPDATE CuDtGeneric SET showType=N'"+xUpForm("showType")+"', "
		for each param in allModel.selectNodes("dsTable[tableName='CuDtGeneric']/fieldList/field[formList!='' and identity!='Y']") 
			if nullText(param.selectSingleNode("fieldName")) = "xImportant" and xUpForm("xxCheckImportant")="Y" then
				sql = sql & "xImportant=" & pkStr(d6date(date()),",")
			elseif nullText(param.selectSingleNode("fieldName"))="fctupublic" and session("CheckYN")<>"N" then	
				sql = sql & "fctupublic='P',"
				mailFlag=true
			'xpostDate 是smalldatetime格式 修改成只剩下時間
			elseif nullText(param.selectSingleNode("fieldName"))="xpostDate" then	
				sql = sql & "xpostDate='"& DateValue(xUpForm("htx_xpostDate")) &"',"	
			elseif nullText(param.selectSingleNode("fieldName"))="Created_Date" then
				sql = sql & "Created_Date='"& DateValue(xUpForm("htx_Created_Date")) &"',"
			elseif nullText(param.selectSingleNode("fieldName"))="deditDate" then
				sql = sql & "deditDate='"& DateValue(xUpForm("htx_deditDate")) &"',"			
			else
				processUpdate param
			end if
		next
		sql = left(sql,len(sql)-1) & " WHERE icuitem=" & pkStr(request.queryString("icuitem"),"")
		'response.write sql & "<HR>"		
		'response.end
		conn.Execute(SQL)
        
        '----Added By Leo   2011-08-04  Start
        '----用農業百年的itemID去判斷，如有更好的判斷方式即可修改。
        if session("itemID") = session("農業百年發展史RootId") then
            strUpdateScript = "UPDATE CuDtGeneric SET refID = " & session("ctNodeId") _
                              & "WHERE icuitem =" & pkStr(request.QueryString("icuitem"),"")
            conn.Execute(strUpdateScript)
        end if 
		'----Added By Leo   2011-08-04   End

		'---vincent : 農委會活動, 若有勾選題目, 幫使用者加分---
		'---知識家的活動---
		Dim ActUpdateFlag
		ActUpdateFlag = false
		if session("ibaseDsd") = "39" then				
			Actsql = "SELECT * FROM Activity WHERE (GETDATE() BETWEEN ActivityStartTime AND ActivityEndTime)"
			set ActRs = conn.execute(Actsql)
			'---it is in the act time---
			if not ActRs.eof then	'---question 單元---								
				ActUpdateFlag = true
			end if
			if session("ctUnitId") = "932" And ActUpdateFlag = true then				
				'---檢查上傳的vGroup值, 
				If xupform("htx_vGroup") = "A" And oldvGroup = "" Then '---題目上線,加分---			
					Call Activity1Action("add", pkStr(request.queryString("icuitem"),""), xupform("htx_ieditor"))					
				ElseIf xupform("htx_vGroup") = "" And oldvGroup = "A" Then '---題目下線,扣分---					
					Call Activity1Action("del", pkStr(request.queryString("icuitem"),""), xupform("htx_ieditor"))
				ElseIf xupform("htx_vGroup") = "A" And oldvGroup = "A" Then '---不做動作---									
				ElseIf xupform("htx_vGroup") = "" And oldvGroup = "" Then	'---不做動作---													
				End If				
			end if						
		end if

'------- 記錄 異動 log -------- start --------------------------------------------------------	
	if checkGIPconfig("UserLogFile") then
		sql = "INSERT INTO userActionLog(loginSID,xTarget,xAction,recordNumber,objTitle) VALUES(" _
			& dfn(session("loginLogSID")) & "'0A','2'," _
			& dfn(request.queryString("icuitem")) _
			& pkstr(xUpForm("htx_sTitle"),")")
		conn.execute sql
	end if	
'------- 記錄 異動 log -------- end --------------------------------------------------------	

    '----參照資料同步更新
    SQLCDG = "Select * from CuDtGeneric where icuitem=" & request.querystring("icuitem")
    set RSCDG = conn.execute(SQLCDG)
    xSlaveTable = refModel.selectSingleNode("tableName").text
    SQLSlave = "Select * from "&xSlaveTable&" where gicuitem = " & request.querystring("icuitem")
    set RSSlave = conn.execute(SQLSlave)    
    SQLRef="Select CDT.icuitem from CuDtGeneric CDT " & _
    		"Where CDT.showType='5' and CDT.refId=" & pkStr(request.queryString("icuitem"),"")
    set RSRef=conn.execute(SQLRef)
    if not RSRef.EOF then
    	while not RSRef.EOF
	    sql = "SELECT n.*,b.sbaseTableName FROM CuDtGeneric AS n LEFT JOIN CtUnit AS u ON u.ctUnitId=n.ictunit" _
		& " Left Join BaseDsd As b ON n.ibaseDsd=b.ibaseDsd" _
		& " WHERE n.icuitem=" & RSRef(0)
	    set RS = Conn.execute(sql)
	    xctUnitId = RS("ictunit")
	    xiBaseDSD = RS("ibaseDsd")
	    if isNull(RS("sBaseTableName")) then
		xsBaseTableName = "CuDTx" & xiBaseDSD
	    else
		xsBaseTableName = RS("sbaseTableName")
	    end if
	    '----更新主表
	    SQLRefUpdateMaster="Update CuDtGeneric SET dEditDate=getdate(),"  
	    for each param in allModel.selectNodes("dsTable[tableName='CuDtGeneric']/fieldList/field[formList!='' and fieldName!='icuitem' and fieldName!='ibaseDsd' and fieldName!='ictunit' and fieldName!='idept' and fieldName!='refId' and fieldName!='dEditDate' and fieldName!='Created_Date']") 	    
	        SQLRefUpdateMaster=SQLRefUpdateMaster&nullText(param.selectSingleNode("fieldName"))&"=N'"&RSCDG(nullText(param.selectSingleNode("fieldName")))&"',"	  
	    next
	    SQLRefUpdateMaster = left(SQLRefUpdateMaster,len(SQLRefUpdateMaster)-1) & " WHERE icuitem=" & RSRef(0)
	    conn.execute(SQLRefUpdateMaster)
  	    '----更新SLave表
	    if xSlaveTable = xsBaseTableName then
		SQLRefUpdateSlave = "Update  " & xsBaseTableName & " set gicuitem=gicuitem,"
		for each param in refModel.selectNodes("fieldList/field") 
		    if not isNull(RSSlave(param.selectSingleNode("fieldName").text)) then
	        	SQLRefUpdateSlave=SQLRefUpdateSlave&nullText(param.selectSingleNode("fieldName"))&"=N'"&RSSlave(nullText(param.selectSingleNode("fieldName")))&"',"	  
		    end if
		next
	    	SQLRefUpdateSlave = left(SQLRefUpdateSlave,len(SQLRefUpdateSlave)-1) & " WHERE gicuitem=" & RSRef(0)
'	    	response.write SQLRefUpdateSlave
'	    	response.end
	    	conn.execute(SQLRefUpdateSlave)
	    end if  	    
    	    RSRef.movenext
    	wend
    end if
	'----審稿email處理
		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysPara.xml"
		xv = htPageDom.load(LoadXML)
  		if htPageDom.parseError.reason <> "" then 
	    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    			Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    			Response.End()
  		end if
		set emailnode=htPageDom.selectSingleNode("SystemParameter/DsdXMLEmail")
		xemail=nullText(emailnode)
		xemailStr=nullText(emailnode.selectSingleNode("@xdesc"))	
		fSql="Select IU.userName,IU.email " & _
			" from CuDtGeneric C " & _
			" Left Join CatTreeNode CTN ON C.ictunit=CTN.ctUnitId AND CTN.ctRootId=" & session("ItemID") & _
			" Left Join CtUserSet2 CUS2 ON CTN.ctNodeId=CUS2.ctNodeId AND CUS2.userId=N'"&session("userId")&"' " & _
			" Left Join CtUnit CT On C.ictunit=CT.ctUnitId " & _
			" Left Join InfoUser IU ON CUS2.userId=IU.userId AND C.idept=IU.deptId " & _
			" where C.icuitem=" & pkStr(request.queryString("icuitem"),"")
		set rs1=conn.execute(fSql)
		if not rs1.EOF then
    		    while not rs1.EOF
        	        S_email=""""+xemailStr+""" <"+xemail+">"
        		R_email=rs1(1)
            
        		email_body="【 " & rs1(0) & " 】小姐先生 您好:" & "<br>" & "<br>" & _
                            "   原上稿資料已有編修, 請至[後台管理網站/資料審稿作業中]審核!"& "<br><br>" & _
                            "謝謝您!"& "<br>" & "<br>" & _
                            xemailStr & "<br>" 	
			if not isNull(rs1(1)) then                                
        			Call Send_email(S_email,R_email, "上稿資料審稿通知" ,email_body)  
			end if 
     
     		    	rs1.movenext  
    		    wend
		end if	
		
	'----關鍵字詞處理
	'----先刪除
	SQLDelete="Delete CuDtkeyword where icuitem=" & pkStr(request.queryString("icuitem"),"")
	conn.execute(SQLDelete)
	'----再新增
	if xUpForm("htx_xKeyword")<>"" then
	    redim iArray(1,0)
	    xStr=""
	    xReturnValue=""
	    SQLInsert=""
	    xKeywordArray=split(xUpForm("htx_xKeyword"),",")
	    weightsum=0
	    for i=0 to ubound(xKeywordArray)
	    	redim preserve iArray(1,i)
		'----分開字詞與權重符號
		xPos=instr(xKeywordArray(i),"*")
		if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
			iArray(0,i)=xStr
			iArray(1,i)=mid(xKeywordArray(i),xPos+1)
		else
			xStr=trim(xKeywordArray(i))
			iArray(0,i)=xStr
			iArray(1,i)=1		
		end if	
		weightsum=weightsum+iArray(1,i)
	    next   
	    '----串SQL字串 
	    for k=0 to ubound(iArray,2)
	    	SQLInsert=SQLInsert+"Insert Into CuDtkeyword values("+dfn(request.queryString("icuitem"))+"'"+iArray(0,k)+"',"+cStr(round(iArray(1,k)*100/weightsum))+");"
	    next
	    if SQLInsert<>"" then conn.execute(SQLInsert)
	end if	
	'----更新hyftd文章索引
	if checkGIPconfig("hyftdGIP") then
		hyftdGIPStr=hyftdGIP("update",request.queryString("icuitem"))
	end if	
	'----940410 MMO物件引用紀錄處理
    	if checkGIPconfig("MMOFolder") then
		MMOReferendSQL=""
		if xUpForm("htx_xBody") = "" then
			MMOReferenedSQL="Delete MMOReferened where iCuItem="&pkStr(request.queryString("iCuItem"),"")
		else
			MMOReferenedSQL="Delete MMOReferened where iCuItem="&pkStr(request.queryString("iCuItem"),"")&";"
			MMOReferenedStr=xUpForm("htx_xBody")
			MMORefPos=instr(xUpForm("htx_xBody"),"MMOID=""")
			while MMORefPos > 0
			     MMOReferenedStr=mid(MMOReferenedStr,MMORefPos+7)
			     quotePos=instr(MMOReferenedStr,"""")
			     xMMOID=Left(MMOReferenedStr,quotePos-1)
			     MMOReferenedSQL=MMOReferenedSQL&"Insert Into MMOReferened values("&request.queryString("iCuItem")&","&xMMOID&");"
			     MMORefPos=instr(MMOReferenedStr,"MMOID=""")
			wend
		end if
		if MMOReferenedSQL<>"" then conn.execute(MMOReferenedSQL)
	end if		
	'======	2006.5.8 by Gary
	if checkGIPconfig("RSSandQuery") then  
	
		SQLRSS = "SELECT YNrss FROM catTreeNode WHERE ctNodeId=" & pkStr(session("ctNodeId"),"")
		'response.write sqlrss
		'response.end
		Set RSS = conn.execute(SQLRSS)
		if not RSS.eof and RSS("YNrss")="Y" then
			session("RSS_method") = "update"
			session("RSS_iCuItem") = request.queryString("iCuItem")	
			postURL = "/ws/ws_RSSPool.asp"
			Server.Execute (postURL) 
		end if
	end if
	'======	
'-------- 950703 Chris 擴充萬用外部附加處理 (extraServiceURI) -----begin------------------------
	for each param in allModel.selectNodes("extraServiceURI")
		xURI = session("mySysSiteURL") & "/wsxd/" & nullText(param) & "?iCuItem=" & request.queryString("iCuItem")
'		response.write "Calling " & xURI & "<HR>"
		set srvXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		srvXmlHttp.Open "GET", xURI, false
		srvXmlHttp.Send()
		set srvXmlHttp = nothing	
	next
'	response.end
'-------- 950703 Chris 擴充萬用外部附加處理 (extraServiceURI) ------end-----------------------

End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>編修表單</title>
     </head>
     <body>
     <script language=vbs>
<%if FTPErrorMSG="" and hyftdGIPStr="" then%>
       	    alert("<%=lMsg%>")
<%elseif FTPErrorMSG<>"" and hyftdGIPStr="" then%>
	alert("<%=lMsg%>"+VBCRLF+VBCRLF+"<%=FTPErrorMSG%>")
<%elseif FTPErrorMSG="" and hyftdGIPStr<>"" then%>
	alert("<%=lMsg%>"+VBCRLF+VBCRLF+"<%=hyftdGIPStr%>")
<%elseif FTPErrorMSG<>"" and hyftdGIPStr<>"" then%>
	alert("<%=lMsg%>"+VBCRLF+VBCRLF+"<%=FTPErrorMSG%>"+VBCRLF+VBCRLF+"<%=hyftdGIPStr%>")
<%end if%>	         	     
       	    <%if request.querystring("S")="Approve" then%>
       	    	window.close
       	    	window.opener.GipApproveListRefresh
			<%elseif session("replyTitle") <> "" and session("replyID") <> 0 and request.querystring("delTask")="delREPLY" then%>
				document.location.href="<%=HTprogPrefix%>View.asp?icuitem=<%=session("replyID")%>&phase=edit"
       	    <%else%>
	    	document.location.href="<%=HTprogPrefix%>List.asp?keep=Y&nowpage=<%=mynowpage%>&pagesize=<%=mypagesize%>"
	    <%end if%>
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
 
<%
 
sub Activity1Action(actiontype, thisicuitem, thisieditor)

	'---題目上線加分,  若已有討論的內容, 也要將討論的人加分
	'---題目下線扣分, 

	Dim MemberExistFlag, QuestionCheckGrade
	MemberExistFlag = false
	QuestionCheckGrade = 0
	
	thissql = "SELECT * FROM ActivityMemberNew WHERE MemberId = '" & thisieditor & "'"
	Set thisrs = conn.execute(thissql)
	if not thisrs.eof then
		MemberExistFlag = true	
		QuestionCheckGrade = thisrs("QuestionCheckGrade")
	end if
	Set thisrs = Nothing
	
	If actiontype = "add" Then		
		'---member exist---add grade---
		if MemberExistFlag = true then 
			thissql = "UPDATE ActivityMemberNew SET QuestionCheckGrade = QuestionCheckGrade + 2 WHERE MemberId = '" & thisieditor & "'"
		else '---member not exist---add member---
			thissql = "INSERT INTO ActivityMemberNew(MemberId, QuestionCheckGrade) VALUES('" & thisieditor & "', 2)"
		end if
		conn.execute(thissql)	
		
		'---檢查有沒有討論的, 若已有討論的.將人加分...
		sql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeForum "
		sql = sql & "ON CuDTGeneric.icuitem = KnowledgeForum.gicuitem "
		sql = sql & "WHERE ParentIcuitem = " & thisicuitem 
		SET thisrs = conn.execute(sql)		
		While Not thisrs.eof
			
			sql = "SELECT * FROM ActivityMemberNew WHERE MemberId = '" & thisrs("iEditor") & "'"
			SET temprs = conn.execute(sql)
			If Not temprs.Eof Then
				thissql = "UPDATE ActivityMemberNew SET DiscussCheckGrade = DiscussCheckGrade + 2 WHERE MemberId = '" & thisrs("iEditor") & "'"
			Else
				thissql = "INSERT INTO ActivityMemberNew(MemberId, DiscussCheckGrade) VALUES('" & thisrs("iEditor") & "', 2)"
			End If
			SET temprs = Nothing
			conn.execute(thissql)
			
			thisrs.movenext
		Wend
		SET thisrs = Nothing
		
	ElseIf actiontype = "del" Then
		If QuestionCheckGrade >= 2 Then
			thissql = "UPDATE ActivityMemberNew SET QuestionCheckGrade = QuestionCheckGrade - 2 WHERE MemberId = '" & thisieditor & "'"
			conn.execute(thissql)
		End If			
	End If

end sub
 
 %>
