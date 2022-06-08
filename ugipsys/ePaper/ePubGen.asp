<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCode="GW1M51"
HTProgPrefix="ePub"
Response.codePage = 65001
Response.charset = "utf-8"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->

<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%
Function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
End Function

Function message(tempstr)
  xs = tempstr
  If xs="" OR isNull(xs) Then
  	message=""
  	exit Function
  elseIf instr(1,xs,"<P",1)>0 or instr(1,xs,"<BR",1)>0 or instr(1,xs,"<td",1)>0 Then
 	message=xs
  	exit Function
  End If
  	xs = replace(xs,vbCRLF&vbCRLF,"<P>")
  	xs = replace(xs,vbCRLF,"<BR/>")
  	message = replace(xs,chr(10),"<BR/>")
End Function

'----ftp參數處理
	FTPErrorMSG=""
	FTPfilePath="public/ePaper"
	SQLP = "Select * from UpLoadSite where upLoadSiteId='file'"
	Set RSP = conn.execute(SQLP)
	If Not RSP.EOF  Then
   		xFTPIP = RSP("upLoadSiteFtpip")
   		xFTPPort = RSP("upLoadSiteFtpport")
   		xFTPID = RSP("upLoadSiteFtpid")
   		xFTPPWD = RSP("upLoadSiteFtppwd")
	End If
'----ftp參數處理end


	epTreeID = session("epTreeID")		'-- 電子報的 tree

	Set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty "ServerHTTPRequest", true

	LoadXML = server.MapPath( "/site/" & session("mySiteID") & "/public/ePaper/ePaper" & epTreeID & ".xml" )
	xv = oxml.load(LoadXML)
  If oxml.parseError.reason <> "" Then
    Response.Write("XML parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  End If

  Set dXml = oxml.selectSingleNode("ePaper")

	ePubID = request.queryString("ePubID")
	formID = "ep" & ePubID
	sqlCom = "SELECT * FROM EpPub WHERE epubId=" & pkStr(ePubID,"")
	Set RSmaster = Conn.execute(sqlcom)
	xMaxNo = RSmaster("maxNo")
    'set xsl type
	ePubType = RSmaster("pubType")
	'dXml.selectSingleNode("epubId").text = ePubID
	'dXml.selectSingleNode("title").text = RSmaster("title")
	dXml.selectSingleNode("ePaperTitle").text = RSmaster("title")
	dXml.selectSingleNode("ePubDate").text = RSmaster("pubDate")
	dXml.selectSingleNode("ePaperXMLCSSURL").text = dXml.selectSingleNode("ePaperXMLCSSURL").text & "/ePaper/ePaper" & ePubType & ".css"
	If nullText(dXml.selectSingleNode("ePaperXMLImgPath")) <> "" Then
		dXml.selectSingleNode("ePaperXMLImgPath").text = dXml.selectSingleNode("ePaperXMLImgPath").text & ePubType & "/"
	End If
	
	SQLCom = "select n.*, u.ibaseDsd from CatTreeNode AS n JOIN CtUnit AS u ON n.ctUnitId=u.ctUnitId "
	If epTreeID = 21 Then
		SQLCom = SQLCom & " Where n.inUse='Y' AND ctRootId = 177 ORDER BY n.catShowOrder "
	Else
		SQLCom = SQLCom & " Where n.inUse='Y' AND ctRootId = " & epTreeID & " ORDER BY n.catShowOrder "
	End If
	Set RS = conn.execute(SqlCom)

	cvbCRLF = vbCRLF
	cTab = ""
	slStr = ""
	
	while Not RS.EOF
	
		slStr = slStr & "<epSection>"
		slStr = slStr & "<secID>" & RS("CtNodeID") & "</secID>"
		slStr = slStr & "<secName>" & RS("CatName")& "</secName>"
		slStr = slStr & "<secURL>" & dXml.selectSingleNode("ePaperURL").text & "</secURL>"
		
		os = ""
		xSql = "SELECT Top " & xMaxNo
		xSql = xSql & " (SELECT count(*) FROM CuDtAttach AS dhtx WHERE bList='Y' AND dhtx.xiCUItem=ghtx.iCuItem) AS attachCount ,"
		xSql = xSql & " ictunit, ibaseDsd, iCuItem, sTitle, xBody, xImgFile, xURL, xPostDate, xabstract, abstract, fileDownLoad, topCat, showType "
		
		'-----------------------------------------------------------------------------
		'判斷是否有RSS欄位, 如果RSS='Y'則為外部匯入資料直接另開視窗到xURL..Apple 10/20
		IfSql = "sp_columns @table_name = 'CuDtGeneric' , @column_name ='RSS'"
		Set IfRS = conn.execute(IfSql)
		If Not IfRS.EOF Then
			xSql = xSql & ", RSS "
		End If
		'-----------------------------------------------------------------------------
		
		xSql = xSql & " FROM CuDtGeneric AS ghtx "
		
		'If RS("CtUnitID") = 807 then xSql = xSql & " INNER JOIN KnowledgeForum AS htx ON ghtx.icuitem = htx.gicuitem "
		'If RS("CtUnitID") = 932 then xSql = xSql & " INNER JOIN KnowledgeForum AS htx ON ghtx.icuitem = htx.gicuitem "
		
		xSql = xSql & " WHERE iCtUnit = " & RS("CtUnitID") & " AND (xPostDate BETWEEN '" & RSmaster("dbDate") & " 00:00:01' AND '" & RSmaster("deDate") & " 23:59:59') " 
		
		'If RS("CtUnitID") = 807 then xSql = xSql & " AND htx.Status <> 'D' "
		'If RS("CtUnitID") = 932 then xSql = xSql & " AND htx.Status <> 'D' "
		
		xSql = xSql & " AND fCTUPublic = 'Y' "
		
		'If RS("CtUnitID") = 821 then 
		If RS("CtUnitID") = 1353 then 
			xSql = xSql & " ORDER BY xImportant DESC, xPostDate DESC"
		Else
			xSql = xSql & " ORDER BY xPostDate DESC"
		End If

		Set RSx = conn.execute(xSql)

		Dim newImgFile : newImgFile = ""
		
		If Not RSx.EOF Then
			
			while Not RSx.EOF
			
				os = os & "<xItemList>"
				
				If Not isNull(RSx("xImgFile")) Then
					
				End If
				newImgFile = RSx("xImgFile")
				If RSx("iCtUnit") = "1351" Then					
					If Not isNull(RSx("xImgFile")) AND RSx("xImgFile") <> "" Then
						newImgFile = RSx("xImgFile")						
					End If									
				End If
				
				os = os & "<xItemURL>" 						& dXml.selectSingleNode("ePaperURL").text & "</xItemURL>"
				os = os & "<xItem>" 							& RSx("iCuItem") 													& "</xItem>"
				os = os & "<iBaseDSD>" 						& RSx("ibaseDsd") 												& "</iBaseDSD>"
				os = os & "<iCtUnit>" 						& RSx("ictunit") 													& "</iCtUnit>"				
				os = os & "<CtNode>"							& RS("CtNodeID")													& "</CtNode>"
				os = os & "<sTitle><![CDATA[" 		& RSx("sTitle") 													& "]]></sTitle>"
				os = os & "<xAbstract><![CDATA[" 	& RSx("xabstract") 												& "]]></xAbstract>"
				os = os & "<Abstract><![CDATA[" 	& RSx("abstract") 												& "]]></Abstract>"				
				os = os & "<xBody><![CDATA[" 			& RSx("xBody") 														& "]]></xBody>"
				os = os & "<showType>"			 			& RSx("showType")													& "</showType>"
				os = os & "<xURL><![CDATA[" 			& RSx("xURL") 														& "]]></xURL>"
				os = os & "<xPostDate>" 					& RSx("xPostDate") 												& "</xPostDate>"
				os = os & "<fileDownLoad>" 				& RSx("fileDownLoad") 										& "</fileDownLoad>"
				os = os & "<topCat>" 							& RSx("topCat") 													& "</topCat>"
				os = os & "<xImgFile>" 						& newImgFile															& "</xImgFile>"
				os = os & "<attachCount>" 				& RSx("attachCount") 											& "</attachCount>"				

				'-----------------------------------------------------------------------------
				'判斷是否有RSS欄位, 如果RSS='Y'則為外部匯入資料直接另開視窗到xURL..Apple 2004/10/20
				If Not IfRS.EOF Then
					If Not isNull(RSx("RSS")) Then
						os = os & "<newWindow>" & RSx("RSS") & "</newWindow>"
					End If
				End If
			
				'判斷是否有附件, 如果attachCount > 0 則為有附件 列出附件列表..Apple 2006/05/08
        If RSx("attachCount") > 0 Then
	       	attSql = "SELECT dhtx.* FROM CuDtAttach AS dhtx WHERE bList='Y' AND dhtx.xiCUItem = " & pkStr(RSx("iCUItem"),"") & " ORDER BY dhtx.listSeq"
                            
         	Set RSAttList = conn.execute(attSql)

         	If  Not RSAttList.EOF Then
	         	os = os & "<AttachmentList>" & vbCRLF
           	RSAttList.moveFirst
  
           	While Not RSAttList.EOF
  	       		os = os & "<Attachment>"
           		os = os & "<URL><![CDATA[public/Attachment/" & RSAttList("nfileName")&"]]></URL>"
           		os = os & "<Caption><![CDATA[" & RSAttList("atitle") & "]]></Caption>"
           		os = os & "</Attachment>"
           		RSAttList.moveNext
           	Wend
           	os = os & "</AttachmentList>"
          Else
          End If
        End If
				os = os & "</xItemList>"
				RSx.moveNext
			Wend
		End If

		slStr = slStr & os & "</epSection>" & cvbCRLF		
		RS.moveNext
		
	Wend
	
	os = ""
	'---for 2008 epaper ap---
	'If epTreeID = "143" Then ' epTreeID = "177" Or epTreeID = "21"
	
	If epTreeID = "21" Then 
		
		Set KMConn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'		KMConn.Open session("KMODBC")
'Set KMConn = Server.CreateObject("HyWebDB3.dbExecute")
KMConn.ConnectionString = session("KMODBC")
KMConn.ConnectionTimeout=0
KMConn.CursorLocation = 3
KMConn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

		
		nstr = "SELECT * FROM EpPubArticle WHERE EPubId = " & ePubID
		Set nrs = Conn.Execute(nstr)
		
		os = os & "<epSection>"
		os = os & "<secID>9999</secID>"
		os = os & "<secName></secName>"
		os = os & "<secURL>" & dXml.selectSingleNode("ePaperURL").text & "</secURL>"
		While Not nrs.Eof 					
			os = os & GetXmlContent( nrs("ArticleId"), nrs("CtRootId"), nrs("categoryid"), nrs("ISFORMER") )			
			nrs.MoveNext		
		Wend
		os = os & "</epSection>"
	End If
	slStr = slStr & os & cvbCRLF
	
	slStr = "<ePaperXML><epSectionList>" & cvbCRLF & slStr & "</epSectionList></ePaperXML>"
	
	Set sxml = server.createObject("microsoft.XMLDOM")
	sxml.async = false
	sxml.setProperty "ServerHTTPRequest", true
	xv = sxml.loadXML(slStr)
	
  If sxml.parseError.reason <> "" Then
    Response.Write("XML parseError on line " &  sxml.parseError.line)
    Response.Write("<BR>Reasonyy: " &  sxml.parseError.reason)
    Response.End()
  End If
	'dXml.selectSingleNode("epSectionList").text = slStr
	Set xFieldNode = sxml.selectSingleNode("ePaperXML/epSectionList").cloneNode(true)
	oXml.selectSingleNode("ePaper").appendChild xFieldNode
	
	oxml.save(Server.MapPath("/site/" & session("mySiteID") & "/public/ePaper/" & formID & ".xml"))
	If oxml.parseError.reason <> "" Then
  	Response.Write("XML parseError on line " &  oxml.parseError.line)
  	Response.Write("<BR>Reasonaa: " &  oxml.parseError.reason)
  	Response.End()
  End If
	'----FTP機制
	If xFTPIP<>"" and xFTPID<>"" and xFTPPWD<>"" Then
		fileAction="MoveFile"
		fileTarget=formID & ".xml"
		fileSource=Server.MapPath("/site/" & session("mySiteID") & "/public/ePaper/" & formID & ".xml")
		ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,"",fileTarget,fileSource
  End If
	'-----------------------------顯示----------
	'----Load epaper.xml
	Set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/site/" & session("mySiteID") & "/public/ePaper/" & formID & ".xml")
	xv = oxml.load(LoadXML)
	
  If oxml.parseError.reason <> "" Then
  	Response.Write("XML parseError on line " &  oxml.parseError.line)
  	Response.Write("<BR>Reasonaa: " &  oxml.parseError.reason)
  	Response.End()
  End If
  '----Load epaper.xsl
	Set oxsl = server.createObject("microsoft.XMLDOM")
	oxsl.load(server.MapPath("/site/" & session("mySiteID") & "/public/ePaper/ePaper"& ePubType &".xsl"))
	Response.ContentType = "text/HTML"
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	outString = replace(outString,"&amp;","&")
	Response.Write replace(outString,"&amp;","&")
	
	'-----------------------------gen html file start ----------
	If checkGIPconfig("epaperGenHtmlFile") Then
    	dim saveHTMLFile
    	const adSaveCreateOverWrite = 2
    	saveHTMLFile = server.mappath("/site/" & session("mySiteID") & "/public/ePaper/ePaper"&epTreeID &"_"& formID &".htm")

      dim objStream
      Set objStream = server.createobject("ADODB.Stream")
      objStream.Open()
      objStream.CharSet = "UTF-8"
      objStream.WriteText(outString)
      objStream.SaveToFile saveHTMLFile, adSaveCreateOverWrite
      objStream.Close()
	End If

	Response.End

	Function GetXmlContent( aid, rid, gid, isformer )

		str = ""
		str = str & "<xItemList>"
		If rid = "1" Then '入口網
			
			sql = ""	
			sql = sql & "SELECT DISTINCT CatTreeNode.CtRootID, CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.showType, CuDTGeneric.xURL, "
			sql = sql & "CuDTGeneric.fileDownLoad, CatTreeRoot.CtRootName, CtUnit.CtUnitId, InfoUser.UserName, CuDTGeneric.xPostDate, Dept.deptName, "
			sql = sql & "CuDtGeneric.ibaseDsd, CatTreeNode.CtNodeID "
			sql = sql & "FROM CuDTGeneric INNER JOIN CatTreeNode ON CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID INNER JOIN CatTreeRoot ON "
			sql = sql & "CatTreeNode.CtRootID = CatTreeRoot.CtRootID INNER JOIN CtUnit ON "
			sql = sql & "CatTreeNode.CtUnitID = CtUnit.CtUnitID INNER JOIN InfoUser ON "
			sql = sql & "CuDTGeneric.iEditor = InfoUser.UserID INNER JOIN Dept ON CuDTGeneric.iDept = Dept.deptID "
			sql = sql & "WHERE (CatTreeRoot.inUse = 'Y') AND  (CtUnit.inUse = 'Y') AND (CatTreeNode.inUse = 'Y') "
			sql = sql & "AND (CatTreeRoot.inUse = 'Y') AND (Dept.inUse = 'Y') "
			sql = sql & "AND (CuDTGeneric.iCuItem = " & aid & ")"
			
			set newrs = Conn.Execute(sql)			
			if not newrs.eof then				
				str = str & "<xItemURL>" 						& dXml.selectSingleNode("ePaperURL").text 	& "</xItemURL>"
				str = str & "<xItem>" 							& newrs("iCuItem") 													& "</xItem>"
				str = str & "<iBaseDSD>" 						& newrs("ibaseDsd") 												& "</iBaseDSD>"
				str = str & "<iCtUnit>" 						& newrs("CtUnitId") 												& "</iCtUnit>"		
				str = str & "<CtNode>" 							& newrs("CtNodeID") 												& "</CtNode>"			
				str = str & "<sTitle><![CDATA[" 		& newrs("sTitle") 													& "]]></sTitle>"
				str = str & "<showType>"			 			& newrs("showType")													& "</showType>"
				str = str & "<xURL><![CDATA[" 			& newrs("xURL") 														& "]]></xURL>"			
				str = str & "<fileDownLoad>" 				& newrs("fileDownLoad") 										& "</fileDownLoad>"
				str = str & "<xPostDate>" 					& FormatDateTime(newrs("xPostDate"),2) 												& "</xPostDate>"      
      end if
			newrs.close
			set newrs = nothing
      
    ElseIf rid = "2" Then  '主題館
      
      sql = ""	
			sql = sql & "SELECT DISTINCT CatTreeNode.CtRootID, CuDTGeneric.iCUItem, CuDTGeneric.sTitle,CuDTGeneric.showType,CuDTGeneric.xURL, CatTreeRoot.CtRootName, "
			sql = sql & "CtUnit.CtUnitId, InfoUser.UserName, CuDTGeneric.xPostDate, Dept.deptName, CuDtGeneric.ibaseDsd, CatTreeNode.CtNodeID, CuDtGeneric.fileDownLoad "
			sql = sql & "FROM CuDTGeneric INNER JOIN CatTreeNode ON CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID INNER JOIN CatTreeRoot ON "
			sql = sql & "CatTreeNode.CtRootID = CatTreeRoot.CtRootID INNER JOIN CtUnit ON "
			sql = sql & "CatTreeNode.CtUnitID = CtUnit.CtUnitID INNER JOIN InfoUser ON "
			sql = sql & "CuDTGeneric.iEditor = InfoUser.UserID INNER JOIN Dept ON CuDTGeneric.iDept = Dept.deptID "
			sql = sql & "WHERE (CatTreeRoot.inUse = 'Y') AND  (CtUnit.inUse = 'Y') AND (CatTreeNode.inUse = 'Y') "
			sql = sql & "AND (CatTreeRoot.inUse = 'Y') AND (Dept.inUse = 'Y') "
			sql = sql & "AND (CuDTGeneric.iCuItem = " & aid & ")"
			set newrs = Conn.Execute(sql)
			if not newrs.eof then
			    if instr(newrs("xURL"),"/Knowledge/")>0 then
				    newURL = dXml.selectSingleNode("ePaperURL").text&newrs("xURL") 
			    else
				    newURL = newrs("xURL") 
			    end if
			end if
			if not newrs.eof then			
				str = str & "<xItemURL>" 						& dXml.selectSingleNode("ePaperURL").text 	& "</xItemURL>"
				str = str & "<xItem>" 							& newrs("iCuItem") 													& "</xItem>"
				str = str & "<iBaseDSD>" 						& newrs("ibaseDsd") 												& "</iBaseDSD>"
				str = str & "<iCtUnit>" 						& newrs("CtUnitId") 												& "</iCtUnit>"		
				str = str & "<xdmp>"			 					& newrs("CtRootID") 												& "</xdmp>"
				str = str & "<CtNode>"	 						& newrs("CtNodeID") 												& "</CtNode>"				
				str = str & "<sTitle><![CDATA["			& newrs("sTitle") 													& "]]></sTitle>"
				str = str & "<showType>"			 			& newrs("showType")													& "</showType>"
				str = str & "<xURL><![CDATA[" 			& newURL															& "]]></xURL>"			
				str = str & "<fileDownLoad>" 				& newrs("fileDownLoad") 										& "</fileDownLoad>"
				str = str & "<xPostDate>" 					& FormatDateTime(newrs("xPostDate"),2) 												& "</xPostDate>"    
			end if
			newrs.close
			set newrs = nothing
			
    ElseIf rid = "3" Then  '知識家
      
      sql = ""
			sql = sql & " SELECT DISTINCT CatTreeNode.CtRootID, CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CatTreeRoot.CtRootName, CuDTGeneric.xPostDate, "
			sql = sql & " CtUnit.CtUnitId, CuDTGeneric.dEditDate, Member.realname, CuDTGeneric.ibaseDsd FROM CuDTGeneric INNER JOIN "
			sql = sql & " CatTreeNode ON CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID INNER JOIN CatTreeRoot ON CatTreeNode.CtRootID = CatTreeRoot.CtRootID "
			sql = sql & " INNER JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
			sql = sql & " WHERE (CuDTGeneric.siteId = N'3') AND (CatTreeRoot.inUse = 'Y') AND (CuDTGeneric.fCTUPublic = 'Y') "
			sql = sql & " AND (CtUnit.inUse = 'Y') AND (CatTreeNode.inUse = 'Y') AND (CatTreeRoot.inUse = 'Y') AND (CatTreeNode.CtUnitID = 932) "
			sql = sql & " AND (CuDTGeneric.iCuItem = " & aid & ")"
			
			set newrs = Conn.Execute(sql)			
			if not newrs.eof then
				str = str & "<xItemURL>" 						& dXml.selectSingleNode("ePaperURL").text 	& "</xItemURL>"
				str = str & "<xItem>" 							& newrs("iCuItem") 													& "</xItem>"
				str = str & "<iBaseDSD>" 						& newrs("ibaseDsd") 												& "</iBaseDSD>"
				str = str & "<iCtUnit>" 						& newrs("CtUnitId") 												& "</iCtUnit>"				
				str = str & "<sTitle><![CDATA[" 		& newrs("sTitle") 													& "]]></sTitle>"
				str = str & "<xURL></xURL>"
				str = str & "<xPostDate>" 					& FormatDateTime(newrs("xPostDate"),2) 												& "</xPostDate>"
      end if
			newrs.close
			set newrs = nothing
      
    ElseIf rid = "4" Then
			
			sql = ""
			IF CBool(ISFORMER) THEN
				sql = sql & " SELECT DISTINCT REPORT.REPORT_ID, CATEGORY.CATEGORY_ID, REPORT.SUBJECT, REPORT.PUBLISHER, REPORT.ONLINE_DATE, "
				sql = sql & " ACTOR_INFO.ACTOR_DETAIL_NAME FROM REPORT INNER JOIN ACTOR_INFO ON REPORT.CREATE_USER = ACTOR_INFO.ACTOR_INFO_ID "
				sql = sql & " INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID INNER JOIN CATEGORY ON "
				sql = sql & " CAT2RPT.DATA_BASE_ID = CATEGORY.DATA_BASE_ID AND CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID "
				sql = sql & " WHERE (REPORT.STATUS = 'PUB') AND (REPORT.ONLINE_DATE < GETDATE()) AND (CATEGORY.DATA_BASE_ID = 'DB020') "
				sql = sql & " AND (REPORT.REPORT_ID = '" & aid & "')"
				
				Set kmrs = KMConn.Execute(sql)
				if not kmrs.eof then 
					str = str & "<xItemURL>" 						& dXml.selectSingleNode("ePaperURL").text 	& "</xItemURL>"
					str = str & "<xItem>" 							& kmrs("REPORT_ID") 												& "</xItem>"
					str = str & "<iBaseDSD>"						& "DB020"																		& "</iBaseDSD>"
					str = str & "<iCtUnit>" 						& kmrs("CATEGORY_ID") 											& "</iCtUnit>"				
					str = str & "<sTitle><![CDATA[" 		& kmrs("SUBJECT") 													& "]]></sTitle>"
					str = str & "<xURL></xURL>"
					str = str & "<xPostDate>" 					& FormatDateTime(kmrs("ONLINE_DATE"),2) 											& "</xPostDate>"
				end if
				kmrs.close
				set kmrs = nothing
			ELSE
					info = GetDetail("DOCUMENT",aid)
					infos = split(info,"|")
					str = str & "<xItemURL>"& dXml.selectSingleNode("ePaperURL").text 	& "</xItemURL>"
					str = str & "<xItem>"& aid & "</xItem>"
					str = str & "<iBaseDSD>"& "DB020"& "</iBaseDSD>"
					str = str & "<iCtUnit>"& gid & "</iCtUnit>"				
					str = str & "<sTitle><![CDATA["& infos(0) & "]]></sTitle>"
					str = str & "<xURL></xURL>"
					str = str & "<xPostDate>"& Replace(Mid(infos(2),1,10),"-","/") & "</xPostDate>"
					str = str & "<isFormer>N</isFormer>"
			END IF
    End If  
    
    str = str & "<rid>" & rid	& "</rid>"
  	str = str & "</xItemList>"
		GetXmlContent = str			
	End Function
	
	Function GetDetail(typeid, key)
		'建立物件
		 dim xmlhttp
		 Set xmlhttp = Server.CreateObject("Microsoft.XMLHTTP")
		 '要檢查的網址
		 URLs = session("mySiteMMOURL") & "/epaper/epaper_querydetail.asp?typeid=" & typeid & "&KEY=" & key

		 IF typeid = "" OR key = "" Then
		 Response.Write "未傳入必要參數"
		 Response.Write.End()
		 End IF
		 'Response.Write xmlhttp
		 xmlhttp.Open "Get", URLs, false
		 xmlhttp.Send
		 
		 IF xmlhttp.status=404 Then
			  Response.Write "找不到頁面"
		 ELSEIF xmlhttp.status=200 Then
				GetDetail = xmlhttp.responsetext
		 Else
			  Response.Write xmlhttp.status
		 End If
	End Function
%>
