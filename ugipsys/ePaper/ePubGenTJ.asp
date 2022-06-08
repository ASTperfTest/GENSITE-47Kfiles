
<%@ CodePage = 65001%>
<% Response.Expires = 0
HTProgCode="GW1M51"
HTProgPrefix="ePub"
Response.codePage = 65001
Response.charset = "utf-8"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%
 
function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function	
	
	
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

	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/public/ePaper/ePaper"&epTreeID &".xml")
	xv = oxml.load(LoadXML)
	'Response.Write /site/" & session("mySiteID") & "/public/ePaper/ePaper"&epTreeID &".xml & "<HR>"
	'Response.End
     

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
         
 


'	dXml.selectSingleNode("epubId").text = ePubID
'	dXml.selectSingleNode("title").text = RSmaster("title")
	dXml.selectSingleNode("ePaperTitle").text = RSmaster("title")
	dXml.selectSingleNode("ePubDate").text = RSmaster("pubDate")
	dXml.selectSingleNode("ePaperXMLCSSURL").text = dXml.selectSingleNode("ePaperXMLCSSURL").text & "/ePaper/ePaper" & epTreeID & ".css"
	If nullText(dXml.selectSingleNode("ePaperXMLImgPath"))<>"" Then _
		dXml.selectSingleNode("ePaperXMLImgPath").text = dXml.selectSingleNode("ePaperXMLImgPath").text & epTreeID & "/"
	'Response.Write dXml.selectSingleNode("ePaperXMLIMGURL").text & "<HR>"
	'Response.End
	ePaperURL = dXml.selectSingleNode("ePaperURL").text

	SQLCom = "select n.*, u.ibaseDsd from CatTreeNode AS n JOIN CtUnit AS u ON n.ctUnitId=u.ctUnitId" _
		& " Where n.inUse='Y' AND ctRootId = "& epTreeID _
		& " ORDER BY n.catShowOrder"
	Set RS = conn.execute(SqlCom)

cvbCRLF = vbCRLF
cTab = ""
        
	
'TJ TJCategoryRoot start


		sqlstrVol=" SELECT  MAX(CAST(dbo.Article.Vol AS float))  as Vol  "
		sqlstrVol=sqlstrVol & " FROM   dbo.Article INNER JOIN  dbo.CuDtGeneric ON dbo.Article.gicuitem=dbo.CuDtGeneric.icuitem "                   
		sqlstrVol=sqlstrVol & " WHERE  dbo.CuDtGeneric.ictunit = 12"		
                sqlstrVol=sqlstrVol & " ORDER BY   dbo.Article.Vol DESC"
		
	        set RSTVol = conn.execute(sqlstrVol)
	        vol=""
	 
	 if Not RSTVol.EOF then
	    vol=RSTVol("Vol")
	 end if


            slStr = "<TJCategoryRoot><Vol>" & vol &"</Vol>" 
            
            
            
            
            
             
    xTRCategoryTemp=""
    sqlTEMP = "SELECT  mvalue  FROM   dbo.CodeMain WHERE   (codeMetaId = N'TJCategory') ORDER BY  mvalue"
	set RSTEMP = conn.execute(sqlTEMP)

  
  
   Do While Not RSTEMP.EOF 
	xTRCategory = "<TJCategory><name>" & RSTEMP("mvalue") & "</name>" 
	
   	
   		sqlstr=" SELECT  dbo.CuDtGeneric.icuitem, dbo.CuDtGeneric.stitle,dbo.Article.Writer,dbo.CuDtGeneric.xabstract  "
		sqlstr=sqlstr & " FROM   dbo.Article INNER JOIN  dbo.CuDtGeneric ON dbo.Article.gicuitem = dbo.CuDtGeneric.icuitem "
		sqlstr=sqlstr & " where dbo.CuDtGeneric.topCat = '" & RSTEMP("mvalue") & "' and dbo.Article.Vol='" & vol & "'"
     
         slStr=  slStr & xTRCategory
        
        
        
	  	set RSstr = conn.execute(sqlstr)
    
	    xTRCategoryTemp=""
	    
				 Do While Not RSstr.EOF 
				 xTRCategoryTemp=""
				   xTRCategoryTemp="<TJ_Article>"
					xTRCategoryTemp=xTRCategoryTemp & "<TJ_ArticleTitle>" & RSstr("stitle") & "</TJ_ArticleTitle>"
					xTRCategoryTemp=xTRCategoryTemp & "<TJ_ArticleBody>" & RSstr("xabstract") & "</TJ_ArticleBody>"
					xTRCategoryTemp=xTRCategoryTemp & "<TJ_ArticleWriter>" & RSstr("Writer") & "</TJ_ArticleWriter>"
          			xurl="<TJ_ArticleUrl>ct.asp?CtNode=122&xItem=" & RSstr("icuitem") & "</TJ_ArticleUrl>"
          		
					 xurl= deAmp(xurl) 	
			          
          			 xTRCategoryTemp=xTRCategoryTemp & xurl
				   
				     xTRCategoryTemp=xTRCategoryTemp & "</TJ_Article>"
				     
				     slStr= slStr & xTRCategoryTemp
				     RSstr.MoveNext
			    Loop
					
					
					
					   slStr=  slStr & "</TJCategory>"           
              
	   RSTEMP.MoveNext
	 
  Loop
            
          
             xTRCategory = "<TJCategory><name>Other</name>" 
   	
   		sqlstr=" SELECT  dbo.CuDtGeneric.icuitem, dbo.CuDtGeneric.stitle,dbo.Article.Writer,dbo.CuDtGeneric.xabstract  "
		sqlstr=sqlstr & " FROM   dbo.Article INNER JOIN  dbo.CuDtGeneric ON dbo.Article.gicuitem = dbo.CuDtGeneric.icuitem "
		sqlstr=sqlstr & " where dbo.CuDtGeneric.topCat is null and dbo.Article.Vol='" & vol & "'"
   
	    slStr=  slStr & xTRCategory
	  	set RSstr = conn.execute(sqlstr)
	  	
	    
	    xTRCategoryTemp=""
	    
				 Do While Not RSstr.EOF 
				 xTRCategoryTemp=""
				   xTRCategoryTemp="<TJ_Article>"
					xTRCategoryTemp=xTRCategoryTemp & "<TJ_ArticleTitle>" & RSstr("stitle") & "</TJ_ArticleTitle>"
					xTRCategoryTemp=xTRCategoryTemp & "<TJ_ArticleBody>" & RSstr("xabstract") & "</TJ_ArticleBody>"
					xTRCategoryTemp=xTRCategoryTemp & "<TJ_ArticleWriter>" & RSstr("Writer") & "</TJ_ArticleWriter>"
          			xurl="<TJ_ArticleUrl>ct.asp?CtNode=122&xItem=" & RSstr("icuitem") & "</TJ_ArticleUrl>"
          		
					xurl= deAmp(xurl) 	
			          
          				xTRCategoryTemp=xTRCategoryTemp & xurl
				   
				     xTRCategoryTemp=xTRCategoryTemp & "</TJ_Article>"
				        slStr=  slStr & xTRCategoryTemp
				     RSstr.MoveNext
			    Loop
					

               slStr=  slStr & "</TJCategory>" 

        
          slStr=slStr&"</TJCategoryRoot>"


	while Not RS.EOF
		slStr = slStr & "<epSection><secID>" & RS("CtNodeID") & "</secID><secName>"&RS("CatName")&"</secName><secURL>"&dXml.selectSingleNode("ePaperURL").text&"</secURL>"
		os = ""
			xSql = "SELECT Top " & xMaxNo
			xSql =xSql & " (SELECT count(*) FROM CuDtAttach AS dhtx WHERE bList='Y' AND dhtx.xiCUItem=ghtx.iCuItem) AS attachCount ,"
			xSql =xSql & " ictunit, ibaseDsd, iCuItem, sTitle, xBody, xImgFile, xURL, xPostDate, xabstract"
			'判斷是否有RSS欄位, 如果RSS='Y'則為外部匯入資料直接另開視窗到xURL..Apple 10/20
			IfSql="sp_columns @table_name = 'CuDtGeneric' , @column_name ='RSS'"
			Set IfRS = conn.execute(IfSql)
			If Not IfRS.EOF Then
				xSql = xSql & ",RSS "
			End If
			'-----------------------------------------------------------------------------
			xSql = xSql & " FROM CuDtGeneric AS ghtx" _
				& " WHERE iCtUnit = " & RS("CtUnitID") _
				& " AND (xPostDate BETWEEN " & pkStr(RSmaster("dbDate"),"") & " AND " & pkStr(RSmaster("deDate"),")") _
				& " AND fCTUPublic = 'Y' "
			'Response.Write xSql & "<HR>"
			'Response.End
			Set RSx = conn.execute(xSql)

			If Not RSx.EOF Then
				while Not RSx.EOF
					os = os & "<xItemList>"
					If Not isNull(RSx("xImgFile")) Then
						os = os & "<xImgFile>/public/data/"&RSx("xImgFile")&"</xImgFile>"
					End If
					os = os & "<xItemURL>"&dXml.selectSingleNode("ePaperURL").text&"</xItemURL>"
					os = os & "<xItem>"&RSx("iCuItem")&"</xItem>"
					os = os & "<sTitle><![CDATA["&RSx("sTitle")&"]]></sTitle>"
					os = os & "<xAbstract><![CDATA["&RSx("xabstract")&"]]></xAbstract>"					
                                        os = os & "<xURL><![CDATA["&RSx("xURL")&"]]></xURL>"
					os = os & "<xPostDate>"&RSx("xPostDate")&"</xPostDate>"
					os = os & "<attachCount>"&RSx("attachCount")&"</attachCount>"

					'判斷是否有RSS欄位, 如果RSS='Y'則為外部匯入資料直接另開視窗到xURL..Apple 2004/10/20
					If Not IfRS.EOF Then
						If Not isNull(RSx("RSS")) Then
							os = os & "<newWindow>" & RSx("RSS") & "</newWindow>"
						End If
					End If
					'-----------------------------------------------------------------------------
					xxBody = RSx("xBody")
					os = os & "<xBody><![CDATA["&xxBody&"]]></xBody>"

    					'判斷是否有附件, 如果attachCount > 0 則為有附件 列出附件列表..Apple 2006/05/08
                        If RSx("attachCount") > 0 Then
                          	attSql = "SELECT dhtx.*"
                          	attSql = attSql & " FROM CuDtAttach AS dhtx"
                          	attSql = attSql & " WHERE bList='Y'"
                          	attSql = attSql & " AND dhtx.xiCUItem=" & pkStr(RSx("iCUItem"),"")
                          	attSql = attSql & " ORDER BY dhtx.listSeq"
                            'Response.Write "attSQL: "& attSql & "<HR>"
                			'Response.End
                          	Set RSAttList = conn.execute(attSql)
                          	'Response.Write "<HR>"
                          	'Response.Write "附件" & RSAttList("nfileName")
                          	If  Not RSAttList.EOF Then
                              	os = os & "<AttachmentList>" & vbCRLF
                              	RSAttList.moveFirst

                              	'Response.Write "附件" & RSAttList("nfileName")

                              	While Not RSAttList.EOF
                              		os = os & "<Attachment>"
                              		os = os & "<URL><![CDATA[public/Attachment/" & RSAttList("nfileName")&"]]></URL>"
                              		os = os & "<Caption><![CDATA[" & RSAttList("atitle") & "]]></Caption>"
                              		os = os & "</Attachment>"
                              		RSAttList.moveNext
                              	Wend
                              	os = os & "</AttachmentList>"
                            Else
                                'Response.Write "No附件"
                            End If
                            'Response.Write "<HR>"
                        End If

					os = os & "</xItemList>"
					RSx.moveNext
				Wend
			End If

		slStr = slStr & os & "</epSection>" & cvbCRLF
           
'		Response.Write slStr & "<HR>"
		RS.moveNext
	Wend
       

	slStr = "<ePaperXML><epSectionList>" & cvbCRLF & slStr & "</epSectionList> </ePaperXML>"
'		Response.Write slStr & "<HR>"
	Set sxml = server.createObject("microsoft.XMLDOM")
	sxml.async = false
	sxml.setProperty "ServerHTTPRequest", true
	xv = sxml.loadXML(slStr)
  If sxml.parseError.reason <> "" Then
    Response.Write("XML parseError on line " &  sxml.parseError.line)
    Response.Write("<BR>Reasonyy: " &  sxml.parseError.reason)
    Response.End()
  End If
'	Response.Write sxml.xml & "<HR>"
'	Response.End

'	dXml.selectSingleNode("epSectionList").text = slStr
	Set xFieldNode = sxml.selectSingleNode("ePaperXML/epSectionList").cloneNode(true)
	oXml.selectSingleNode("ePaper").appendChild xFieldNode
	'Response.Write "/site/" & session("mySiteID") & "/public/ePaper/" & formID & ".xml"
	'Response.End

	oxml.save(Server.MapPath("/site/" & session("mySiteID") & "/public/ePaper/" & formID & ".xml"))
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
	oxsl.load(server.MapPath("/site/" & session("mySiteID") & "/public/ePaper/ePaper"&epTreeID &".xsl"))
'Response.Write "Hello2"
'Response.End

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

%>
