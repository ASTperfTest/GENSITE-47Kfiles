<%@  codepage="65001" %>
<% Response.Expires = 0
HTProgCap="電子報管理"
HTProgFunc="電子報發行清單"
HTProgCode="GW1M51"
HTProgPrefix="ePub"
%>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link type="text/css" rel="stylesheet" href="/css/list.css">
    <link type="text/css" rel="stylesheet" href="/css/layout.css">
    <link type="text/css" rel="stylesheet" href="/css/setstyle.css">
</head>
<!--#INCLUDE FILE="ePubListParam.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
    
	Dim Title
	sql = "SELECT title FROM EpPub WHERE (ePubID = " & pkStr(request("epubid"), "") & ")"
	Set Rs = Conn.Execute(sql)
	If Not Rs.Eof Then
		Title = Rs("title")
	End If 
	
	If Request("CtRootId") = "4" Then	
	    
    '宣告變數
     Dim URLs, xml,DataToSend,nCntChd,DataCount,CategoryCount
      
     '建立物件
     Set xml = Server.CreateObject ("Microsoft.XMLHTTP")
     
     '要檢查的網址
     tstr = now()
     If Request("sTitle") <> "" Then
		DataToSend = "keyword=" & Request("sTitle")
	 Else
	    DataToSend = "keyword="
     End If
     
	 DataToSend =  DataToSend & "&keywordfield=&folder=&category=2"
	 DataToSend =  DataToSend & "&docclass=&author=&datetime=&tag=&docclassvalue=公佈日期:[19000101160000 TO "
	 DataToSend =  DataToSend & year(tstr) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) 
	 DataToSend =  DataToSend & "]AND 可閱讀分眾導覽\(前端入口網\):A,B&containchildfolder=false"
	 DataToSend =  DataToSend & "&containchildcategory=true&enablekeywordsynonyms=false&enabletagsynonyms=false&sort=_l_last_modified_datetime"
     URLs = session("KmAPIURL") & "/search/advancedresult?who=" & session("KMAPIActor") & "&format=xml"
     
     IF CInt(Request.QueryString("PageNumber")) > 0 Then
     URLs = URLs & "&pi=" & CInt(Request.QueryString("PageNumber")) -1
     Else
     URLs = URLs & "&pi=0"
     End IF
    
	 IF cint(Request.QueryString("PageSize")) > 0 Then
	 URLs = URLs & "&ps=" & cint(Request.QueryString("PageSize"))
	 Else
	 URLs = URLs & "&ps=10"
	 End If
	 URLs = URLs & "&api_key=" & session("KMAPIKey")
     xml.Open "POST", URLs, false
	 xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
     xml.Send DataToSend

     IF xml.status=404 Then
          Response.Write "找不到頁面"
     ELSEIF xml.status=200 Then
		    Set xmlDoc = server.CreateObject("MSXML2.DOMDocument")
		    xmlDoc.LoadXML(xml.responsexml.xml)
			IF xmlDoc.Xml <> "" Then
				Set objNodeList = xmlDoc.documentElement.selectNodes("//b:anyType")
				DataCount = objNodeList.Length
				Dim result(),categoryIDs
				Redim result(DataCount)
				IF DataCount > 0 Then
				Set oNodeList = objNodeList(0).childNodes(4).childNodes
				For i = 0 To objNodeList.Length -1
					Set dataList = Server.CreateObject("Scripting.Dictionary")
					categoryIDs = ""
					For x = 0 To objNodeList(i).childNodes(4).childNodes.Length -1
						categoryIDs =  categoryIDs & objNodeList(i).childNodes(4).childNodes(x).text
						IF x <> objNodeList(i).childNodes(4).childNodes.Length -1 Then
							categoryIDs = categoryIDs & "|"
						End IF
					Next
					
					CategoryID = GetDetail("CATEGORY_ID",categoryIDs)
					CategoryName = GetDetail("CATEGORY_NAME",CategoryID)
					OUName = GetDetail("USER_OU",objNodeList(i).childNodes(1).text)
					UserName = GetDetail("USER_NAME",objNodeList(i).childNodes(1).text)

					dataList.Add "REPORT_ID", objNodeList(i).childNodes(24).text
					dataList.Add "CATEGORY_NAME", CategoryName
					dataList.Add "CATEGORY_ID", CategoryID
					dataList.Add "SUBJECT", objNodeList(i).childNodes(23).text
					dataList.Add "PUBLISHER", OUName
					dataList.Add "ONLINE_DATE", objNodeList(i).childNodes(10).text
					dataList.Add "ACTOR_DETAIL_NAME", UserName
				
				Set result(i) = dataList
				Next
				End If
			End If
     ELSE
          Response.Write "連線發生錯誤，代碼 "&xml.status
     END IF
      
     '關閉物件
     Set xml = Nothing
	 
	ElseIf Request("CtRootId") = "3" Then		
		
		sql = ""
		sql = sql & " SELECT DISTINCT CatTreeNode.CtRootID, CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CatTreeRoot.CtRootName, "
		sql = sql & " CtUnit.CtUnitName, CuDTGeneric.dEditDate, Member.realname FROM CuDTGeneric INNER JOIN CatTreeNode ON "
		sql = sql & " CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID INNER JOIN CatTreeRoot ON CatTreeNode.CtRootID = CatTreeRoot.CtRootID "
		sql = sql & " INNER JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
		sql = sql & " WHERE (CuDTGeneric.siteId = N'3') AND (CatTreeRoot.inUse = 'Y') AND (CuDTGeneric.fCTUPublic = 'Y') "
		sql = sql & " AND (CtUnit.inUse = 'Y') AND (CatTreeNode.inUse = 'Y') AND (CatTreeRoot.inUse = 'Y') AND (CatTreeNode.CtUnitID = 932) "
		
		If Request("sTitle") <> "" Then
			sql = sql & "AND (CuDTGeneric.sTitle LIKE " & pkDate("%" & Request("sTitle") & "%", "") & ") "
		End If
		
		If Request("StartDate") <> "" And Request("EndDate") <> "" Then
			sql = sql & "AND (CuDTGeneric.dEditDate BETWEEN " & pkStr(Request("StartDate"), "") & " AND " & pkStr(Request("EndDate"), "") & ") "
		ElseIf Request("StartDate") <> "" Then
			sql = sql & "AND (CuDTGeneric.dEditDate >= " & pkStr(Request("StartDate"), "") & ") "
		ElseIf Request("EndDate") <> "" Then
			sql = sql & "AND (CuDTGeneric.dEditDate <= " & pkStr(Request("EndDate"), "") & ") "
		End If
		
		If Request("Status") = "Y" Then
			sql = sql & "AND (CuDTGeneric.fCTUPublic = 'Y') "
		Else
			sql = sql & "AND (CuDTGeneric.fCTUPublic <> 'Y') "
		End If
		sql = sql & " ORDER BY CuDtGeneric.dEditDate DESC "
				
	Else
		
		sql = ""	
		sql = sql & "SELECT DISTINCT CatTreeNode.CtRootID, CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CatTreeRoot.CtRootName, "
		sql = sql & "CtUnit.CtUnitName, InfoUser.UserName, CuDTGeneric.dEditDate, Dept.deptName FROM CuDTGeneric INNER JOIN "
		sql = sql & "CatTreeNode ON CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID INNER JOIN CatTreeRoot ON "
		sql = sql & "CatTreeNode.CtRootID = CatTreeRoot.CtRootID INNER JOIN CtUnit ON "
		sql = sql & "CatTreeNode.CtUnitID = CtUnit.CtUnitID INNER JOIN InfoUser ON "
		sql = sql & "CuDTGeneric.iEditor = InfoUser.UserID INNER JOIN Dept ON CuDTGeneric.iDept = Dept.deptID "
		sql = sql & "WHERE (CatTreeRoot.inUse = 'Y') AND  (CtUnit.inUse = 'Y') AND (CatTreeNode.inUse = 'Y') "
		sql = sql & "AND (CatTreeRoot.inUse = 'Y') AND (Dept.inUse = 'Y')"
		
		If Request("CtRootId") <> "" Then		
			If Request("CtRootId") = "1" Then			
				sql = sql & "AND (CuDtGeneric.siteId = " & pkStr(Request("CtRootId"), "") & ") AND (CatTreeNode.CtRootID = 34) "			
			ElseIf Request("CtRootId") = "2" Then			
				sql = sql & "AND (CuDtGeneric.siteId = " & pkStr(Request("CtRootId"), "") & ") AND (CatTreeRoot.vGroup = 'XX')  "								
			End If
		End If
	
		If Request("CtUnitName") <> "" Then
			sql = sql & "AND (CtUnit.CtUnitName LIKE " & pkDate("%" & Request("CtUnitName") & "%", "") & ") "
		End If	
		
		If Request("sTitle") <> "" Then
			sql = sql & "AND (CuDTGeneric.sTitle LIKE " & pkDate("%" & Request("sTitle") & "%", "") & ") "
		End If
		
		If Request("StartDate") <> "" And Request("EndDate") <> "" Then
			sql = sql & "AND (CuDTGeneric.dEditDate BETWEEN " & pkStr(Request("StartDate"), "") & " AND " & pkStr(Request("EndDate"), "") & ") "
		ElseIf Request("StartDate") <> "" Then
			sql = sql & "AND (CuDTGeneric.dEditDate >= " & pkStr(Request("StartDate"), "") & ") "
		ElseIf Request("EndDate") <> "" Then
			sql = sql & "AND (CuDTGeneric.dEditDate <= " & pkStr(Request("EndDate"), "") & ") "
		End If
		
		If Request("Status") = "Y" Then
			sql = sql & "AND (CuDTGeneric.fCTUPublic = 'Y') "
		Else
			sql = sql & "AND (CuDTGeneric.fCTUPublic <> 'Y') "
		End If
		
		sql = sql & " ORDER BY CuDtGeneric.dEditDate DESC "
				
		'If Request("DeptId") <> "" Then
		'	sql = sql & "AND (CuDTGeneric.iDept = " & pkStr(Request("DeptId"), "") & ")"
		'End If
	
	End If
	
	PageNumber = Request.QueryString("PageNumber")  '現在頁數

	curRecCount = 0
	
	If Request("CtRootId") = "4" Then		
	IF xmlDoc.Xml <> "" Then
		Set objParentNode = xmlDoc.documentElement.selectNodes("//a:anyType")
	    If objParentNode.Length = 4 Then
	        If Int(objParentNode(3).childNodes(0).text) > 0 and DataCount > 0 Then
		    totRec = Int(objParentNode(3).childNodes(0).text) '總筆數
    	   
			    If totRec > 0 Then 
    			   '每頁筆數
			      PageSize = cint(Request.QueryString("PageSize"))

			      If PageSize <= 0 Then  
				    PageSize = 10  
			      End If 
				  
				  TotalPageCount = Int(totRec / PageSize) + 1
    			  
			      If cint(PageNumber) < 1 Then 
				     PageNumber = 1
			      ElseIf cint(PageNumber) > TotalPageCount Then 
				     PageNumber = totRec 
			      End If            	

			      totPage =  TotalPageCount'總頁數      
    			  
			      curRecCount = totRec - (PageSize * (PageNumber-1) )    
				    If curRecCount > PageSize Then
					    curRecCount = PageSize
				    End If
			    End If    
	    End If
	End If
	End If
	Else
	Set Rs = Server.CreateObject("ADODB.RecordSet")
'----------HyWeb GIP DB CONNECTION PATCH----------
'		Rs.Open sql,Conn,3,1
	Set Rs = Conn.execute(sql)

'----------HyWeb GIP DB CONNECTION PATCH----------
		If Not Rs.Eof Then 
		
			totRec = Rs.Recordcount       '總筆數
	   
			If totRec > 0 Then 
			  
			  PageSize = cint(Request.QueryString("PageSize"))
			  
			  If PageSize <= 0 Then  
				PageSize = 10  
			  End If 
			  
			  Rs.PageSize = PageSize       '每頁筆數

			  If cint(PageNumber) < 1 Then 
				 PageNumber = 1
			  ElseIf cint(PageNumber) > Rs.PageCount Then 
				 PageNumber = Rs.PageCount 
			  End If            	

			  Rs.AbsolutePage = PageNumber
			  totPage = Rs.PageCount       '總頁數      
			  
			  curRecCount = totRec - (PageSize * (PageNumber-1) )    
				If curRecCount > PageSize Then
					curRecCount = PageSize
				End If
			  
			End If    
		End If   
	End If

	Function CheckExist( epubid, articleid, ctrootid, categoryid )
	
		sql = "SELECT * FROM EpPubArticle WHERE EpubId = " & epubid & " AND ArticleId = " & articleid & " AND CtRootId = " & ctrootid & " AND CategoryId = '" & categoryid & "'"
		Set RsC = Conn.Execute(sql)
		If Not RsC.Eof Then
			CheckExist = true
		Else
			CheckExist = false
		End If
		RsC = Empty		
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
<body>
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="50%" align="left" nowrap class="FormName">
                電子報管理&nbsp;<font size="2">【資料查詢結果--<%=Title%>】</td>
            <td width="50%" class="FormLink" align="right">
                <a href="epaper_setList.asp?epubid=<%=Request("epubid")%>&eptreeid=<%=Request("eptreeid")%>"
                    target="">電子報文章列表</a> <a href="epaper_query.asp?epubid=<%=Request("epubid")%>&eptreeid=<%=Request("eptreeid")%>"
                        target="">重新查詢</a> <a href="Javascript:window.history.back();" title="回前頁">回前頁</a>
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="2">
                <hr noshade size="1" color="#000080">
            </td>
        </tr>
        <tr>
            <td class="Formtext" colspan="2" height="15">
            </td>
        </tr>
        <tr>
            <td align="center" colspan="2" width="80%" height="230" valign="top">
                <form method="POST" name="iForm">
                    <input type="hidden" name="CtRootId" value="<%=Request("CtRootId")%>" />
                    <input type="hidden" name="CtUnitName" value="<%=Request("CtUnitName")%>" />
                    <input type="hidden" name="sTitle" value="<%=Request("sTitle")%>" />
                    <input type="hidden" name="StartDate" value="<%=Request("StartDate")%>" />
                    <input type="hidden" name="EndDate" value="<%=Request("EndDate")%>" />
                    <input type="hidden" name="Status" value="<%=Request("Status")%>" />
                    <input type="hidden" name="DeptId" value="<%=Request("DeptId")%>" />
                    <input type="hidden" name="flag" value="0" />
                    <center>
                        <table cellspacing="1" cellpadding="0" width="100%" align="center" border="0">
                            <tbody>
                                <tr>
                                    <td valign="top" width="95%" colspan="2" height="230">
                                        <div id="Page">
                                            <p align="center">
                                                <font size="2" color="rgb(63,142,186)">第 <font size="2" color="#FF0000">
                                                    <%=PageNumber%>
                                                    /<%=totPage%></font>頁|共<font color="#FF0000"><%=totRec%></font> <font size="2" color="rgb(63,142,186)">
                                                        筆| 跳至第
                                                        <select id="PageNumber" number="PageNumber" size="1" style="color: #FF0000" onchange="PageNumberOnChange()">
                                                            <% For i = 1 To totPage %>
                                                            <option value="<%=i%>" <% If i = cint(PageNumber) Then %> selected <% End If %>>
                                                                <%=i%>
                                                            </option>
                                                            <% Next %>
                                                        </select>
                                                        頁
                                                        <% If cint(PageNumber) <> 1 Then %>
                                                        |<a href="javascript:PreviousPage()">上一頁</a>
                                                        <% End If %>
                                                        <% If Request("CtRootId") = "4" Then %>
                                                        <% If cint(PageNumber) <> TotalPageCount And cint(PageNumber) < TotalPageCount Then %>
                                                        |<a href="javascript:NextPage()">下一頁</a>
                                                        <% End If %>
                                                        <% Else %>
                                                        <% If cint(PageNumber) <> Rs.PageCount And cint(PageNumber) < Rs.PageCount Then %>
                                                        |<a href="javascript:NextPage()">下一頁</a>
                                                        <% End If %>
                                                        <% End If %>
                                                    </font>| 每頁筆數:
                                                    <select id="PageSize" name="PageSize" size="1" style="color: #FF0000" onchange="PageSizeOnChange()">
                                                        <option value="10" <% If PageSize = 10 Then %> selected <% End If %>>10</option>
                                                        <option value="20" <% If PageSize = 20 Then %> selected <% End If %>>20</option>
                                                        <option value="30" <% If PageSize = 30 Then %> selected <% End If %>>30</option>
                                                        <option value="50" <% If PageSize = 50 Then %> selected <% End If %>>50</option>
                                                    </select>
                                                </font>
                                        </div>
                                        <center>
                                            <table width="100%" id="ListTable">
                                                <tbody>
                                                    <tr align="left">
                                                        <th align="middle" width="7%">
                                                            <div align="center">
                                                                <input name="selectall" type="button" class="cbutton" onclick="CKBoxSelectAll()"
                                                                    value="全選/全不選"></div>
                                                        </th>
                                                        <th>
                                                            預覽</th>
                                                        <th>
                                                            目錄樹</th>
                                                        <th>
                                                            單元名稱</th>
                                                        <th>
                                                            資料標題</th>
                                                        <th>
                                                            帳號</th>
                                                        <th>
                                                            單位</th>
                                                        <th>
                                                            編修日期</th>
                                                    </tr>
                                                    <% If totRec > 0 Then %>
                                                    <% For i = 1 To PageSize%>
                                                    <tr>
                                                        <td align="middle">
                                                            <div align="center">
                                                                <% If Request("CtRootId") = "4" Then %>
                                                                <% If CheckExist( Request("epubid"), result(i-1)("REPORT_ID"), Request("CtRootId"), result(i-1)("CATEGORY_ID") ) = true Then %>
                                                                <input type="checkbox" name="ckbox" value="<%=result(i-1)("REPORT_ID")%>-<%=Request("CtRootId")%>-<%=result(i-1)("CATEGORY_ID")%>"
                                                                    checked>
                                                                <% Else %>
                                                                <input type="checkbox" name="ckbox" value="<%=result(i-1)("REPORT_ID")%>-<%=Request("CtRootId")%>-<%=result(i-1)("CATEGORY_ID")%>">
                                                                <% End If %>
                                                                <% Else %>
                                                                <% If CheckExist( Request("epubid"), Rs("iCuItem"), Request("CtRootId"), Request("CtRootId") ) = true Then %>
                                                                <input type="checkbox" name="ckbox" value="<%=Rs("iCUItem")%>-<%=Request("CtRootId")%>-<%=Request("CtRootId")%>"
                                                                    checked>
                                                                <% Else %>
                                                                <input type="checkbox" name="ckbox" value="<%=Rs("iCUItem")%>-<%=Request("CtRootId")%>-<%=Request("CtRootId")%>">
                                                                <% End If%>
                                                                <% End If %>
                                                            </div>
                                                        </td>
                                                        <% If Request("CtRootId") = "4" Then %>
                                                        <td>
															<% If session("myWWWSiteURL") <> "" Then %>
															
																<a target=_blank href="<%=session("myWWWSiteURL")%>/category/categorycontent.aspx?ReportId=<%=result(i-1)("REPORT_ID")%>&CategoryId=<%=result(i-1)("CATEGORY_ID")%>&ActorType=002&kpi=0">View</a></td>
															<% Else %>
																<a href="" target="_wMof">View</a></td>
															<% End If %>
                                                        <td>
                                                            知識庫</td>
                                                        <td>
                                                            <%=result(i-1)("CATEGORY_NAME")%>
                                                        </td>
                                                        <td>
                                                            <%=result(i-1)("SUBJECT")%>
                                                        </td>
                                                        <td>
                                                            <%=result(i-1)("ACTOR_DETAIL_NAME")%>
                                                            &nbsp;</td>
                                                        <td>
                                                            <%=result(i-1)("PUBLISHER")%>
                                                            &nbsp;</td>
                                                        <td>
                                                            <%=Replace(Mid(result(i-1)("ONLINE_DATE"),1,10),"-","/") %>
                                                        </td>
                                                        <% ElseIf Request("CtRootId") = "3" Then %>
                                                        <td>
                                                            <a href="#" target="_wMof">View</a></td>
                                                        <td>
                                                            知識家</td>
                                                        <td>
                                                            <%=Rs("CtUnitName")%>
                                                        </td>
                                                        <td>
                                                            <%=Rs("sTitle")%>
                                                        </td>
                                                        <td>
                                                            <%=Rs("realname")%>
                                                            &nbsp;</td>
                                                        <td>
                                                            知識家</td>
                                                        <td>
                                                            <%=FormatDateTime(Rs("dEditDate"),2)%>
                                                        </td>
                                                        <% Else %>
                                                        <% If Request("CtRootId") = "1" Then %>
                                                        <td>
                                                            <a href="#" target="_wMof">View</a></td>
                                                        <td>
                                                            入口網</td>
                                                        <% ElseIf Request("CtRootId") = "2" Then %>
                                                        <td>
                                                            <a href="#" target="_wMof">View</a></td>
                                                        <td>
                                                            主題網</td>
                                                        <% End If %>
                                                        <td>
                                                            <%=Rs("CtUnitName")%>
                                                        </td>
                                                        <td>
                                                            <%=Rs("sTitle")%>
                                                        </td>
                                                        <td>
                                                            <%=Rs("UserName")%>
                                                        </td>
                                                        <td>
                                                            <%=Rs("deptName")%>
                                                        </td>
                                                        <td>
                                                            <%=FormatDateTime(Rs("dEditDate"),2)%>
                                                        </td>
                                                        <% End If %>
                                                    </tr>
                                                    <%
														If Request("CtRootId") <> "4" Then
														Rs.MoveNext
														If Rs.Eof Then Exit For 
														Else
														If i >= DataCount  Then Exit For
														End If
													Next 
                                                    %>
                                                    <% Else %>
                                                    <tr>
                                                        <td align="middle" colspan="8">
                                                            <div align="center">
                                                                此條件下無資料</div>
                                                        </td>
                                                    </tr>
                                                    <% End If %>
                                                </tbody>
                                            </table>
                                        </center>
                        </table>
                    </center>
                </form>
                <input name="button" type="button" class="cbutton" onclick="SaveItems()" value="新增存檔">
                <input type="button" class="cbutton" onclick="ResetForm()" value="重設"></td>
        </tr>
    </table>
</body>
</html>

<script language="javascript">
	
	function PageNumberOnChange() 
	{
		var doc = document.iForm;
		doc.action = "epaper_resultList.asp?epubid=<%=Request("epubid")%>&eptreeid=<%=Request("eptreeid")%>&PageNumber=" + doc.PageNumber.options[doc.PageNumber.selectedIndex].value + "&PageSize=<%=PageSize%>";
		doc.submit();
	}
	function PageSizeOnChange() 
	{
		var doc = document.iForm;
		doc.action = "epaper_resultList.asp?epubid=<%=Request("epubid")%>&eptreeid=<%=Request("eptreeid")%>&PageNumber=<%=PageNumber%>&PageSize=" + doc.PageSize.options[doc.PageSize.selectedIndex].value;
		doc.submit();
	}
	function PreviousPage()
	{
		var doc = document.iForm;
		doc.action = "epaper_resultList.asp?epubid=<%=Request("epubid")%>&eptreeid=<%=Request("eptreeid")%>&PageNumber=<%=PageNumber-1%>&PageSize=<%=PageSize%>";
		doc.submit();
	}
	function NextPage()
	{
		var doc = document.iForm;
		doc.action = "epaper_resultList.asp?epubid=<%=Request("epubid")%>&eptreeid=<%=Request("eptreeid")%>&PageNumber=<%=PageNumber+1%>&PageSize=<%=PageSize%>";
		doc.submit();
	}	
	function CKBoxSelectAll()
	{
		var doc = document.iForm;
		var flag = false;
		if( doc.flag.value == "0" ) {
			if( <%=curRecCount%> == 1 ) {
				doc.ckbox.checked = true;			
			}
			else {
				for( var i = 0; i < <%=curRecCount%>; i++ ) {			
					doc.ckbox[i].checked = true;			
				}
			}
			doc.flag.value = "1";
		}
		else {
			if( <%=curRecCount%> == 1 ) {
				doc.ckbox.checked = false;			
			}
			else {
				for( var i = 0; i < <%=curRecCount%>; i++ ) {			
					doc.ckbox[i].checked = false;			
				}
			}
			doc.flag.value = "0";
		}
	}
	function ResetForm()
	{
		var doc = document.iForm;
		doc.Items.value = "";
		doc.action = "epaper_resultList.asp?epubid=<%=Request("epubid")%>&eptreeid=<%=Request("eptreeid")%>&PageNumber=1&PageSize=10";
		doc.submit();
	}
	function SaveItems()
	{
		var doc = document.iForm;
		var flag = false;
		var checkitems = "";
		var uncheckitems = "";
		if( <%=curRecCount%> == 1 ) {
			if( doc.ckbox.checked ) {
				flag = true;
			}				
		}
		else {
			for( var i = 0; i < <%=curRecCount%>; i++ ) {			
				if( doc.ckbox[i].checked ) {
					flag = true;
					break;
				}			
			}
		}
		if( !flag ) {
			alert("請至少選擇一篇文章");
			return;
		}
		else {
			if( <%=curRecCount%> == 1 ) {
				if( doc.ckbox.checked ) {
					checkitems = doc.ckbox.value + ";";	
					uncheckitems = ";";
				}
				else {
					checkitems = ";";	
					uncheckitems = doc.ckbox.value + ";";
				}
			}										
			else {
				for( var i = 0; i < <%=curRecCount%>; i++ ) {
					if( doc.ckbox[i].checked ) {
						checkitems += doc.ckbox[i].value + ";";	
					}			
					else{
						uncheckitems += doc.ckbox[i].value + ";";	
					}
				}
			}
			AddDelArticle( checkitems, uncheckitems);
		}		
	}
	function AddDelArticle( checkitems, uncheckitems)
	{
		var oxmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
		
		var httpStr = "../epaper/AddArticle.asp?epubid=<%=Request("epubid")%>&check=" + checkitems + "&uncheck=" + uncheckitems;
		oxmlhttp.open( "GET", httpStr, false );
  	oxmlhttp.send();
  	var oRtn = oxmlhttp.responseText;  	 
  	if(oRtn == 1) {  				
  		alert("(新增/刪除)成功");
  		document.iForm.action = "epaper_resultList.asp?epubid=<%=Request("epubid")%>&eptreeid=<%=Request("eptreeid")%>&PageNumber=<%=PageNumber%>&PageSize=<%=PageSize%>";
  		document.iForm.submit();
  	}  	
	}
</script>

