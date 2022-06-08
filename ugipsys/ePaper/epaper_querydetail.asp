<%
	Response.Expires = 0
	Response.Expiresabsolute = Now() - 1 
	Response.AddHeader "pragma","no-cache" 
	Response.AddHeader "cache-control","private" 
	Response.CacheControl = "no-cache"

	typeid = request("typeid")
	key = request("key")
	
	dim result
	IF typeid = "USER_OU" THEN
	result = GetUserOUNameByID(key)
	ELSEIF typeid = "CATEGORY_NAME" THEN
	result = GetCategoryNameByID(key)
	ELSEIF typeid = "USER_NAME" THEN
	result = GetUserNameByID(key)
	ELSEIF typeid = "DOCUMENT" THEN
	result = GetDocumentDetail(key)
	ELSEIF typeid = "CATEGORY_ID" THEN
	result = GetCategoryIDByIds(key)
	END IF
	
	Response.Write result
	
	Function GetCategoryNameByID(categoryid)
	
	 '建立物件
     Set xmlHttp = Server.CreateObject ("Microsoft.XMLHTTP")
     '要檢查的網址
     URLs = session("KmAPIURL") & "/category/"

     IF categoryid <> "" Then
     URLs = URLs & categoryid & "?load_path=false&format=xml&tid=0&who=" & session("KMAPIActor") & "&pi=0&ps=10&api_key=" & session("KMAPIKey")
     Else
     Response.Write "未傳入CategoryID"
     End IF

     xmlHttp.Open "Get", URLs, false
     xmlHttp.Send
	 IF xmlHttp.status=404 Then
          Response.Write "找不到頁面"
     ELSEIF xmlHttp.status=200 Then
		  Set domDocument = server.CreateObject("MSXML2.DOMDocument")
		  domDocument.LoadXML(xmlHttp.responsexml.xml)
		  Set objCategoryNodes = domDocument.documentElement.childNodes(2).childNodes
		  GetCategoryNameByID = objCategoryNodes(3).text
	 Else
		  Response.Write xmlHttp.status
	 End If
	End Function
	
	Function GetUserOUNameByID(subjectID)
		'建立物件
		 Set xmlOUHttp = Server.CreateObject ("Microsoft.XMLHTTP")
		 '要檢查的網址
		 URLs =  session("KmAPIURL") & "/subject/fetch/user/ous/"

		 IF subjectID <> "" Then
		 URLs = URLs & subjectID & "?load_path=false&format=xml&tid=0&who=" & session("KMAPIActor") & "&pi=0&ps=10&api_key=" & session("KMAPIKey")
		 Else
		 Response.Write "未傳入SubjectID"
		 End IF

		 xmlOUHttp.Open "Get", URLs, false
		 xmlOUHttp.Send
		 
		 IF xmlOUHttp.status=404 Then
			  Response.Write "找不到頁面"
		 ELSEIF xmlOUHttp.status=200 Then
				Set ouDocument = server.CreateObject("MSXML2.DOMDocument")
				ouDocument.LoadXML(xmlOUHttp.responsexml.xml)
				Set objOUNodes = ouDocument.documentElement.selectNodes("//a:Subject")
				IF objOUNodes.Length > 0 Then
					Set OuDetailNodes = objOUNodes(0).childNodes
					GetUserOUNameByID = OuDetailNodes(1).text
				Else
					GetUserOUNameByID = Empty
				End IF
		 Else
			  Response.Write xmlOUHttp.status
		 End If
	End Function
	
	Function GetUserNameByID(subjectID)
		'建立物件
		 Set xmlUserHttp = Server.CreateObject ("Microsoft.XMLHTTP")
		 '要檢查的網址
		 URLs = session("KmAPIURL") & "/user/exact/"

		 IF subjectID <> "" Then
		 URLs = URLs & subjectID & "?load_path=false&format=xml&tid=0&who=" & session("KMAPIActor") & "&pi=0&ps=10&api_key=" & session("KMAPIKey")
		 Else
		 Response.Write "未傳入SubjectID"
		 End IF

		 xmlUserHttp.Open "Get", URLs, false
		 xmlUserHttp.Send
		 
		 IF xmlUserHttp.status=404 Then
			  Response.Write "找不到頁面"
		 ELSEIF xmlUserHttp.status=200 Then
				Set userDocument = server.CreateObject("MSXML2.DOMDocument")
				userDocument.LoadXML(xmlUserHttp.responsexml.xml)
				Set objUserNodes = userDocument.documentElement.selectNodes("//a:DisplayName")
				IF objUserNodes.Length > 0 Then
					Set UserDetailNodes = objUserNodes(0).childNodes
					GetUserNameByID = objUserNodes(0).text
				Else
					GetUserNameByID = Empty
				End IF
		 Else
			  Response.Write xmlUserHttp.status
		 End If
	End Function
	
	Function GetDocumentDetail(documentid)
		'建立物件
		 Set xmldocumentHttp = Server.CreateObject ("Microsoft.XMLHTTP")
		 '要檢查的網址
		 URLs = session("KmAPIURL") & "/document/"

		 IF documentid <> "" Then
		 URLs = URLs & documentid & "?load_path=false&format=xml&tid=0&who=" & session("KMAPIActor") & "&pi=0&ps=10&api_key=" & session("KMAPIKey")
		 Else
		 Response.Write "未傳入Documentid"
		 End IF
		 xmldocumentHttp.Open "Get", URLs, false
		 xmldocumentHttp.Send
		 
		 IF xmldocumentHttp.status=404 Then
				Response.Write "找不到頁面"
		 ELSEIF xmldocumentHttp.status=200 Then
				Set document = server.CreateObject("MSXML2.DOMDocument")
				document.LoadXML(xmldocumentHttp.responsexml.xml)
				'組合TITLE|USERID|CREATIONDATE
				Set titleNodes = document.documentElement.selectNodes("//a:VersionTitle")
				Set titleChildNodes = titleNodes(0).childNodes
				Set userIdNodes = document.documentElement.selectNodes("//a:VersionCreatorId")
				Set userIdChildNodes = UserIdNodes(0).childNodes
				Set creationDateNodes = document.documentElement.selectNodes("//a:VersionCreationDatetime")
				Set creationDateChildNodes = creationDateNodes(0).childNodes
				GetDocumentDetail = titleChildNodes(0).text & "|" & userIdChildNodes(0).text & "|" & creationDateChildNodes(0).text
		 Else
			  Response.Write xmlUserHttp.status
		 End If
	
	End Function
	
	Function GetCategoryIDByIds(categoryIds)
		 '建立物件
		 Set xmlCategoryIDHttp = Server.CreateObject ("Microsoft.XMLHTTP")
		 '要檢查的網址
		 URLs = session("KmAPIURL") & "/category/path/"
		 dim result
		 aryCategoryId = split(categoryIds,"|")
		 IF UBound(aryCategoryId) > -1 Then
		     For i = 0 To UBound(aryCategoryId)
			     Url = URLs & aryCategoryId(i) & "?load_path=false&format=xml&tid=0&who=" & session("KMAPIActor") & "&pi=0&ps=10&api_key=" & session("KMAPIKey")
    			 
			     xmlCategoryIDHttp.Open "Get", Url, false
			     xmlCategoryIDHttp.Send
    			
			     IF xmlCategoryIDHttp.status=404 Then
			      Response.Write "找不到頁面"
			     ELSEIF xmlCategoryIDHttp.status=200 Then
				    Set categoryDocument = server.CreateObject("MSXML2.DOMDocument")
				    categoryDocument.LoadXML(xmlCategoryIDHttp.responsexml.xml)
				    Set objCategoryNodes = categoryDocument.documentElement.selectNodes("//a:CategoryId")
				    For c = 0 To objCategoryNodes.Length -1
					    IF objCategoryNodes(c).childNodes(0).text = session("CategoryTreeRootNode") Then
						    result = aryCategoryId(i)
						    Exit For
					    End IF
				    Next
				    IF result <> "" Then
				    Exit For
				    End IF
			     Else
				      Response.Write xmlCategoryIDHttp.status
			     End If
		     Next
		 Else
		 Response.Write "未傳入categoryIds"
		 End IF
		 
		 IF result <> "" Then
			GetCategoryIDByIds = result
		 Else
			GetCategoryIDByIds = Empty
		 End IF
	
	End Function

%>
