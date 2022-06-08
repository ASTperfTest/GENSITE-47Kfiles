<%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/10 上午 09:59:38
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\Search\SearchResultListTest.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("Keyword", "FromSiteUnit", "FromKnowledgeTank", "FromKnowledgeHome", "FromTopic", "AdvanceSearch", "Subject", "Keywords", "Author", "Description", "Journal", "Phonetic", "Authorize", "ActorInfoId", "CategoryId", "Range", "Order", "Sort", "Depth")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "-", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
genParamsMessage=now() & vbTab & "Error(1):REQUEST變數含特殊字元"

'-------- (可修改)只檢查單引號，如 Request 變數未來將要放入資料庫，請一定要設定(防 SQL Injection) --------
sqlInjParamsArray=array()
sqlInjParamsPattern=array("'")	'#### 要檢查的 pattern(程式會自動更新):DB ####
sqlInjParamsMessage=now() & vbTab & "Error(2):REQUEST變數含單引號"

'-------- (可修改)只檢查 HTML標籤符號，如 Request 變數未來將要輸出，請設定 (防 Cross Site Scripting)--------
xssParamsArray=array()
xssParamsPattern=array("<", ">", "%3C", "%3E")	'#### 要檢查的 pattern(程式會自動更新):HTML ####
xssParamsMessage=now() & vbTab & "Error(3):REQUEST變數含HTML標籤符號"

'-------- (可修改)檢查數字格式 --------
chkNumericArray=array()
chkNumericMessage=now() & vbTab & "Error(4):REQUEST變數不為數字格式"

'-------- (可修改)檢察日期格式 --------
chkDateArray=array()
chkDateMessage=now() & vbTab & "Error(5):REQUEST變數不為日期格式"

'##########################################
ChkPattern genParamsArray, genParamsPattern, genParamsMessage
ChkPattern sqlInjParamsArray, sqlInjParamsPattern, sqlInjParamsMessage
ChkPattern xssParamsArray, xssParamsPattern, xssParamsMessage
ChkNumeric chkNumericArray, chkNumericMessage
ChkDate chkDateArray, chkDateMessage
'--------- 檢查 request 變數名稱 --------
Sub ChkPattern(pArray, patern, message)
	for each str in pArray	
		p=request(str)
		for each ptn in patern
			if (Instr(p, ptn) >0) then
				message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
				Log4U(message) '寫入到 log
				OnErrorAction
			end if
		next
	next
End Sub

'-------- 檢查數字格式 --------
Sub ChkNumeric(pArray, message)
	for each str in pArray
		p=request(str)
		if not isNumeric(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'--------檢察日期格式 --------
Sub ChkDate(pArray, message)
	for each str in pArray
		p=request(str)
		if not IsDate(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'onError
Sub OnErrorAction()
	if (onErrorPath<>"") then response.redirect(onErrorPath)
	response.end
End Sub

'Log 放在網站根目錄下的 /Logs，檔名： YYYYMMDD_log4U.txt
Function Log4U(strLog)
	if (activeLog4U) then
		fldr=Server.mapPath("/") & "/Logs"
		filename=Year(Date()) & Right("0"&Month(Date()), 2) & Right("0"&Day(Date()),2)
		
		filename = filename & "_log4U.txt"
		
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		'產生新的目錄
		If (Not fso.FolderExists(fldr)) Then
			Set f = fso.CreateFolder(fldr)
		Else
			ShowAbsolutePath = fso.GetAbsolutePathName(fldr)
		End If
		
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		'開啟檔案
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile( fldr & "\" & filename , ForAppending, True, -1)
		f.Write strLog  & vbCrLf
	end if
End Function
'##### 結束：此段程式碼為自動產生，註解部份請勿刪除 #####
%><%
	Response.Expires = 0
	Response.Expiresabsolute = Now() - 1 
	Response.AddHeader "pragma","no-cache" 
	Response.AddHeader "cache-control","private" 
	Response.CacheControl = "no-cache"
	Server.ScriptTimeout = 600 
%>
<!--#Include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/HyftdFun.inc" -->
<%		

	on error resume next
	'-----------------------------------------------------------------------  			
	Dim QueryString, Keyword
	Dim StartDate, EndDate, HStartDate, HEndDate
	Dim FromSiteUnit, FromKnowledgeTank, FromKnowledgeHome, FromTopic
	Dim Range, Order, Sort, Depth
	Dim PageSize, PageNumber, TotalPage
	Dim CategoryId
	Dim RelatedNumber, MaxRelated
	Dim KnowledgeTankDatabaseId
	Dim AdvanceSearch
	Dim Subject, Keywords, Author, Description, Journal
	Dim Phonetic, Authorize
	Dim ActorInfoId
	'-------------------------------------------------------------------------------
	'---Keyword Setting----------------------------------------------------------------
	Keyword = Request("Keyword") 
	Keyword = Replace(Keyword, "'", "''")
	If IsNull(Keyword) Then Keyword = ""	
	'-------------------------------------------------------------------------------
	'---Date Setting----------------------------------------------------------------
	StartDate = Request("StartDate")
	If IsNull(StartDate) Or StartDate = "" Or Not IsDate(StartDate) Then
		StartDate = ""
		HStartDate = ""
	Else
		YMDArray = Split(StartDate, "/")		
		HStartDate = CompareDate( YMDArray(0), YMDArray(1), YMDArray(2) )
		YMDArray = Empty
	End If	
	EndDate = Request("EndDate")
	If IsNull(EndDate) Or EndDate = "" Or Not IsDate(EndDate) Then
		EndDate = ""
		HEndDate = ""
	Else
		YMDArray = Split(EndDate, "/")
		HEndDate = CompareDate( YMDArray(0), YMDArray(1), YMDArray(2) )		
		YMDArray = Empty
	End If		
	'-------------------------------------------------------------------------------
	'---From Setting----------------------------------------------------------------
	FromSiteUnit = Request("FromSiteUnit")
	If IsNull(FromSiteUnit) Or FromSiteUnit = "" Or (FromSiteUnit <> "0" And FromSiteUnit <> "1") Then FromSiteUnit = "0"
	FromKnowledgeTank = Request("FromKnowledgeTank")
	If IsNull(FromKnowledgeTank) Or FromKnowledgeTank = "" Or (FromKnowledgeTank <> "0" And FromKnowledgeTank <> "1") Then FromKnowledgeTank = "0"	
	FromKnowledgeHome = Request("FromKnowledgeHome")
	If IsNull(FromKnowledgeHome) Or FromKnowledgeHome = "" Or (FromKnowledgeHome <> "0" And FromKnowledgeHome <> "1") Then FromKnowledgeHome = "0"	
	FromTopic = Request("FromTopic")
	If IsNull(FromTopic) Or FromTopic = "" Or (FromTopic <> "0" And FromTopic <> "1") Then FromTopic = "0"	
	SiteUnitId = Application("SiteUnitId")														'---站內單元---1
	KnowledgeTankId = Application("KnowledgeTankId")									'---知識庫---0
	KnowledgeHomeId = Application("KnowledgeHomeId")									'---知識家---3
	TopicId = Application("TopicId")																	'---主題網---2
	KnowledgeTankDatabaseId = Application("KnowledgeTankDatabaseId")	'---知識庫可搜尋的DBID---	
	'-------------------------------------------------------------------------------
	'---SearchType Setting----------------------------------------------------------
	AdvanceSearch = Request("AdvanceSearch")
	If AdvanceSearch = "" Then AdvanceSearch = "0"	
	'-------------------------------------------------------------------------------
	'---Parameter Setting----------------------------------------------------------
	Subject = Request("Subject")
	If Subject = "" Then Subject = "0" 		
	Keywords = Request("Keywords")
	If Keywords = "" Then Keywords = "0"
	Author = Request("Author") 
	If Author = "" Then Author = "0"
	Description = Request("Description") 
	If Description = "" Then Description = "0"
	Journal = Request("Journal")
	If Journal = "" Then Journal = "0"
	Phonetic = Request("Phonetic") 
	If Phonetic = "" Then Phonetic = "0" 
	Authorize = Request("Authorize")
	If Authorize = "" Then Authorize = "0"	
	ActorInfoId = Request("ActorInfoId")
	If ActorInfoId = "" Then ActorInfoId = ""	
	'-------------------------------------------------------------------------------
	'---CategoryId Setting----------------------------------------------------------
	CategoryId = Request("CategoryId")
	If IsNull(CategoryId) Then CategoryId = ""	
	'-------------------------------------------------------------------------------
	'---Range Setting---------------------------------------------------------------
	Range = Request("Range")
	If IsNull(Range) Or Range = "" Or (Range <> "0" And Range <> "1") Then Range = "0"	
	'-------------------------------------------------------------------------------	
	'---Order Setting---------------------------------------------------------------	
	Order = Request("Order")
	If IsNull(Order) Or Order = "" Or (Order <> "0" And Order <> "1" And Order <> "2" And Order <> "3" And Order <> "4") Then Order = "0"	
	'-------------------------------------------------------------------------------
	'---Sort Setting----------------------------------------------------------------
	Sort = Request("Sort")
	If IsNull(Sort) Or Sort = "" Or Order = "0" Or (Sort <> "0" And Sort <> "1") Then Sort = "0"	
	'-------------------------------------------------------------------------------
	'---Sort Setting----------------------------------------------------------------
	Depth = Request("Depth")
	If IsNull(Depth) Or Depth = "" Or (Depth <> "0" And Depth <> "1" And Depth = "2" And Depth = "3" And Depth = "4") Then Depth = "0"	
	'-------------------------------------------------------------------------------
	'---Page Setting----------------------------------------------------------------
	PageSize = 15000
	PageNumber = 1
	TotalPage = 0
	TotalRecordCount = 0			
	'-------------------------------------------------------------------------------
  '---initial hyftd parameter-----------------------------------------------------  
 	Set HyftdObj = Server.CreateObject("hysdk.hyft.1")
 	Call Hyftd_Initial_Parameter ( PageNumber, PageSize, Debug, Order, Sort, Phonetic, Authorize )	   		  
 	'-------------------------------------------------------------------------------
	'---Set Encoding----------------------------------------------------------------
	Call Hyftd_Set_Encoding(HyftdObj, "big5")			
	'-------------------------------------------------------------------------------
	'---Connect to Hyftd Server and initial query-----------------------------------

	HyftdConnId = Hyftd_Connection( HyftdObj, HyftdServer, HyftdPort, HyftdGroupName, HyftdUserId, HyftdUserPassword )		

	Call Hyftd_Initial_Query(HyftdObj, HyftdConnId)		
	'-------------------------------------------------------------------------------
	'---Add Query Condition---------------------------------------------------------	

	'-------------------------------------------------------------------------------
	'---Data From-------------------------------------------------------------------		
	
	'Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", "1", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )								
		
	Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", "3", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )
	'Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "associatedata", "孤雌生殖", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )
	'Call Hyftd_Add_And( HyftdObj, HyftdConnId )	
	'Call Hyftd_Add_Or( HyftdObj, HyftdConnId )				
		
	'Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", "3", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )											
	
	'Call Hyftd_Add_Or( HyftdObj, HyftdConnId )		
	
	'Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", "4", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )														
	
	'Call Hyftd_Add_Or( HyftdObj, HyftdConnId )		
	
	'-------------------------------------------------------------------------------
	'---Execute Hyftd Query And Get Query Id----------------------------------------	
	HyftdSortQueryId = Hyftd_Sort_Query( HyftdObj, HyftdConnId, HyftdSortName, HyftdSortType, HyftdRecordFrom, PageSize )		
	'-------------------------------------------------------------------------------
	'---Get Total Record Count------------------------------------------------------
	HyftdTotalRecordCount = Hyftd_Total_Sysid( HyftdObj, HyftdConnId, HyftdSortQueryId )	
	'-------------------------------------------------------------------------------
	'---Get Current Record Count----------------------------------------------------
	HyftdCurrentRecordCount = Hyftd_Num_Sysid( HyftdObj, HyftdConnId, HyftdSortQueryId )		
	'-------------------------------------------------------------------------------
	'---Get Record Content----------------------------------------------------------			
	If HyftdCurrentRecordCount > 0 Then		
		Dim index					
		dim counter
		For index = 0 To HyftdCurrentRecordCount - 1										
			
			'---------------------------------------------------------------------------
			'---Fetch Current Sysid-----------------------------------------------------
			HyftdDataId = Hyftd_Fetch_Sysid( HyftdObj, HyftdConnId, HyftdSortQueryId, index )				
			'response.write HyftdDataId & "<br />"
      '---------------------------------------------------------------------------
			'sql = ""
			'sysid = split(HyftdDataId, "@")
			
			'for each item in sysid
			'	sql = sql & item & ","				
			'next
			'sql = sql & "0,"
			'sql = "INSERT INTO HyftdIndexDelete VALUES(" & left(sql, len(sql) - 1) & ",GETDATE(),0)"
			'conn.execute(sql)
			arr = split(HyftdDataId, "@")
			sql = "SELECT CuDTGeneric.sTitle, CuDTGeneric.ClickCount, CONVERT(char(10), CuDTGeneric.xPostDate, 111) AS xPostDate, CtUnit.CtUnitName, CuDTGeneric.showType, " & _
      				"ISNULL(CuDTGeneric.xURL, '') AS xURL, ISNULL(CuDTGeneric.fileDownLoad, '') AS fileDownLoad, CatTreeNode.CtNodeID, SUBSTRING(ISNULL(CuDTGeneric.xBody, ''), 0, 60) AS xBody, CatTreeNode.CtRootID " & _
      				"FROM CuDTGeneric INNER JOIN CtUnit ON CuDTGeneric.iCTUnit = CtUnit.CtUnitID INNER JOIN CatTreeNode " & _
      				"ON CtUnit.CtUnitID = CatTreeNode.CtUnitID WHERE (CuDTGeneric.iCUItem = '" & arr(1) & "') "      			
      	if arr(0) = "1" then
      		sql = sql & "AND (CatTreeNode.CtRootID = 34) "
					sql = sql & "AND (CatTreeNode.CtNodeID = " & arr(4) & ")"				
      	End If		      
				
	      Set cudtrs = conn.execute(sql)
	      If cudtrs.EOF Then
					counter = counter + 1
	      	'response.write "<font color=""red"">" & HyftdDataId & "</font><br />"
					sysid = split(HyftdDataId, "@")
					delsql = ""
					for each item in sysid
						delsql = delsql & item & ","				
					next
					'delsql = delsql & "0,"
					delsql1 = "INSERT INTO HyftdIndexDelete VALUES(" & left(delsql, len(delsql) - 1) & ",GETDATE(),0)"
					response.write delsql1 & "<br />"
					'conn.execute(delsql1)
	      End If      
	      cudtrs.close
	      Set cudtrs = Nothing
			
			'response.write sql & "<hr />"
		Next					
		response.write index & "<hr />"
		response.write counter & "<hr />"
  End If  	

	'-----------------------------------------------------------------
	'---word tank------------------------------------------------------
	'-----------------------------------------------------------------      
	'End If		
	if err.number = 0 then
		HyftdX = HyftdObj.hyft_close(HyftdConnId)    
		HyftdObj = Empty
		'response.write err.description & " at: " & err.number
 		'Set HyftdObj = Nothing  
 	else
 		HyftdX = HyftdObj.hyft_close(HyftdConnId)    
		HyftdObj = Empty
 		'Set HyftdObj = Nothing  
 		response.write err.description & " at: " & err.number
	end if				
%>
