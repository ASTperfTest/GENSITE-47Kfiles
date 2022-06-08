<%@CodePage = 65001%><%
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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\Search\HSearchResultList.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("Keyword", "FromSiteUnit", "FromKnowledgeTank", "FromKnowledgeHome", "FromTopic", "AdvanceSearch", "Subject", "Keywords", "Author", "Description", "Journal", "Phonetic", "Authorize", "ActorInfoId", "debug", "CategoryId", "Range", "Order", "Sort", "Depth", "PageSize", "PageNumber", "RelatedNumber", "MaxRelated", "gstyle")
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
%>
<%
	Response.Expires = 0
	Response.Expiresabsolute = Now() - 1 
	Response.AddHeader "pragma","no-cache" 
	Response.AddHeader "cache-control","private" 
	Response.CacheControl = "no-cache"
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
	Dim debug
	'-------------------------------------------------------------------------------
	'---Keyword Setting----------------------------------------------------------------
	Keyword = Request("Keyword") 
	if Instr(Keyword , ";") = Len(Keyword) then
		Keyword = Left(Keyword , Len(Keyword)-1)
	end if
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
	debug = Request("debug")
	If debug = "" Then debug = ""	
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
	PageSize = Request("PageSize")
	PageNumber = Request("PageNumber")			
	If IsNull(PageSize) Or PageSize = "" Or (PageSize <> "10" And PageSize <> "30" And PageSize <> "50") Then PageSize = 10	
	If IsNull(PageNumber) Or PageNumber = "" Or CInt(PageNumber) <= 0 Then PageNumber = 1	
	TotalPage = 0
	TotalRecordCount = 0			
	'-------------------------------------------------------------------------------
  '---other parameter-------------------------------------------------------------  
  RelatedNumber = Request("RelatedNumber")
  If IsNull(RelatedNumber) Or RelatedNumber = "" Then RelatedNumber = "0"
  MaxRelated = Request("MaxRelated") 
	If IsNull(MaxRelated) Or MaxRelated = "" Then MaxRelated = "1"	
	'-------------------------------------------------------------------------------
	'---放置由hyftd回傳的結果-------------------------------------------------------
	Dim ResultArray()  
	ReDim ResultArray(PageSize, 18)  
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
	Dim BolDone
	Dim StrQuery
	KeywordArray = Split(Keyword, ";")
	If UBound(KeywordArray) > 0 Then
		For i = 0 To UBound(KeywordArray)
			If i = 0 Then
				StrQuery = KeywordArray(i)
			Else
				StrQuery = StrQuery & " * " & KeywordArray(i)
			End If		
		Next
	Else
		StrQuery = Keyword
	End If			
	'---for query---
	'-------------------------------------------------------------------------------
	'---Data From-------------------------------------------------------------------		
	Dim SiteFlag
	SiteFlag = false	
	If FromKnowledgeTank = "1" Then	'---知識庫---若有抓取知識庫的資料,只能抓取DB020的----------		
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", KnowledgeTankId, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )								
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "databaseid", KnowledgeTankDatabaseId, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
		Call Hyftd_Add_And( HyftdObj, HyftdConnId )				
		If SiteFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )		
		SiteFlag = true
	End If		
	If FromSiteUnit = "1" Then			'---站內單元---
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", SiteUnitId, 			"", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )			
		If SiteFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )
		SiteFlag = true		
	End If				
	If FromKnowledgeHome = "1" Then	'---知識家---		
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", KnowledgeHomeId, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )											
		If SiteFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )		
		SiteFlag = true
	End If	
	If FromTopic = "1" Then					'---主題館---		
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", TopicId, 				"", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )														
		If SiteFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )		
		SiteFlag = true
	End If		
	'-------------------------------------------------------------------------------
	'---Keyword---------------------------------------------------------------------
	Dim TypeFlag
	If AdvanceSearch = "1" Then		
		If Subject = "1" Then
			Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "subject", 		StrQuery, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )			
			TypeFlag = true
		End If
		If Keywords = "1" Then
			Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "keywords", 		StrQuery, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )			
			If TypeFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )			
			TypeFlag = true
		End If
		If Author = "1" Then
			Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "author", 		StrQuery, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )			
			If TypeFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )			
			TypeFlag = true
		End If
		If Description = "1" Then
			Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "description", 		StrQuery, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )			
			If TypeFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )			
			TypeFlag = true
		End If
		If Journal = "1" Then
			Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "journal", 		StrQuery, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )			
			If TypeFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )			
			TypeFlag = true
		End If		
	Else	
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "associatedata", StrQuery, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )
	End If
	
	If SiteFlag = true Then Call Hyftd_Add_And( HyftdObj, HyftdConnId )	
			
	'-------------------------------------------------------------------------------
	'---PostDate Range--------------------------------------------------------------
	If HStartDate <> "" Or HEndDate <> "" Then
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "onlinedate", HStartDate, HEndDate, HyftdQueryType, HyftdPhonetic, HyftdAuthorize )
		Call Hyftd_Add_And( HyftdObj, HyftdConnId )
	End If	
	'-------------------------------------------------------------------------------	
	If CategoryId <> "" Then		
		If Depth = "1" Then	'---siteid---				
			'Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", CategoryId & "$", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
			'Call Hyftd_Add_And( HyftdObj, HyftdConnId )				
		ElseIf Depth = "2" Then
			If FromTopic = "1" Then
				Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "categoryname", CategoryId, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
				Call Hyftd_Add_And( HyftdObj, HyftdConnId )				
			End If
			If FromKnowledgeHome = "1" Then
				Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "categoryname", CategoryId & "$", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
				Call Hyftd_Add_And( HyftdObj, HyftdConnId )				
			End If
			If FromSiteUnit = "1" Then
				Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "categoryid", CategoryId & "$", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
				Call Hyftd_Add_And( HyftdObj, HyftdConnId )				
			End If
			If FromKnowledgeTank = "1" Then
				'Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "categoryid", CategoryId & "$", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
				Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "actorinfoid", CategoryId, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
				Call Hyftd_Add_And( HyftdObj, HyftdConnId )				
			End If
		ElseIf Depth = "3" Then
			If FromTopic = "1" Then
				Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "categoryid", CategoryId & "$", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
				Call Hyftd_Add_And( HyftdObj, HyftdConnId )			
			End If
			If FromKnowledgeTank = "1" Then
				'Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "categoryid", CategoryId & "$", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
				'Call Hyftd_Add_And( HyftdObj, HyftdConnId )				
			End If
		ElseIf Depth = "4" Then
			If FromKnowledgeTank = "1" Then	'---categoryid---											
				Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "categoryid", CategoryId & "$", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
				Call Hyftd_Add_And( HyftdObj, HyftdConnId )					
			End If			
		End If
	End If	
	'---加入actorinfoid---
	If FromKnowledgeTank = "1" Then
		If ActorInfoId <> "" Then
			Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "actorinfoid", ActorInfoId, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
			Call Hyftd_Add_And( HyftdObj, HyftdConnId )					
		Else
			If Depth = "0" Then
				If request("gstyle") = "003" or request("gstyle") = "005" then
					Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "actorinfoid", "000 + 001 + 002 + 004", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
				else
					Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "actorinfoid", "000 + 001 + 002", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
				end if					
				'Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "actorinfoid", "002", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )															
				'Call Hyftd_Add_Or( HyftdObj, HyftdConnId )					
				'Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "actorinfoid", "002", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )															
				'Call Hyftd_Add_Or( HyftdObj, HyftdConnId )					
				Call Hyftd_Add_And( HyftdObj, HyftdConnId )					
			Else				
				If request("gstyle") = "003" or request("gstyle") = "005" then
					Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "actorinfoid", "001 + 002 + 004", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )																			
				else
					Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "actorinfoid", "001 + 002", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )																			
				end if 				
				Call Hyftd_Add_And( HyftdObj, HyftdConnId )					
			End If			
		End If
	End If
	'-------------------------------------------------------------------------------
	'---Execute Hyftd Query And Get Query Id----------------------------------------	
	HyftdSortQueryId = Hyftd_Sort_Query( HyftdObj, HyftdConnId, HyftdSortName, HyftdSortType, HyftdRecordFrom, PageSize )		
	'-------------------------------------------------------------------------------


'response.write GetFieldIndexString(HyftdObj,HyftdConnId, HyftdSortQueryId, "actorinfoid")


	'---Get Total Record Count------------------------------------------------------
	HyftdTotalRecordCount = Hyftd_Total_Sysid( HyftdObj, HyftdConnId, HyftdSortQueryId )	
	'-------------------------------------------------------------------------------
	'---Get Current Record Count----------------------------------------------------
	HyftdCurrentRecordCount = Hyftd_Num_Sysid( HyftdObj, HyftdConnId, HyftdSortQueryId )		
	'-------------------------------------------------------------------------------
	'---Get Record Content----------------------------------------------------------
	If HyftdCurrentRecordCount > 0 Then		
		Dim index					
		For index = 0 To HyftdCurrentRecordCount - 1										
			If ((PageNumber - 1) * PageSize + index) = HyftdTotalRecordCount Then				
				Exit For
			End If
			'---------------------------------------------------------------------------
			'---Fetch Current Sysid-----------------------------------------------------
			HyftdDataId = Hyftd_Fetch_Sysid( HyftdObj, HyftdConnId, HyftdSortQueryId, index )				
			'if request("debug") = "true" then
			'	response.write HyftdDataId & "<hr>"
			'end if
      '---------------------------------------------------------------------------
      '---Get Related Number------------------------------------------------------
      HyftdDataRelatedNumber = Hyftd_Get_Related_Number( HyftdObj )
      '---------------------------------------------------------------------------
      '---Set Value Into ResultArray----------------------------------------------             
      ResultArray(index, 0) = HyftdDataId
      ResultArray(index, 1) = HyftdDataRelatedNumber
      TempArray = Split( HyftdDataId, "@")
      ResultArray(index, 2) = TempArray(0)	'---siteid---
      ResultArray(index, 3) = TempArray(1)	'---reportid---icuitem
      ResultArray(index, 4) = TempArray(2)	'---databaseid---basedsd
      ResultArray(index, 5) = TempArray(3)	'---categoryid---ctunit      
      ResultArray(index, 6) = ""
	    ResultArray(index, 7) = ""
	    ResultArray(index, 8) = ""
	    ResultArray(index, 9) = ""	      	
	    ResultArray(index, 10) = ""
	    ResultArray(index, 11) = ""
	    ResultArray(index, 12) = ""
	    ResultArray(index, 13) = ""
	    ResultArray(index, 14) = ""
	    ResultArray(index, 15) = ""
	    ResultArray(index, 16) = "" '---actorinfoid---
			if TempArray(0) <> "0" then
				ResultArray(index, 17) = TempArray(4) '---ctnodeid---
			else
				ResultArray(index, 17) = ""
			end if
      If TempArray(0) = "0" Then  

	  
      	sql = "SELECT REPORT.REPORT_ID, ISNULL(REPORT.SUBJECT, '') AS SUBJECT, ISNULL(REPORT.CLICK_COUNT, 0) AS CLICK_COUNT, " & _
	    				"CONVERT(char(10), REPORT.ONLINE_DATE, 111) AS ONLINE_DATE, ISNULL(CATEGORY.CATEGORY_NAME, '') AS CATEGORY_NAME " & _
	    				"FROM REPORT INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID " & _
	    				"INNER JOIN DATA_BASE ON CAT2RPT.DATA_BASE_ID = DATA_BASE.DATA_BASE_ID LEFT OUTER JOIN CATEGORY " & _
	    				"ON CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID AND CAT2RPT.DATA_BASE_ID = CATEGORY.DATA_BASE_ID " & _
	    				"WHERE (REPORT.REPORT_ID = '" + ResultArray(index, 3) + "') AND (CAT2RPT.DATA_BASE_ID = '" + ResultArray(index, 4) + "') " & _
	    				"AND (CAT2RPT.CATEGORY_ID = '" + ResultArray(index, 5) + "')"
	    	Set kmrs = KMConn.execute(sql)	    		      
	    	If Not kmrs.EOF Then
	    		ResultArray(index, 6) = kmrs("CATEGORY_NAME")
	      	ResultArray(index, 7) = kmrs("SUBJECT")
	      	ResultArray(index, 8) = kmrs("CLICK_COUNT")
	      	ResultArray(index, 9) = kmrs("ONLINE_DATE")	      	
	    	End If
	    	
	    	kmrs.close
	    	Set kmrs = Nothing
      	
      ElseIf TempArray(0) = "1" Or TempArray(0) = "2" Or TempArray(0) = "3" Then
      	
      	sql = "SELECT CuDTGeneric.sTitle, CuDTGeneric.ClickCount, CONVERT(char(10), CuDTGeneric.xPostDate, 111) AS xPostDate, CtUnit.CtUnitName, CuDTGeneric.showType, " & _
      				"ISNULL(CuDTGeneric.xURL, '') AS xURL, ISNULL(CuDTGeneric.fileDownLoad, '') AS fileDownLoad, CatTreeNode.CtNodeID, SUBSTRING(ISNULL(CuDTGeneric.xBody, ''), 0, 60) AS xBody, CatTreeNode.CtRootID " & _
      				"FROM CuDTGeneric INNER JOIN CtUnit ON CuDTGeneric.iCTUnit = CtUnit.CtUnitID INNER JOIN CatTreeNode " & _
      				"ON CtUnit.CtUnitID = CatTreeNode.CtUnitID  " & _
					" WHERE (CuDTGeneric.iCUItem = '" & ResultArray(index, 3) & "') "      			
      	If TempArray(0) = "1" Then
      		sql = sql & "AND (CatTreeNode.CtRootID = " & Application("SiteCatTreeRoot") & ") "
					sql = sql & "AND (CatTreeNode.CtNodeID = " & ResultArray(index, 17) & ")"
      	End If		      
	      Set cudtrs = conn.execute(sql)
	      If Not cudtrs.EOF Then
	      	ResultArray(index, 6) = cudtrs("CtUnitName")
	      	ResultArray(index, 7) = cudtrs("sTitle")
	      	ResultArray(index, 8) = cudtrs("ClickCount")
	      	ResultArray(index, 9) = cudtrs("xPostDate")
	      	ResultArray(index, 10) = cudtrs("showType")
	      	ResultArray(index, 11) = cudtrs("xURL")
	      	ResultArray(index, 12) = cudtrs("fileDownLoad")
	      	ResultArray(index, 13) = cudtrs("CtNodeID")
	      	ResultArray(index, 14) = cudtrs("xBody")
	      	ResultArray(index, 15) = cudtrs("CtRootID")	      
	      End If      
	      cudtrs.close
	      Set cudtrs = Nothing
    	End If                              
      If index = 0 And RelatedNumber = "0" Then          	    	      	
       	MaxRelated = HyftdDataRelatedNumber
       	RelatedNumber = "1"
      End If  
      k = Hyftd_One_Assdata_Len(HyftdObj, HyftdConnId,"brief",HyftdDataId)    	      
      f = Hyftd_One_Assdata(HyftdObj, HyftdConnId,k)  
			if instr(f, "001") > 0 and instr(f, "002") > 0 then
				ResultArray(index, 16) = "002"
			elseif instr(f, "001") > 0 then
				ResultArray(index, 16) = "001"
			elseif instr(f, "002") > 0 then
				ResultArray(index, 16) = "002"
			end if
      'ResultArray(index, 16) = Mid(f, InStr(f, "<actorinfoid>") + 13, 3)      
		Next					
  End If  	
	
	'-------------------------------------------------------------------------------
	'---setting page number---------------------------------------------------------
	If HyftdTotalRecordCount <> 0 Then TotalPage = Int(HyftdTotalRecordCount / PageSize + 0.99)			
	'-------------------------------------------------------------------------------
	'---Consume Time----------------------------------------------------------------
	HyftdEndTime = GetTime()
	HyftdTimeDiff = Round((HyftdEndTime / 1000) - (HyftdStartTime / 1000), 3)				
	'-------------------------------------------------------------------------------				
	'---using grouping name---
	If Depth = "0" Then
		HyftdGroupingName1 = "siteid"
	ElseIf Depth = "1" Then
		If FromTopic = "1" Then
			HyftdGroupingName1 = "categoryname"
		ElseIf FromKnowledgeHome = "1" Then
			HyftdGroupingName1 = "categoryname"
		ElseIf FromKnowledgeTank = "1" Then
			HyftdGroupingName1 = "actorinfoid"
		Else
			HyftdGroupingName1 = "categoryid"
		End If
	ElseIf Depth = "2" Then
		If FromKnowledgeHome = "1" Then
			HyftdGroupingName1 = "categoryname"
		Else
			HyftdGroupingName1 = "categoryid"
		End If
	ElseIf Depth = "3" Then
		HyftdGroupingName1 = "categoryid"
	ElseIf Depth = "4" Then		
		HyftdGroupingName1 = "categoryid"
	End If					
	'---using grouping len---
	Dim GroupingLen
	GroupingLen = "0"							
	If FromKnowledgeTank = "1" Then		
		If Depth = "0" Then
			GroupingLen = "0"
		ElseIf Depth = "1" Then
			If FromKnowledgeTank = "1" Then
				GroupingLen = "3"	'---for actorinfoid---
				'GroupingLen = "9"
			End If
		ElseIf Depth = "2" Then
			If FromKnowledgeTank = "1" Then
				GroupingLen = "11"
			End If
		ElseIf Depth = "3" Then
			If FromKnowledgeTank = "1" Then
				GroupingLen = "13"
			End If
		ElseIf Depth = "4" Then
			If FromKnowledgeTank = "1" Then
				GroupingLen = "13"
			End If
		End If		
	End If	

	'---get grouping---
	HyftdQueryGroupingId = Hyftd_Query_Grouping( HyftdObj, HyftdConnId, HyftdGroupingName1, GroupingLen, HyftdGroupingType1 )
	'---get how many group---
	HyftdGroupingCount1 = Hyftd_Num_Grouping( HyftdObj, HyftdConnId, HyftdQueryGroupingId )		
	ReDim HyftdGrouping1(HyftdGroupingCount1, 3)	
	'-------------------------------------------------------------------------------
	'response.write HyftdGroupingCount1 & "~" & hyftdgroupingname1 & "~" & groupinglen	
	Dim pathGroupName(5)					
	Dim pathGroupData(5)
	'pathGroupName = Array("","","","","")
	'pathGroupData = Array("","","","","")	
	'---上方分類路徑-----------------------------------------------------------
	If CInt(Depth) >= 0 Then
		pathGroupName(0) = "全部"
		pathGroupData(0) = ""	
	End If
	If CInt(Depth) >= 1 Then			
		If FromKnowledgeTank = "1" Then 
			pathGroupName(1) = "知識庫"
			pathGroupData(1) = "0"				
		End If
		If FromSiteUnit = "1" Then
			pathGroupName(1) = "站內單元"
			pathGroupData(1) = "1"				
		End If
		If FromTopic = "1" Then
			pathGroupName(1) = "主題網"
			pathGroupData(1) = "2"				
		End If
		If FromKnowledgeHome = "1" Then
			pathGroupName(1) = "知識家"
			pathGroupData(1) = "3"				
		End If				
	End If	
	If CInt(Depth) >= 2 Then
		pathGroupData(2) = CategoryId
		If FromSiteUnit = "1" Then
			temp = split(CategoryId, "@")
			sql = "SELECT CtUnitName FROM CtUnit WHERE (CtUnitID = " & temp(1) & ")"
			Set siters = conn.execute(sql)
			if not siters.eof then pathGroupName(2) = siters("CtUnitName")			
			siters = nothing
		End If
		If FromKnowledgeTank = "1" Then
			If CInt(Depth) = 2 Then
				ActorInfoId = CategoryId	
			End If
			pathGroupData(2) = ActorInfoId
			If ActorInfoId = "001" Then
				pathGroupName(2) = "生產者"
			ElseIf ActorInfoId = "002" Then
				pathGroupName(2) = "消費者"
			End If
			'temp = split(CategoryId, "@")
			'sql = "SELECT CATEGORY_NAME FROM CATEGORY WHERE (DATA_BASE_ID = '" & temp(0) & "') AND (CATEGORY_ID = '" & Mid(temp(1), 1, 3) & "')"						
			'Set tankrs = kmconn.execute(sql)
			'if not tankrs.eof then
			'	pathGroupName(2) = tankrs("CATEGORY_NAME")
			'end if
			'tankrs = nothing
		End If
		If FromKnowledgeHome = "1" Then
			temp = split(CategoryId, "@")
			sql = "SELECT mValue FROM CodeMain WHERE (codeMetaID = 'KnowledgeType') AND (mCode = '" & temp(1) & "')"				
			Set homers = conn.execute(sql)
			if not homers.eof then pathGroupName(2) = homers("mValue")			
			homers = nothing
		End If
		If FromTopic = "1" Then			
			If Depth = "2" Then	
				sql = "SELECT CtRootName FROM CatTreeRoot WHERE (CtRootID = " & CategoryId & ")"				
				Set topicrs = conn.execute(sql)
				if not topicrs.eof then pathGroupName(2) = topicrs("CtRootName")				
				topicrs = nothing				
			Else
				temp = split(CategoryId, "@")
				sql = "SELECT  CatTreeRoot.CtRootName, CatTreeRoot.CtRootID FROM CatTreeNode INNER JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID INNER JOIN " & _
							"CatTreeRoot ON CatTreeNode.CtRootID = CatTreeRoot.CtRootID WHERE (CatTreeNode.CtUnitID = " & temp(1) & ")"								
				Set topicrs = conn.execute(sql)								
				if not topicrs.eof then
					pathGroupName(2) = topicrs("CtRootName")
					pathGroupData(2) = topicrs("CtRootID")
				end if
				topicrs = nothing									
			End If		
		End If		
	End If	
	If CInt(Depth) >= 3 Then
		pathGroupData(3) = CategoryId
		If FromKnowledgeTank = "1" Then
			temp = split(CategoryId, "@")
			sql = "SELECT CATEGORY_NAME FROM CATEGORY WHERE (DATA_BASE_ID = '" & temp(0) & "') AND (CATEGORY_ID = '" & Mid(temp(1), 1, 5) & "')"				
			Set tankrs = kmconn.execute(sql)
			if not tankrs.eof then pathGroupName(3) = tankrs("CATEGORY_NAME")			
			tankrs = nothing
		End If		
		If FromTopic = "1" Then
			temp = split(CategoryId, "@")
			sql = "SELECT CtUnitName FROM CtUnit WHERE (CtUnitID = " & temp(1) & ")"				
			Set topicrs = conn.execute(sql)
			if not topicrs.eof then pathGroupName(3) = topicrs("CtUnitName")			
			topicrs = nothing
		End If		
	End If		
	If CInt(Depth) >= 4 Then
		pathGroupData(4) = CategoryId
		If FromKnowledgeTank = "1" Then
			temp = split(CategoryId, "@")
			sql = "SELECT CATEGORY_NAME FROM CATEGORY WHERE (DATA_BASE_ID = '" & temp(0) & "') AND (CATEGORY_ID = '" & temp(1) & "')"				
			Set tankrs = kmconn.execute(sql)
			if not tankrs.eof then pathGroupName(4) = tankrs("CATEGORY_NAME")			
			tankrs = nothing
		End If					
	End If		
	'---end of 上方分類路徑-----------------------------------------------------------
	'---for 後分類--------------------------------------------------------------------
	Dim CoaItem, KmItem
	If Depth = "0" Then
		If HyftdGroupingCount1 > 0 Then
			For index = 0 To HyftdGroupingCount1 - 1
				'---get the name and number of grouping---
				HyftdGrouping1(index, 1) = Hyftd_Fetch_Grouping( HyftdObj, HyftdConnId, HyftdQueryGroupingId, index )
				HyftdGrouping1(index, 0) = Hyftd_Get_Grouping_Name( HyftdObj )							
			Next
		End If	
	Else
		If HyftdGroupingCount1 > 0 Then
			For index = 0 To HyftdGroupingCount1 - 1
				HyftdGrouping1(index, 1) = Hyftd_Fetch_Grouping( HyftdObj, HyftdConnId, HyftdQueryGroupingId, index )
				HyftdGrouping1(index, 0) = Hyftd_Get_Grouping_Name( HyftdObj )									
				'response.write HyftdGrouping1(index, 0) & "~" & HyftdGrouping1(index, 1) & "<hr>"
				'If FromTopic = "1" And Depth = "1" Then
				'	CoaItem = CoaItem & HyftdGrouping1(index, 0) & ","
				'Else
				If HyftdGrouping1(index, 0) = "001" Then
					HyftdGrouping1(index, 2) = "生產者"
				ElseIf HyftdGrouping1(index, 0) = "002" Then
					HyftdGrouping1(index, 2) = "消費者"
				Else
					'response.write HyftdGrouping1(index, 0) & "<hr>"
					TempArray = Split(HyftdGrouping1(index, 0), " ")
					If FromKnowledgeTank = "1" Then					
						KmItem = KmItem & "'" & TempArray(1) & "',"							
					ElseIf FromTopic = "1" And Depth = "1" Then
						CoaItem = CoaItem & TempArray(0) & ","
					ElseIf FromKnowledgeHome = "1" And (Depth = "1" Or Depth = "2") Then
						CoaItem = CoaItem & "'" & TempArray(1) & "',"
					Else
						CoaItem = CoaItem & TempArray(1) & ","
					End If					
				End If
				'End If
			Next			
			If Len(CoaItem) > 0 Then CoaItem = Mid(CoaItem, 1, Len(CoaItem) - 1)
			If Len(KmItem) > 0 Then KmItem = Mid(KmItem, 1, Len(KmItem) - 1)
		End If										
		'---Connect To DB To Get The Category Name--------------------------------------
		Dim CategoryName()
		Dim CountC
		CountC = 0
		If Depth <> "0" Then
			If Len(CoaItem) > 0 Then
				If FromTopic = "1" And Depth = "1" Then
					sql = "SELECT CtRootID AS CtUnitID, CtRootName AS CtUnitName FROM CatTreeRoot " & _
								"WHERE CtRootID IN (" & CoaItem & ")"		
				ElseIf FromKnowledgeHome = "1" And (Depth = "1" Or Depth = "2" ) Then
					sql = "SELECT mCode AS CtUnitID, mValue AS CtUnitName FROM CodeMain WHERE codeMetaID = 'KnowledgeType' " & _
								"AND mCode IN (" & CoaItem & ")"
				Else
					sql = "SELECT CtUnitName, CtUnitID FROM CtUnit WHERE CtUnitID IN (" & CoaItem & ")"
				End If					
				Set CoaRs = Conn.Execute(sql)
				While Not CoaRs.EOF 
					CountC = CountC + 1				
					CoaRs.MoveNext
				Wend			
			End If
			If Len(KmItem) > 0 Then
				sql = "SELECT CATEGORY_ID, CATEGORY_NAME FROM CATEGORY WHERE DATA_BASE_ID = '" & KnowledgeTankDatabaseId & "' AND CATEGORY_ID IN (" & KmItem & ")"
				Set KmRs = KmConn.Execute(sql)
				While Not KmRs.EOF 
					CountC = CountC + 1								
					KmRs.MoveNext
				Wend			
			End If
			If Len(CoaItem) + Len(KmItem) > 0 Then
				ReDim CategoryName(CountC, 2)
				CountC = 0
				If Len(CoaItem) > 0 Then
					CoaRs.MoveFirst
					While Not CoaRs.EOF
						CategoryName(CountC, 0) = CStr(CoaRs("CtUnitID"))
						CategoryName(CountC, 1) = CoaRs("CtUnitName")							
						'response.write CategoryName(CountC, 0) & "~" & CategoryName(CountC, 1) & "<hr>"
						CountC = CountC + 1
						CoaRs.MoveNext 				
					Wend
					CoaRs.close
					Set CoaRs = Nothing
				End If
				If Len(KmItem) > 0 Then
					KmRs.MoveFirst
					While Not KmRs.EOF			
						CategoryName(CountC, 0) = KmRs("CATEGORY_ID")
						CategoryName(CountC, 1) = KmRs("CATEGORY_NAME")							
						'response.write CategoryName(CountC, 0) & "~" & CategoryName(CountC, 1) & "<hr>"
						CountC = CountC + 1
						KmRs.MoveNext			
					Wend
					KmRs.close
					Set KmRs = Nothing
				End If
				If CountC > 0 Then									
					For index = 0 To HyftdGroupingCount1 - 1
						'If FromTopic = "1" And Depth = "1" Then
						'	TempArray(0) = ""
						'	TempArray(1) = HyftdGrouping1(index, 0 )
						'Else
						TempArray = Split(HyftdGrouping1(index, 0 ), " ")
						'End If														
						For index1 = 0 To CountC - 1			
					
							If FromKnowledgeTank = "1" Then					
								If TempArray(1) = CategoryName(index1, 0) Then
									'response.write HyftdGrouping1(index, 0 ) & "~" & CategoryName(index1, 0) & "<hr>"
									HyftdGrouping1(index, 2) = CategoryName(index1, 1)
									Exit For
								End If								
							ElseIf FromTopic = "1" And Depth = "1" Then
								If TempArray(0) = CategoryName(index1, 0) Then
									'response.write HyftdGrouping1(index, 0 ) & "~" & CategoryName(index1, 0) & "<hr>"
									HyftdGrouping1(index, 2) = CategoryName(index1, 1)
									Exit For
								End If	
							ElseIf FromKnowledgeHome = "1" And Depth = "1" Then
								If TempArray(1) = CategoryName(index1, 0) Then
									'response.write HyftdGrouping1(index, 0 ) & "~" & CategoryName(index1, 0) & "<hr>"
									HyftdGrouping1(index, 2) = CategoryName(index1, 1)
									Exit For
								End If	
							Else
								If TempArray(1) = CategoryName(index1, 0) Then
									'response.write HyftdGrouping1(index, 0 ) & "~" & CategoryName(index1, 0) & "<hr>"
									HyftdGrouping1(index, 2) = CategoryName(index1, 1)
									Exit For
								End If	
							End If					
							
							'If TempArray(1) = CategoryName(index1, 0) Then
								'response.write HyftdGrouping1(index, 0 ) & "~" & CategoryName(index1, 0) & "<hr>"
								'	HyftdGrouping1(index, 2) = CategoryName(index1, 1)
								'	Exit For
							'End If			
						Next		
					Next							
				End If
			End If
		End If	
	End If		
	'-------------------------------------------------------------------------------		
	'---Time Distribute-------------------------------------------------------------
	HyftdQueryGroupingId = Hyftd_Query_Grouping( HyftdObj, HyftdConnId, HyftdGroupingName, "0", HyftdGroupingType )
	HyftdGroupingCount = Hyftd_Num_Grouping( HyftdObj, HyftdConnId, HyftdQueryGroupingId )
	ReDim HyftdGrouping(HyftdGroupingCount, 2)
	If HyftdGroupingCount > 0 Then
		For index = 0 To HyftdGroupingCount - 1
			HyftdGrouping(index, 1) = Hyftd_Fetch_Grouping( HyftdObj, HyftdConnId, HyftdQueryGroupingId, index )
			HyftdGrouping(index, 0) = Hyftd_Get_Grouping_Name( HyftdObj )
		Next			
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
 		'response.write err.description & " at: " & err.number
	end if				
%>		
		<% For index = 0 To PageSize - 1 %>
			<% If ResultArray(index, 0) <> "" And index < 10 Then
			
				If ResultArray(index, 2) = "3" And ResultArray(index, 7) <> "" Then	'---知識家---
				Response.write ("<doclist>")
			%>
					<%  Response.write ("<icuitem>")
						Response.Write (ResultArray(index, 3))
						Response.write ("</icuitem>")
						Response.write ("<title><![CDATA[")
						HomeUrl = Replace( session("myURL") & "knowledge/knowledge_cp.aspx?ArticleId={0}&ArticleType=A&CategoryId=E", "{0}", ResultArray(index, 3) )								
						'Response.Write "<a href=""" & HomeUrl & """ target=""_blank"">" & ResultArray(index, 7) & "</a>"
						Response.Write (ResultArray(index, 7))
						Response.write ("]]></title>")	
						Response.write ("<postTime>")
						Response.Write (ResultArray(index, 9))	
						Response.write ("</postTime>")						
					%>
				<% Response.write ("</doclist>") 
				End If%>
			<% Else %>
				<% Exit For %>
			<% End If %>
		<% Next%>
<%
	conn.close
	Set conn = Nothing
%>