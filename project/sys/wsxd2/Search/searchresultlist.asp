<%@CodePage = 65001%>
<%
	Response.Expires = 0
	Response.Expiresabsolute = Now() - 1 
	Response.AddHeader "pragma","no-cache" 
	Response.AddHeader "cache-control","private" 
	Response.CacheControl = "no-cache"
%>
<!--#include virtual = "/inc/HyftdFun.inc" -->
<!--#include virtual = "/inc/MSClient.inc" -->
<%		



	'on error resume next
	'-----------------------------------------------------------------------  			
	Dim QueryString, Keyword
	Dim StartDate, EndDate, HStartDate, HEndDate
	Dim FromSiteUnit, FromKnowledgeTank, FromKnowledgeHome, FromTopic
	Dim FromPedia,FromVideo,FromTechCD '新增加三種
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
	Dim index					
	Dim NullRecordCount
	NullRecordCount = 0
	
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
													'---站內單元---1
	FromPedia = Request("FromPedia")
	If IsNull(FromPedia) Or FromPedia = "" Or (FromPedia <> "0" And FromPedia <> "1") Then FromPedia = "0"	
	
	FromVideo = Request("FromVideo")
	If IsNull(FromVideo) Or FromVideo = "" Or (FromVideo <> "0" And FromVideo <> "1") Then FromVideo = "0"	
	
	FromTechCD = Request("FromTechCD")
	If IsNull(FromTechCD) Or FromTechCD = "" Or (FromTechCD <> "0" And FromTechCD <> "1") Then FromTechCD = "0"	
   
  
	'==================================
    Application("Pedia") = "4"
	SiteUnitId 		        = Application("SiteUnitId")	
	KnowledgeTankId  = Application("KnowledgeTankId")									    '---知識庫---0
	KnowledgeHomeId = Application("KnowledgeHomeId")								    '---知識家---3
	TopicId 		        = Application("TopicId")											        '---主題網---2
	PediaId			        = Application("Pedia")												        '---農業---4											
	LpId			            = Application("Video")			                                            '---影音專區--5	
	KnowledgeTankCDDatabaseId = Application("KnowledgeTankCDDatabaseId")	'---知識庫可搜尋的DBID---	
	
	'---知識庫可搜尋的DBID---	
	KnowledgeTankDatabaseId = "DB020"       '加值系統資料庫
	KnowledgeTankDatabaseId = "DB024"      'Lambda KM資料庫
	'-------------------------------------------------------------------------------r

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
	ReDim ResultArray(PageSize, 19)  
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
	'KeywordArray = Split(Keyword, " ")
	'If UBound(KeywordArray) > 0 Then
	'	For i = 0 To UBound(KeywordArray)
	'		If i = 0 Then
	'			StrQuery = KeywordArray(i)
	'		Else
	'			StrQuery = StrQuery & " * " & KeywordArray(i)
	'		End If		
	'	Next
	'Else
	'	StrQuery = Keyword
	'End If
	StrQuery = trim(Keyword)	
	StrQuery = replace(StrQuery,"   ","*")
	StrQuery = replace(StrQuery,"  ","*")
	StrQuery = replace(StrQuery," ","*")
	StrQuery = replace(StrQuery,";;;","*")
	StrQuery = replace(StrQuery,";;","*")
	StrQuery = replace(StrQuery,";","*")	
	StrQuery = replace(StrQuery,"***","*")
	StrQuery = replace(StrQuery,"**","*")	
	StrQuery = replace(StrQuery,"*"," * ")	
		
	KeywordArray = Split(StrQuery, " * ")
	
	'StrQuery = replace(StrQuery,"*"," + ")          '" * " 代表 and，" + " 代表 or，" - "  代表 not。	
	'response.Write StrQuery
	'---for query---
	'-------------------------------------------------------------------------------
	'---Data From-------------------------------------------------------------------		
	Dim SiteFlag
	SiteFlag = false	

'KnowledgeTankDatabaseId = "DB024"
'FromKnowledgeTank=1
'FromTechCD=0
'FromSiteUnit=0
'FromKnowledgeHome=0
'FromTopic=0
'FromVideo=0
'FromPedia=0

	If FromKnowledgeTank = "1" Then	'---知識庫---若有抓取知識庫的資料,只能抓取DB024的----------		
		'Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", KnowledgeTankId, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )
		'Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "databaseid", KnowledgeTankDatabaseId, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
		'Call Hyftd_Add_And( HyftdObj, HyftdConnId )				
		
	    Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "databaseid", KnowledgeTankDatabaseId, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )													
		If SiteFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )		
		SiteFlag = true			
	End If

	if FromTechCD = "1" then                   '----培肥專區    
	   Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "databaseid", "DB023", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
		If SiteFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )
		SiteFlag = true		
	end if
	
	If FromSiteUnit = "1" Then				'---站內單元---	
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", SiteUnitId, 	"", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )			
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "databaseid", "41", 	"", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )		
		Call Hyftd_Add_not( HyftdObj, HyftdConnId )	
		If SiteFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )
		SiteFlag = true		
	End If				
	If FromKnowledgeHome = "1" Then	'---知識家---		    	
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", KnowledgeHomeId, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )															
		If SiteFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )		
		SiteFlag = true
	End If	

	If FromTopic = "1" Then					'---主題館---		
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", TopicId, 	"", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )
        If SiteFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )				
		SiteFlag = true
	End If		
	If FromPedia = "1" Then					'---農業小百科---			
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "siteid", 	PediaId, 	"", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )												
		If SiteFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )		
		SiteFlag = true
	End If		
	If FromVideo = "1" Then					'---影音專區---		    		
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "databaseid", 	"41", 	"", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )														
		If SiteFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )		
		SiteFlag = true
	End If	
	
	'---Keyword---------------------------------------------------------------------
	Dim TypeFlag
	If AdvanceSearch = "1" Then		
		If Subject = "1" Then
			Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "subject", StrQuery, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )			
			TypeFlag = true
		End If
		If Keywords = "1" Then
			Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "keywords", StrQuery, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )			
			If TypeFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )			
			TypeFlag = true
		End If
		If Author = "1" Then
			Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "author", 	StrQuery, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )			
			If TypeFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )			
			TypeFlag = true
		End If
		If Description = "1" Then
			Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "description", StrQuery, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )			
			If TypeFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )			
			TypeFlag = true
		End If
		If Journal = "1" Then
			Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "journal", StrQuery, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )			
			If TypeFlag = true Then Call Hyftd_Add_Or( HyftdObj, HyftdConnId )			
			TypeFlag = true
		End If		
	Else	
		Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "associatedata", StrQuery, "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )
    End If
	
	If SiteFlag = true Then Call Hyftd_Add_And( HyftdObj, HyftdConnId )	
 



    '文件狀態
    Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "status", "Y", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )    
    Call Hyftd_Add_Query( HyftdObj, HyftdConnId, "status", "PUB", "", HyftdQueryType, HyftdPhonetic, HyftdAuthorize )    
    Call Hyftd_Add_Or( HyftdObj, HyftdConnId )	
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

'response.Write (HyftdSortName)
'response.write GetFieldIndexString(HyftdObj,HyftdConnId, HyftdSortQueryId, "actorinfoid")

    responsedString = ""

	'---Get Total Record Count------------------------------------------------------
	HyftdTotalRecordCount = Hyftd_Total_Sysid( HyftdObj, HyftdConnId, HyftdSortQueryId )	
	'-------------------------------------------------------------------------------
	'---Get Current Record Count----------------------------------------------------
	HyftdCurrentRecordCount = Hyftd_Num_Sysid( HyftdObj, HyftdConnId, HyftdSortQueryId )		
	'-------------------------------------------------------------------------------
	'---Get Record Content----------------------------------------------------------				
	'response.write HyftdCurrentRecordCount
	If HyftdCurrentRecordCount > 0 Then				
		For index = 0 To HyftdCurrentRecordCount - 1		
            										
            If ((PageNumber - 1) * PageSize + index) = HyftdTotalRecordCount Then				
	            Exit For
            End If
            '---------------------------------------------------------------------------
            '---Fetch Current Sysid-----------------------------------------------------
            HyftdDataId = Hyftd_Fetch_Sysid( HyftdObj, HyftdConnId, HyftdSortQueryId, index )				
            
	        response.write "<!--" & HyftdDataId & "<hr>-->" & vbcrlf
            
            
            TempArray = Split( HyftdDataId, "@")
            ResultArray(index, 0) = HyftdDataId
            ResultArray(index, 2) = TempArray(0)	'---siteid---
            ResultArray(index, 3) = TempArray(1)  '---reportid---icuitem
            ResultArray(index, 4) = TempArray(2)	'---databaseid---basedsd
            ResultArray(index, 5) = TempArray(3)	'---categoryid---ctunit      
                
          '去除重複的記錄
          responsedStringResult = ResultArray(index, 2) & "@" & ResultArray(index, 3)
          if 0 = 0 then
            responsedString=responsedString & responsedStringResult & ", "
            'response.Write responsedString & "<hr>"
                        '---------------------------------------------------------------------------
                        '---Get Related Number------------------------------------------------------
                        HyftdDataRelatedNumber = Hyftd_Get_Related_Number( HyftdObj )
                        'response.write HyftdDataRelatedNumber & "<hr>"
                        '---------------------------------------------------------------------------
                        '---Set Value Into ResultArray----------------------------------------------             
                        
                        ResultArray(index, 1) = HyftdDataRelatedNumber
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
	    				            " , isnull(report.document_id, 0) as document_id "  & _
	    				            " , isnull(report.xURL, '') as xURL "  & _
	    				            " FROM REPORT INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID " & _
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
	      	                ResultArray(index, 11) = kmrs("xURL")
	      	                ResultArray(index, 18) = kmrs("document_id")	      		      	    
	    	            else
	                        ResultArray(index, 0) = ""
	                        NullRecordCount = NullRecordCount+1
	                        
	                        response.Write "<!--"
	                        response.Write "REPORT_ID:" & ResultArray(index, 3) & ", "
	                        response.Write "DATA_BASE_ID:" & ResultArray(index, 4) & ", "
	                        response.Write "CATEGORY_ID:" & ResultArray(index, 5) 
	                        response.Write "-->"
	    	            End If      
	    	            kmrs.close
	    	            Set kmrs = Nothing
                  	
                  ElseIf TempArray(0) = "1" Or TempArray(0) = "2" Or TempArray(0) = "3" Or TempArray(0) = "4" Then
                  	
      	            sql = "SELECT CuDTGeneric.sTitle, CuDTGeneric.ClickCount, CONVERT(char(10), CuDTGeneric.xPostDate, 111) AS xPostDate, CtUnit.CtUnitName, CuDTGeneric.showType, " & _
      				            "ISNULL(CuDTGeneric.xURL, '') AS xURL, ISNULL(CuDTGeneric.fileDownLoad, '') AS fileDownLoad, CatTreeNode.CtNodeID, SUBSTRING(ISNULL(CuDTGeneric.xBody, ''), 0, 60) AS xBody, CatTreeNode.CtRootID " & _
      				            "FROM CuDTGeneric INNER JOIN CtUnit ON CuDTGeneric.iCTUnit = CtUnit.CtUnitID INNER JOIN CatTreeNode " & _
      				            "ON CtUnit.CtUnitID = CatTreeNode.CtUnitID WHERE (CuDTGeneric.iCUItem = '" & ResultArray(index, 3) & "') "      				            
      				'sql = sql & vbcrlf & "     and CuDTGeneric.fCTUPublic = 'Y'"
		            'sql = sql & vbcrlf & "     and mGIPcoanew.dbo.CatTreeNode.CtRootID not in  (177)"		 '177 農業知識入口網(2008電子報)
		            'sql = sql & vbcrlf & "     and mGIPcoanew.dbo.CatTreeNode.CtNodeId not in  (4783, 4784, 4785, 4786)"		 ' 最新消息
		            'sql = sql & vbcrlf & "     and  ( (iCtUnit='126' and NewtopCat !='01') or  iCtUnit != '126')"		 ' 最新消息
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
	      	                ResultArray(index, 14) = cudtrs("xBody")
	      	                ResultArray(index, 15) = cudtrs("CtRootID")	   
	      	                
	      	                select case cudtrs("CtNodeID")
	      	                case 4783, 4784, 4785, 4786
	      	                    ResultArray(index, 13) = 4782
	      	                case else	      	                
	      	                    ResultArray(index, 13) = cudtrs("CtNodeID")
	      	                end select
	      	                   
	                      else
	                            ResultArray(index, 0) = ""
	                            NullRecordCount = NullRecordCount+1

	                            'response.Write "<!--"
	                            response.Write "iCUItem:" & ResultArray(index, 3) & ", "
	                            response.Write "CtNodeID:" & ResultArray(index, 7) 
	                            'response.Write "-->"

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
          
          else
            '去除重複記錄
            ResultArray(index, 0) = ""
            NullRecordCount = NullRecordCount +1
          end if
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
	HyftdTeachCDGroupingName1 = "databaseid"
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
	HyftdQueryGroupingId = Hyftd_Query_Grouping( HyftdObj, HyftdConnId, HyftdGroupingName1, GroupingLen, HyftdGroupingType1 )'HyftdGroupingName1指siteId
	'---get how many group---
	HyftdGroupingCount1 = Hyftd_Num_Grouping( HyftdObj, HyftdConnId, HyftdQueryGroupingId )		
	'
	ReDim HyftdGrouping1(HyftdGroupingCount1, 3)	
	
	
	'Function:獲得以databaseId
	HyftdQueryTeachCDGroupingId= Hyftd_Query_Grouping(HyftdObj, HyftdConnId, HyftdTeachCDGroupingName1, GroupingLen, HyftdGroupingType1 )'HyftdTeachCDGroupingName1 指databaseId
	HyftdTeachCDGroupingCount1 = Hyftd_Num_Grouping( HyftdObj, HyftdConnId, HyftdQueryTeachCDGroupingId )	
	ReDim HyftdTeachCDGrouping1(HyftdTeachCDGroupingCount1, 3)	
  
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
			pathGroupName(1) = "主題館"
			pathGroupData(1) = "2"				
		End If
		If FromKnowledgeHome = "1" Then
			pathGroupName(1) = "知識家"
			pathGroupData(1) = "3"				
		End If		
		If FromPedia = "1" Then
			pathGroupName(1) = "農業小百科"
			pathGroupData(1) = "4"				
		End If		
		If FromVideo = "1" Then
			pathGroupName(1) = "影音專區"
			pathGroupData(1) = "5"				
		End If	
		If FromTechCD = "1" Then
			pathGroupName(1) = "技術光碟"
			pathGroupData(1) = "6"				
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
		'
		 If HyftdTeachCDGroupingCount1 > 0 Then'增加databaseId的查詢條件
			For index = 0 To HyftdTeachCDGroupingCount1 - 1
				 HyftdTeachCdGrouping1(index, 1) = Hyftd_Fetch_Grouping( HyftdObj, HyftdConnId, HyftdQueryTeachCDGroupingId, index )
				 HyftdTeachCdGrouping1(index, 0) = Hyftd_Get_Grouping_Name( HyftdObj )							
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

<script language="vbs">
	Dim CanTarget
	Dim followCanTarget

	Sub popCalendar(dateName,followName)        
 		CanTarget = dateName
 		followCanTarget = followName
		xdate = document.all(CanTarget).value
		if not isDate(xdate) then	xdate = date()
		document.all.calendar1.setDate xdate
	
 		If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 		Else        
    	ex = window.event.clientX + 100
			ey = document.body.scrolltop + window.event.clientY + 10
      'ey = window.event.srcElement.parentElement.offsetTop
      if ex > 520 then ex = 520
       document.all.calendar1.style.pixelleft = ex - 80
       document.all.calendar1.style.pixeltop = ey
       document.all.calendar1.style.visibility = ""
 		End If              
	End Sub 

	Sub calendar1_onscriptletevent(n,o) 
    document.all("CalendarTarget").value = n         
    document.all.calendar1.style.visibility = "hidden"    
    if n <> "Cancle" then
      document.all(CanTarget).value = document.all.CalendarTarget.value        
      if followCanTarget <> "" then
        document.all(followCanTarget).value = document.all.CalendarTarget.value        	
      end if
		end if
	End Sub 	
</script>

<script language="javascript">
	
	function CheckSubmit() 
	{
		var bolFlag = false;
		var depth = <%=Depth%>;
		if( document.iForm.QueryString.value == "" )  {
			document.iForm.Keyword.value = "<%=Keyword%>";
		}
		else {
			if( document.iForm.Range.value == "0" ) {
				document.iForm.Keyword.value = "<%=Keyword%>;" + document.iForm.QueryString.value;
			}
			else {
				document.iForm.Keyword.value = document.iForm.QueryString.value;
			}			
		}		
		document.iForm.FromSiteUnit.disabled 		= false;
		document.iForm.FromKnowledgeTank.disabled 	= false;
		document.iForm.FromKnowledgeHome.disabled 	= false;
		document.iForm.FromTopic.disabled 			= false;
		document.iForm.FromPedia.disabled 			= false;
		document.iForm.FromVideo.disabled 			= false;
		document.iForm.FromTechCD.disabled 			= false;
		
		if( document.iForm.FromSiteUnit.checked		||
			document.iForm.FromKnowledgeTank.checked|| 
			document.iForm.FromKnowledgeHome.checked|| 
			document.iForm.FromTopic.checked 		||
			document.iForm.FromPedia.checked 		||
			document.iForm.FromVideo.checked 		||
			document.iForm.FromTechCD.checked 
			) {
				bolFlag = true;
		}
		if( !bolFlag ) {
			document.iForm.FromSiteUnit.checked		 = true; 
			document.iForm.FromKnowledgeTank.checked = true;
			document.iForm.FromKnowledgeHome.checked = true;
			document.iForm.FromTopic.checked 		 = true;
			document.iForm.FromPedia.checked 		 = true;
			document.iForm.FromVideo.checked 		 = true;
			document.iForm.FromTechCD.checked 		 = true;
		}
		if( document.iForm.Range.selectedIndex == 1 ) {
			document.iForm.CategoryId.value = "";
			document.iForm.Depth.value = "0";							
		}			
		document.iForm.RelatedNumber.value 		= "";
		document.iForm.MaxRelated.value 		= "";
		document.iForm.PageNumber.selectedIndex = 0;
		document.iForm.PageSize.selectedIndex 	= 0;
		document.iForm.Order.selectedIndex 		= 0;
		document.iForm.Sort.selectedIndex 		= 0;					
		document.iForm.submit();
	}
	
	function CheckSubmitForTime( value ) 
	{		
		document.iForm.Keyword.value = "<%=Keyword%>";	
		var bolFlag = false;
		document.iForm.FromSiteUnit.disabled 		= false;
		document.iForm.FromKnowledgeTank.disabled 	= false;
		document.iForm.FromKnowledgeHome.disabled 	= false;
		document.iForm.FromTopic.disabled 			= false;
		document.iForm.FromPedia.disabled 			= false;
		document.iForm.FromVideo.disabled 			= false;
		document.iForm.FromTechCD.disabled			= false;
		
		if( document.iForm.FromSiteUnit.checked 	            || 
			document.iForm.FromKnowledgeTank.checked    || 
			document.iForm.FromKnowledgeHome.checked   || 
			document.iForm.FromTopic.checked 		            ||
			document.iForm.FromPedia.checked 		            ||
			document.iForm.FromVideo.checked 		            ||
			document.iForm.FromTechCD.checked		
			) {
				bolFlag = true;
		}
		if( !bolFlag ) {
			document.iForm.FromSiteUnit.checked 		= true; 
			document.iForm.FromKnowledgeTank.checked 	= true;
			document.iForm.FromKnowledgeHome.checked 	= true;
			document.iForm.FromTopic.checked 			= true;
			document.iForm.FromPedia.checked 			= true;
			document.iForm.FromVideo.checked 			= true;
			document.iForm.FromTechCD.checked 			= true;
		}
		if( value != "" ) {
			var year = value.substring(0, 4);
			var month = value.substring(4, 6);
			var daysinmonth = daysInMonth(month, year);
			document.iForm.StartDate.value = year + "/" + month + "/01";
			document.iForm.EndDate.value = year + "/" + month + "/" + daysinmonth;
		}			
		document.iForm.RelatedNumber.value 		= "";		
		document.iForm.MaxRelated.value 		= "";
		document.iForm.PageSize.selectedIndex 	= 0;
		document.iForm.PageNumber.selectedIndex = 0;
		document.iForm.submit();
	}
	
	function CheckSubmitForCategory( catid ) 
	{	
		var depth = <%=Depth%>;		
		var newdepth = <%=CInt(Depth) + 1%>;	
		if( newdepth > 4 )	{
			newdepth = 4;
		}
		
		
		document.iForm.Depth.value = newdepth;		
		document.iForm.Keyword.value = "<%=Keyword%>";	
		var bolFlag = false;
		if( depth == 0 ) {
			document.iForm.FromSiteUnit.disabled 		= false;
			document.iForm.FromKnowledgeTank.disabled 	= false;
			document.iForm.FromKnowledgeHome.disabled 	= false;
			document.iForm.FromTopic.disabled 			= false;
			document.iForm.FromPedia.disabled 			= false;
			document.iForm.FromVideo.disabled 			= false;
			document.iForm.FromTechCD.disabled 			= false;
			
			document.iForm.FromSiteUnit.checked 		= false;
			document.iForm.FromKnowledgeTank.checked 	= false;
			document.iForm.FromKnowledgeHome.checked 	= false;
			document.iForm.FromTopic.checked 			= false;
			document.iForm.FromPedia.checked 			= false;
			document.iForm.FromVideo.checked 			= false;
			document.iForm.FromTechCD.checked 			= false;
			if( catid == "1" ) {
					document.iForm.FromSiteUnit.checked = true;
					bolFlag = true;
			}							
			if( catid == "0" ) {
					document.iForm.FromKnowledgeTank.checked = true;
					bolFlag = true;
			}				
			if( catid == "3" ) {
					document.iForm.FromKnowledgeHome.checked = true;
					bolFlag = true;
			}							
			if( catid == "2" ) {
					document.iForm.FromTopic.checked = true;
					bolFlag = true;
			}	
			if( catid == "4" ) {
					document.iForm.FromPedia.checked = true;
					bolFlag = true;
			}	
			if( catid == "5" ) {
					window.location="kp.asp?xdURL=Search/SearchResultList.asp&mp=1&search.x=35&search.y=14&Keyword=<%=server.URLEncode(keyword)%>&FromVideo=1&memID=&gstyle=";
					return;
			}	
			if( catid == "6" ) {
					window.location="kp.asp?xdURL=Search/SearchResultList.asp&mp=1&search.x=35&search.y=14&Keyword=<%=server.URLEncode(keyword)%>&FromTechCD=1&memID=&gstyle=";
					return;
			}				
		}
		else {
			document.iForm.FromSiteUnit.disabled = false;
			document.iForm.FromKnowledgeTank.disabled = false;
			document.iForm.FromKnowledgeHome.disabled = false;
			document.iForm.FromTopic.disabled = false;
			if( document.iForm.FromSiteUnit.checked || document.iForm.FromKnowledgeTank.checked 
					|| document.iForm.FromKnowledgeHome.checked || document.iForm.FromTopic.checked ) {
					bolFlag = true;
			}			
		}			
		if( !bolFlag )		{
			document.iForm.FromSiteUnit.checked 		= true;
			document.iForm.FromKnowledgeTank.checked 	= true;
			document.iForm.FromKnowledgeHome.checked 	= true;
			document.iForm.FromTopic.checked 			= true;
			document.iForm.FromPedia.checked 			= true;
			document.iForm.FromVideo.checked 			= true;
			document.iForm.FromTechCD.checked 			= true;
		}
		//alert(  "bolFlag:"+bolFlag+" "+catid);
		document.iForm.CategoryId.value = catid;
		document.iForm.RelatedNumber.value = "";
		document.iForm.MaxRelated.value = "";
		document.iForm.PageSize.selectedIndex = 0;
		document.iForm.PageNumber.selectedIndex = 0;
		document.iForm.submit();
	}
	
	function RangeOnChange() 	
	{
		document.iForm.Keyword.value = "<%=Keyword%>";		
		if( document.iForm.Range.selectedIndex == 1 ) {	
			document.iForm.FromSiteUnit.disabled 		= false;
			document.iForm.FromKnowledgeTank.disabled 	= false;
			document.iForm.FromKnowledgeHome.disabled 	= false;
			document.iForm.FromTopic.disabled 			= false;
			document.iForm.FromPedia.disabled 			= false;
			document.iForm.FromVideo.disabled 			= false;
			document.iForm.FromTechCD.disabled 			= false;			
		}
		else {
			document.iForm.FromSiteUnit.disabled 		= true;
			document.iForm.FromKnowledgeTank.disabled 	= true;
			document.iForm.FromKnowledgeHome.disabled 	= true;
			document.iForm.FromTopic.disabled 			= true;	
			document.iForm.FromPedia.disabled 			= true;
			document.iForm.FromVideo.disabled 			= true;
			document.iForm.FromTechCD.disabled 			= true;
		}
	}
	
	function MakePageNumberSame() {
	  document.iForm.PageNumber.selectedIndex = document.getElementById("PageNumber_down").value - 1 ;
	  PageNumberOnChange() ;
	} 
	
	
	function PageNumberOnChange() {
		
		//document.iForm.PageSize.selectedIndex = 0;
		document.iForm.Keyword.value = "<%=Keyword%>";		
		var bolFlag = false;
		document.iForm.FromSiteUnit.disabled 		= false;
		document.iForm.FromKnowledgeTank.disabled 	= false;
		document.iForm.FromKnowledgeHome.disabled	= false;
		document.iForm.FromTopic.disabled 			= false;
		document.iForm.FromPedia.disabled 			= false;
		document.iForm.FromVideo.disabled 			= false;
		document.iForm.FromTechCD.disabled 			= false;
		
		if( document.iForm.FromSiteUnit.checked 		||
			document.iForm.FromKnowledgeTank.checked 	||
			document.iForm.FromKnowledgeHome.checked 	||
			document.iForm.FromTopic.checked			||
			document.iForm.FromPedia.checked 			||
			document.iForm.FromVideo.checked 			||
			document.iForm.FromTechCD.checked 	
			) {
				bolFlag = true;			
		}
		if( !bolFlag ) {
			document.iForm.FromSiteUnit.checked 		= true; 
			document.iForm.FromKnowledgeTank.checked 	= true;
			document.iForm.FromKnowledgeHome.checked 	= true;
			document.iForm.FromTopic.checked 			= true;
			document.iForm.FromPedia.checked 			= true;
			document.iForm.FromVideo.checked 			= true;
			document.iForm.FromTechCD.checked 			= true;
		}
		document.iForm.RelatedNumber.value = "1";		
		document.iForm.submit();
	}
	
	function MakePageSizeSame() {
	  if ( document.getElementById("PageSize_down").value == 10 ) {
	    document.iForm.PageSize.selectedIndex = 0 ;
	  }
	  else if ( document.getElementById("PageSize_down").value == 30 ) {
	    document.iForm.PageSize.selectedIndex = 1 ;
	  }
	  else if ( document.getElementById("PageSize_down").value == 50 ) {
	    document.iForm.PageSize.selectedIndex = 2 ;
	  }
	  PageSizeOnChange() ;
	}
	
	function PageSizeOnChange() 
	{
		document.iForm.Keyword.value = "<%=Keyword%>";		
		document.iForm.PageNumber.selectedIndex = 0;
		var bolFlag = false;
		document.iForm.FromSiteUnit.disabled		= false;
		document.iForm.FromKnowledgeTank.disabled   = false;
		document.iForm.FromKnowledgeHome.disabled 	= false;
		document.iForm.FromTopic.disabled 			= false;
		document.iForm.FromPedia.disabled 			= false;
		document.iForm.FromVideo.disabled 			= false;
		document.iForm.FromTechCD.disabled 			= false;
		
		if( document.iForm.FromSiteUnit.checked 		|| 
			document.iForm.FromKnowledgeTank.checked 	|| 
			document.iForm.FromKnowledgeHome.checked 	|| 
			document.iForm.FromTopic.checked 			||
			document.iForm.FromPedia.checked 			||
			document.iForm.FromVideo.checked 			||
			document.iForm.FromTechCD.checked 				) {
				bolFlag = true;
		        }
		if( !bolFlag ) {
			document.iForm.FromSiteUnit.checked 		= true; 
			document.iForm.FromKnowledgeTank.checked 	= true;
			document.iForm.FromKnowledgeHome.checked 	= true;
			document.iForm.FromTopic.checked 			= true;
			document.iForm.FromPedia.checked 			= true;
			document.iForm.FromVideo.checked 			= true;
			document.iForm.FromTechCD.checked 			= true;
		}
		
	
		document.iForm.RelatedNumber.value = "1";				
		document.iForm.submit();
	}
	
	
	function PreviousPageOnClick( value ) 
	{
		document.iForm.Keyword.value = "<%=Keyword%>";		
		document.iForm.PageNumber.selectedIndex = value - 1;		
		var bolFlag = false;
		document.iForm.FromSiteUnit.disabled 		        = false;
		document.iForm.FromKnowledgeTank.disabled 	= false;
		document.iForm.FromKnowledgeHome.disabled 	= false;
		document.iForm.FromTopic.disabled 			        = false;
		document.iForm.FromPedia.disabled 			        = false;
		document.iForm.FromVideo.disabled 			        = false;
		document.iForm.FromTechCD.disabled 			    = false;
		
		if( document.iForm.FromSiteUnit.checked 		        || 
			document.iForm.FromKnowledgeTank.checked 	|| 
			document.iForm.FromKnowledgeHome.checked 	||
			document.iForm.FromTopic.checked 			        ||
			document.iForm.FromPedia.checked 			        ||
			document.iForm.FromVideo.checked 			        ||
			document.iForm.FromTechCD.checked 				) {
				bolFlag = true;
		}
		if( !bolFlag ) {
			document.iForm.FromSiteUnit.checked 		        = true; 
			document.iForm.FromKnowledgeTank.checked 	= true;
			document.iForm.FromKnowledgeHome.checked 	= true;
			document.iForm.FromTopic.checked 			        = true;
			document.iForm.FromPedia.checked 			        = true;
			document.iForm.FromVideo.checked 			        = true;
			document.iForm.FromTechCD.checked 			    = true;
		}
		document.iForm.RelatedNumber.value = "1";				
		document.iForm.submit();
	}
	
	function NextPageOnClick( value ) 
	{
		document.iForm.Keyword.value = "<%=Keyword%>";		
		document.iForm.PageNumber.selectedIndex = value - 1;	
		var bolFlag = false;
		document.iForm.FromSiteUnit.disabled 	        	= false;
		document.iForm.FromKnowledgeTank.disabled 	= false;
		document.iForm.FromKnowledgeHome.disabled 	= false;
		document.iForm.FromTopic.disabled 			        = false;
		document.iForm.FromPedia.disabled 			        = false;
		document.iForm.FromVideo.disabled 			        = false;
		document.iForm.FromTechCD.disabled 			    = false;
		
		if( document.iForm.FromSiteUnit.checked 	            ||
			document.iForm.FromKnowledgeTank.checked    || 
			document.iForm.FromKnowledgeHome.checked   || 
			document.iForm.FromTopic.checked 		            ||
			document.iForm.FromPedia.checked 		            ||
			document.iForm.FromVideo.checked 		            ||
			document.iForm.FromTechCD.checked 			 ) {
				bolFlag = true;
		}
		if( !bolFlag ) {
			document.iForm.FromSiteUnit.checked 		        = true; 
			document.iForm.FromKnowledgeTank.checked 	= true;
			document.iForm.FromKnowledgeHome.checked 	= true;
			document.iForm.FromTopic.checked 			        = true;
			document.iForm.FromPedia.checked 			        = true;
			document.iForm.FromVideo.checked 			        = true;
			document.iForm.FromTechCD.checked 			    = true;
		}
		document.iForm.RelatedNumber.value = "1";				
		document.iForm.submit();
	}
	
	function OrderOnChange() 
	{		
		document.iForm.Keyword.value = "<%=Keyword%>";		
		var bolFlag = false;
		document.iForm.FromSiteUnit.disabled 		        = false;
		document.iForm.FromKnowledgeTank.disabled 	= false;
		document.iForm.FromKnowledgeHome.disabled 	= false;
		document.iForm.FromTopic.disabled 			        = false;
		document.iForm.FromPedia.disabled 			        = false;
		document.iForm.FromVideo.disabled 			        = false;
		document.iForm.FromTechCD.disabled 			    = false;
		
		if( document.iForm.FromSiteUnit.checked 	            ||
			document.iForm.FromKnowledgeTank.checked    || 
			document.iForm.FromKnowledgeHome.checked   ||
			document.iForm.FromTopic.checked 		            ||
			document.iForm.FromPedia.checked 		            ||
			document.iForm.FromVideo.checked 		            ||
			document.iForm.FromTechCD.checked 				) {
				bolFlag = true;
		}
		if( !bolFlag ) {
			document.iForm.FromSiteUnit.checked 		        = true; 
			document.iForm.FromKnowledgeTank.checked 	= true;
			document.iForm.FromKnowledgeHome.checked 	= true;
			document.iForm.FromTopic.checked 			        = true;
			document.iForm.FromPedia.checked 			        = true;
			document.iForm.FromVideo.checked 			        = true;
			document.iForm.FromTechCD.checked 			= true;
		}
		document.iForm.RelatedNumber.value = "";				
		document.iForm.MaxRelated.value = "";
		document.iForm.PageSize.selectedIndex = 0;
		document.iForm.PageNumber.selectedIndex = 0;	
		document.iForm.submit();
	}
		
	function SortOnChange() 
	{	
		document.iForm.Keyword.value = "<%=Keyword%>";		
		var bolFlag = false;
		document.iForm.FromSiteUnit.disabled 		        = false;
		document.iForm.FromKnowledgeTank.disabled 	= false;
		document.iForm.FromKnowledgeHome.disabled 	= false;
		document.iForm.FromTopic.disabled 			        = false;
		document.iForm.FromPedia.disabled 			        = false
		document.iForm.FromVideo.disabled 			        = false;
		document.iForm.FromTechCD.disabled 		    	= false;
		
		if( document.iForm.FromSiteUnit.checked 	|| 
			document.iForm.FromKnowledgeTank.checked||
			document.iForm.FromKnowledgeHome.checked|| 
			document.iForm.FromTopic.checked 		||
			document.iForm.FromPedia.checked 		||
			document.iForm.FromVideo.checked 		||
			document.iForm.FromTechCD.checked 
			) {
				bolFlag = true;
		}
		if( !bolFlag ) {
			document.iForm.FromSiteUnit.checked 		= true; 
			document.iForm.FromKnowledgeTank.checked 	= true;
			document.iForm.FromKnowledgeHome.checked 	= true;
			document.iForm.FromTopic.checked 			= true;
			document.iForm.FromPedia.checked 			= true;
			document.iForm.FromVideo.checked 			= true;
			document.iForm.FromTechCD.checked 			= true;
		}
		document.iForm.RelatedNumber.value = "";		
		document.iForm.MaxRelated.value = "";		
		document.iForm.PageSize.selectedIndex = 0;
		document.iForm.PageNumber.selectedIndex = 0;		
		document.iForm.submit();
	}	
	
	function CategoryPathOnClick( depth, catid ) 
	{
		document.iForm.Depth.value = depth;		
		document.iForm.Keyword.value = "<%=Keyword%>";
		var bolFlag = false;
		if( depth == 0 || depth == 1 ) {
	
				document.iForm.FromSiteUnit.disabled 		= false;
				document.iForm.FromKnowledgeTank.disabled 	= false;
				document.iForm.FromKnowledgeHome.disabled 	= false;
				document.iForm.FromTopic.disabled 			= false;
				document.iForm.FromPedia.disabled 			= false;
				document.iForm.FromVideo.disabled 			= false;
				document.iForm.FromTechCD.disabled			= false;
				
				document.iForm.ActorInfoId.value = "";
				if( document.iForm.FromSiteUnit.checked 	||
					document.iForm.FromKnowledgeTank.checked|| 
					document.iForm.FromKnowledgeHome.checked|| 
					document.iForm.FromTopic.checked  		||
					document.iForm.FromPedia.checked		||
					document.iForm.FromVideo.checked 		||
					document.iForm.FromTechCD.checked
					) {
					bolFlag = true;
				}
				if( !document.iForm.FromSiteUnit.checked ) {
					document.iForm.FromSiteUnit.checked = false;
				}
				if( !document.iForm.FromKnowledgeTank.checked ) {
					document.iForm.FromKnowledgeTank.checked = false;
				}
				if( !document.iForm.FromKnowledgeHome.checked ) {
					document.iForm.FromKnowledgeHome.checked = false;
				}
				if( !document.iForm.FromTopic.checked ) {
					document.iForm.FromTopic.checked = false;
				}									
				if( !document.iForm.FromPedia.checked ) {
					document.iForm.FromPedia.checked = false;
				}
				if( !document.iForm.FromVideo.checked ) {
					document.iForm.FromVideo.checked = false;
				}
				if( !document.iForm.FromTechCD.checked ) {
					document.iForm.FromTechCD.checked = false;
				}				
			document.iForm.CategoryId.value = "";
		}
		else {
			document.iForm.FromSiteUnit.disabled 		= false;
			document.iForm.FromKnowledgeTank.disabled 	= false;
			document.iForm.FromKnowledgeHome.disabled 	= false;
			document.iForm.FromTopic.disabled 			= false;
			document.iForm.FromPedia.disabled 			= false;
			document.iForm.FromVideo.disabled 			= false;
			document.iForm.FromTechCD.disabled 			= false;
			
			if( document.iForm.FromSiteUnit.checked 		|| 
				document.iForm.FromKnowledgeTank.checked 	|| 
				document.iForm.FromKnowledgeHome.checked 	|| 
				document.iForm.FromTopic.checked 			||
				document.iForm.FromPedia.checked 			||
				document.iForm.FromVideo.checked 			||
				document.iForm.FromTechCD.checked 
				) {
				bolFlag = true;
			}		
			document.iForm.CategoryId.value = catid;
		}
		if( !bolFlag ) {
			document.iForm.FromSiteUnit.checked 		= true; 
			document.iForm.FromKnowledgeTank.checked 	= true;
			document.iForm.FromKnowledgeHome.checked 	= true;
			document.iForm.FromTopic.checked 			= true;
			document.iForm.FromPedia.checked 			= true;
			document.iForm.FromVideo.checked 			= true;
			document.iForm.FromTechCD.checked			= true;
		}
		document.iForm.RelatedNumber.value = "";
		document.iForm.MaxRelated.value = "";
		document.iForm.PageSize.selectedIndex = 0;
		document.iForm.PageNumber.selectedIndex = 0;
		document.iForm.submit();
	}	
	
	function daysInMonth(month,	year) 
	{
		var m = [31,28,31,30,31,30,31,31,30,31,30,31];
		if (month != 2) 
			return m[month - 1];
		if (year%4 != 0) 
			return m[1];
		if (year%100 == 0 && year%400 != 0) 
			return m[1];
		return m[1] + 1;
	} 
</script>
<form name="iForm" method="post" action="kp.asp?xdURL=Search/SearchResultList.asp&mp=1">
<div class="searchResult">
<table class="layout">
<tr>
<td class="left">
	<table class="layouttable">
	
	
	<input name="CategoryId"  	type="hidden" value="<%=CategoryId%>" >		
	<input name="Phonetic" 	    type="hidden" value="<%=HyftdPhonetic%>" >
	<input name="Authorize" 	    type="hidden" value="<%=HyftdAuthorize%>" >
	<input name="queryStyle" 	    type="hidden" value="<%=Auth%>" >		
	<input name="QueryType" 	    type="hidden" value="0" >		
	<input name="IsThisPage" 	    type="hidden" value="1" >
	<input name="Depth" 		    type="hidden" value="<%=Depth%>" >
	<input name="Keyword" 		type="hidden" value="<%=QueryString%>" >
	<input name="RelatedNumber" type="hidden" value="<%=RelatedNumber%>" >
	<input name="MaxRelated" 	type="hidden" value="<%=MaxRelated%>" >
	<input name="AdvanceSearch" type="hidden" value="<%=AdvanceSearch%>" >
	<input name="Subject" 		    type="hidden" value="<%=Subject%>" >
	<input name="Keywords" 		type="hidden" value="<%=Keywords%>" >
	<input name="Author" 		    type="hidden" value="<%=Author%>" >
	<input name="Description" 	type="hidden" value="<%=Description%>" >
	<input name="Journal" 		    type="hidden" value="<%=Journal%>" >
	<input name="ActorInfoId" 	type="hidden" value="<%=ActorInfoId%>" >
	<input name="debug" 		    type="hidden" value="<%=debug%>" >
		
	<tr><th scope="col">結果內再查詢</th></tr>
	<tr>
		<td>		
			<div><label>關鍵字：<input name="QueryString" value="" type="text" size="8"></label></div><hr/>
			<div>資料發布日期：<br />
				<label>
					自<input name="StartDate" type="text" value="<%=StartDate%>" size="8" />
					<img src="/xslgip/style1/images3/calendar.gif" alt="設定開始日期" width="16" height="16" border="0" align="absmiddle" onclick="VBS: popCalendar 'StartDate',''">
					起<br />
					至<input name="EndDate" type="text" value="<%=EndDate%>" size="8" />
					<img src="/xslgip/style1/images3/calendar.gif" alt="設定結束日期" width="16" height="16" border="0" align="absmiddle" onclick="VBS: popCalendar 'EndDate',''">
					止</label>
					<input type="hidden" name="CalendarTarget">
	`				<OBJECT id="calendar1" style="LEFT: 230px; VISIBILITY: hidden; TOP: 0px; POSITION: ABSOLUTE;" type="text/x-scriptlet" height="160" width="245" data="../inc/calendar.asp"></OBJECT>
			</div><hr/>
			<div>資料來源：<br/>
				<label><input type="checkbox" name="FromSiteUnit" value="1" 	        	<% If FromSiteUnit = "1" Then  %> checked 		        <% End If %> <% If Range = "0" Then %> disabled="true" <% End If %> />站內單元</label><br/>
				<label><input type="checkbox" name="FromKnowledgeTank" value="1" 	<% If FromKnowledgeTank = "1" Then %> checked 	<% End If %> <% If Range = "0" Then %> disabled="true" <% End If %> />知識庫</label><br/>
				<label><input type="checkbox" name="FromKnowledgeHome" value="1" 	<% If FromKnowledgeHome = "1" Then %> checked 	<% End If %> <% If Range = "0" Then %> disabled="true" <% End If %> />知識家</label><br/>
				<label><input type="checkbox" name="FromTopic" value="1" 		        	<% If FromTopic = "1" Then %> checked 			        <% End If %> <% If Range = "0" Then %> disabled="true" <% End If %> />主題館</label>
				<label><input type="checkbox" name="FromPedia" value="1" 			        <% If FromPedia = "1" Then %> checked 			        <% End If %> <% If Range = "0" Then %> disabled="true" <% End If %> />農業小百科</label><br/>
				<label><input type="checkbox" name="FromVideo" value="1" 			        <% If FromVideo = "1" Then %> checked 			        <% End If %> <% If Range = "0" Then %> disabled="true" <% End If %> />影音專區</label><br/>
				<label><input type="checkbox" name="FromTechCD" value="1" 			    <% If FromTechCD= "1" Then %> checked 			    <% End If %> <% If Range = "0" Then %> disabled="true" <% End If %> />技術光碟 </label><br/>
		
			</div><hr/>
			<select name="Range" onchange="RangeOnChange()">
				<option value="0" <% If Range = "0" Then %> selected="selected" <% End If %> >搜尋此結果範圍下</option>
				<option value="1" <% If Range = "1" Then %> selected="selected" <% End If %> >搜尋全部範圍</option>
			</select><hr/>
			<input name="search" type="image" src="/xslgip/style1/images3/searchBtn_1.gif" alt="search" class="btn" onclick="CheckSubmit()" />
			<!--input name="search" type="image" src="/xslgip/style1/images3/SearchBtn_2.gif" alt="search" class="btn" /-->			
		</td>
	</tr>
	</table>

<!--後分類-->
	<table class="layouttable">
	<tr><th scope="col">分類中符合的項目</th></tr>
	<tr>
		<td>
			<div class="catapath">							
			分類路徑：
			<%											
				if CInt(Depth) >= 0 Then
					response.write "<a href=""javascript:CategoryPathOnClick('0', '" & pathGroupData(0) & "')"">" & pathGroupName(0) & "</a>"
				end if
				if CInt(Depth) >= 1 Then
					response.write " >> " & pathGroupName(1)
					'response.write " >> " & "<a href=""javascript:CategoryPathOnClick('1', '" & pathGroupData(1) & "')"">" & pathGroupName(1) & "</a>" 
				end if
				if CInt(Depth) >= 2 Then
					response.write " >> " & pathGroupName(2)
					'response.write " >> " & "<a href=""javascript:CategoryPathOnClick('2', '" & pathGroupData(2) & "')"">" & pathGroupName(2) & "</a>"
				end if
				if CInt(Depth) >= 3 Then
					response.write " >> " & pathGroupName(3)
					'response.write " >> " & "<a href=""javascript:CategoryPathOnClick('3', '" & pathGroupData(3) & "')"">" & pathGroupName(3) & "</a>"
				end if
				if CInt(Depth) >= 4 Then
					response.write " >> " & pathGroupName(4)
					'response.write " >> " & "<a href=""javascript:CategoryPathOnClick('4', '" & pathGroupData(4) & "')"">" & pathGroupName(4) & "</a>" 
				end if
			%>
			</div>
			<ul>
			<%
				Dim Str0, Str1, Str2, Str3, Str4,Str5,Str6,Str7				
							
				If Depth = "0" Then
				    Str0 = "<li><img src=""/xslgip/style1/images3/tree_closed.gif"" width=""20"" height=""11"" border=""0"" align=""absmiddle"">"
				    Str0 = Str0 & "<a href=""javascript:CheckSubmitForCategory('{2}');"">{0}({1})</a></li>"
				    
					Str1 = Replace( Replace( Replace( Str0, "{0}", "知識庫"  				), "{1}", "0" ), "{2}", "0")
					Str2 = Replace( Replace( Replace( Str0, "{0}", "站內單元"				), "{1}", "0" ), "{2}", "1")						
					Str3 = Replace( Replace( Replace( Str0, "{0}", "主題館"  				), "{1}", "0" ), "{2}", "2")						
					Str4 = Replace( Replace( Replace( Str0, "{0}", "知識家"  				), "{1}", "0" ), "{2}", "3")							
					Str5 = Replace( Replace( Replace( Str0, "{0}", "農業小百科" 			), "{1}", "0" ), "{2}", "4")	
					Str6 = Replace( Replace( Replace( Str0, "{0}", "影音專區"  			), "{1}", "0" ), "{2}", "5")	
					Str7 = Replace( Replace( Replace( Str0, "{0}", "技術光碟"  	), "{1}", "0" ), "{2}", "6")	
				   Dim databaseIdCDCount
				   databaseIdCDCount=0
				   
					'獲取databaseid分成培肥的資料及知識庫(分類夾的名稱及數量)
					For index = 0 To HyftdTeachCDGroupingCount1 - 1
						If HyftdTeachCDGrouping1(index, 0) = KnowledgeTankDatabaseId Then
							Str1 = Replace( Replace( Replace( Str0, "{0}", "知識庫"  		            		), "{1}", HyftdTeachCDGrouping1(index, 1) ), "{2}", "0")							
						End If
					
						If HyftdTeachCDGrouping1(index, 0) = "DB023" Then
						    
							Str7 = Replace( Replace( Replace( Str0, "{0}", "技術光碟"  	), "{1}", HyftdTeachCDGrouping1(index, 1) ), "{2}", "6")	
						End If
						
						If HyftdTeachCDGrouping1(index, 0) = "41" Then
						   databaseIdCDCount=HyftdTeachCDGrouping1(index, 1)
							Str6 = Replace( Replace( Replace( Str0, "{0}", "影音專區"  			            ), "{1}", HyftdTeachCDGrouping1(index, 1) ), "{2}", "5")	
						End If
					Next
				
					For index = 0 To HyftdGroupingCount1 - 1
						
						If HyftdGrouping1(index, 0) = "1" Then
							Str2 = Replace( Replace( Replace( Str0, "{0}", "站內單元"				), "{1}", HyftdGrouping1(index, 1)-databaseIdCDCount ), "{2}", "1")							
						ElseIf HyftdGrouping1(index, 0) = "2" Then
							Str3 = Replace( Replace( Replace( Str0, "{0}", "主題館"  				), "{1}", HyftdGrouping1(index, 1) ), "{2}", "2")	
						ElseIf HyftdGrouping1(index, 0) = "3" Then
							Str4 = Replace( Replace( Replace( Str0, "{0}", "知識家"  				), "{1}", HyftdGrouping1(index, 1) ), "{2}", "3")	
						ElseIf HyftdGrouping1(index, 0) = "4" Then
							Str5 = Replace( Replace( Replace( Str0, "{0}", "農業小百科"  			), "{1}", HyftdGrouping1(index, 1) ), "{2}", "4")	
						End If						
					Next
					
				
					Response.Write Str2 & Str1 & Str4 & Str3& Str5& Str6& Str7
				Else
				    Str0 = "<li><img src=""/xslgip/style1/images3/tree_closed.gif"" width=""20"" height=""11"" border=""0"" align=""absmiddle"">"
				    Str0 = Str0 & "{0}({1})</li>"
				    				
					Stra = "<li><img src=""/xslgip/style1/images3/tree_closed.gif"" width=""20"" height=""11"" border=""0"" align=""absmiddle"">{0}({1})</li>"
					For index = 0 To HyftdGroupingCount1 - 1							
						'TempArray = Split(HyftdGrouping1(index, 0), " ")
						If Depth = "1" Then
							'If FromSiteUnit = "1" Then
							'	Str1 = Str1 & Replace( Replace( Stra, "{0}", HyftdGrouping1(index, 2) ), "{1}", HyftdGrouping1(index, 1) )	
							'Else
							  If FromKnowledgeTank="1"  Then
							    Str1=""
							  Else
							  	Str1 = Str1 & Replace( Replace( Replace(Str0, "{0}", HyftdGrouping1(index, 2) ), "{1}", HyftdGrouping1(index, 1) ), "{2}", Replace(HyftdGrouping1(index, 0), " ", "@")) 'TempArray(1) )	
							  End If
							
						ElseIf Depth = "2" Then
							If FromKnowledgeHome = "1" Or FromSiteUnit = "1" Then
								Str1 = Str1 & Replace( Replace( Stra, "{0}", HyftdGrouping1(index, 2) ), "{1}", HyftdGrouping1(index, 1) )	
							Else
								Str1 = Str1 & Replace( Replace( Replace(Str0, "{0}", HyftdGrouping1(index, 2) ), "{1}", HyftdGrouping1(index, 1) ), "{2}", Replace(HyftdGrouping1(index, 0), " ", "@")) 'TempArray(1) )										
							End If
						ElseIf Depth = "3" Then
							If FromTopic = "1" Then
								Str1 = Str1 & Replace( Replace( Stra, "{0}", HyftdGrouping1(index, 2) ), "{1}", HyftdGrouping1(index, 1) )	
							Else
								Str1 = Str1 & Replace( Replace( Replace(Str0, "{0}", HyftdGrouping1(index, 2) ), "{1}", HyftdGrouping1(index, 1) ), "{2}", Replace(HyftdGrouping1(index, 0), " ", "@")) 'TempArray(1) )										
							End If
						ElseIf Depth = "4" Then
							If FromKnowledgeTank = "1" Then
								Str1 = Str1 & Replace( Replace( Stra, "{0}", HyftdGrouping1(index, 2) ), "{1}", HyftdGrouping1(index, 1) )	
							Else
								Str1 = Str1 & Replace( Replace( Replace(Str0, "{0}", HyftdGrouping1(index, 2) ), "{1}", HyftdGrouping1(index, 1) ), "{2}", Replace(HyftdGrouping1(index, 0), " ", "@")) 'TempArray(1) )										
							End If
						Else
							Str1 = Str1 & Replace( Replace( Replace(Str0, "{0}", HyftdGrouping1(index, 2) ), "{1}", HyftdGrouping1(index, 1) ), "{2}", Replace(HyftdGrouping1(index, 0), " ", "@")) 'TempArray(1) )		
						End If						
					Next
					Response.Write Str1
				End If			
			%>
			</ul>
		</td>
	</tr>
	</table>

	<!--依時間分布-->
	<table class="layouttable">
	<tr><th scope="col">時序分布</th></tr>
	<tr>
		<td>
		<ul>
		<% 
			If HyftdGroupingCount > 0 Then 
				
				For index = 0 To HyftdGroupingCount - 1
				
					If Not IsNull(HyftdGrouping(index, 0)) AND HyftdGrouping(index, 0) <> "" Then
						
						Output = "<li><a href=""javascript:CheckSubmitForTime('" & HyftdGrouping(index, 0) & "');"">"
						If Len(HyftdGrouping(index, 0)) = 6 Then
							Output = Output & Mid(HyftdGrouping(index, 0), 1, 4) & "/" & Mid(HyftdGrouping(index, 0), 5, 2)
							Output = Output & "(" & HyftdGrouping(index, 1) & ")"
						End If
						Output = Output & "</a></li>"
						
						Response.Write Output
						
					End If
				
				Next
				
		 	End If 
		%>
		</ul>
		</td>
	</tr>
	</table>
</td>

<td class="center">
	<table class="layouttable">
	<tr>
		<td>
			<p>查詢條件：<%=Keyword%><br/>查詢結果： <%=HyftdTotalRecordCount-NullRecordCount%> 筆，依
	    <% If Order = "0" Then %>
	  		關聯度 
	  	<% ElseIf Order = "1" Then	%>
    		標題 
    	<% ElseIf Order = "2" Then		%>			
    		發佈日期 
    	<% ElseIf Order = "3" Then		%>		
    		資料來源 
    	<% ElseIf Order = "4" Then		%>		
    		點閱次數
		<% End If %>
		
		<% If Sort = "0" Then %>
			遞減 
		<% ElseIf Sort = "1" Then %>
			遞增 
		<% End If %>
			排序， 搜尋共費 <%=HyftdTimeDiff%> 秒</p>
			<!--div class="relatedword">
			<ul>
				更多相關詞：
				<li><a href="#">......</a></li>				
			</ul>
			</div-->
		</td>
	</tr>
	</table>
	
	<div class="Page">		
		第<span class="Number">
		<%
			If HyftdTotalRecordCount-NullRecordCount = 0 Then
				Response.Write "0"
			Else
				Response.Write PageNumber
			End If
		%>/<%=TotalPage%></span>頁，共<span class="Number"><%=HyftdTotalRecordCount-NullRecordCount%></span>筆，
		<% If CInt(PageNumber) > 1 Then %>
			<a href="javascript:PreviousPageOnClick(<%=PageNumber - 1%>);"><img src="/xslgip/style1/images3/arrow_left.gif" alt="左方箭頭" hspace="5" />上一頁</a>，
		<% End If %>
		跳至第
		<select name="PageNumber" onchange="PageNumberOnChange()">
		<%		
			For index = 1 To TotalPage
				If CInt(PageNumber) = index Then
					Response.Write "<option value=""" & index & """ selected=""selected"">" & index & "</option>"
				Else
					Response.Write "<option value=""" & index & """>" & index & "</option>"
				End If
			Next
		%>
		</select>
		<label for="select">頁</label>，
		<% If CInt(PageNumber) < TotalPage Then %>
			<a href="javascript:NextPageOnClick(<%=PageNumber + 1%>);">下一頁<img src="/xslgip/style1/images3/arrow_right.gif" alt="右方箭頭" hspace="5" /></a>，
		<% End If %>
		每頁
		<select id="Pa" name="PageSize" onchange="PageSizeOnChange()">
			<option value="10" <% If PageSize = 10 Then %> selected="selected" <% End If %> >10</option>
			<option value="30" <% If PageSize = 30 Then %> selected="selected" <% End If %> >30</option>
			<option value="50" <% If PageSize = 50 Then %> selected="selected" <% End If %> >50</option>
		</select>
		<label for="select">筆</label>資料	
	</div>

	<div class="listorder">
		瀏覽排序依：
		<select name="Order" onchange="OrderOnChange()">
			<option value="0" <% If Order = "0" Then %> selected="selected" <% End If %> >關聯度</option>
			<option value="1" <% If Order = "1" Then %> selected="selected" <% End If %> >標題</option>
			<option value="2" <% If Order = "2" Then %> selected="selected" <% End If %> >發布日期</option>
			<option value="3" <% If Order = "3" Then %> selected="selected" <% End If %> >資料來源</option>			
			<option value="4" <% If Order = "4" Then %> selected="selected" <% End If %> >點閱次數</option>				
	
		</select>
		<%if Order<>"0" then%>
		<select name="Sort" onchange="SortOnChange()">
			<option value="0" <% If Sort = "0" Then %> selected="selected" <% End If %> >遞減</option>			
			<option value="1" <% If Sort = "1" Then %> selected="selected" <% End If %> >遞增</option>			
		</select>
		<%end if %>
	</div>

	<div class="searchlist">		
		<table width="100%" border="0" cellspacing="0" cellpadding="0" summary="搜尋結果條列顯示">
		<tr>
			<th scope="col">&nbsp;</th>
			<th scope="col">標題</th>
			<th scope="col">發布日期</th>
			<th scope="col">資料來源</th>
			<th scope="col">點閱次數</th>
		</tr>		
	
		<% rspIndex = 0
		    For index = 0 To PageSize - 1 %>
			<% If ResultArray(index, 0) <> "" Then 
		       rspIndex = rspIndex+1 	
			%>
				<tr>
					<td><%=rspIndex%></td>
						<td><%
							If ResultArray(index, 2) = "0" Then	'---知識庫---

					            select case ResultArray(index, 4)
					                case KnowledgeTankDatabaseId
								            TankUrl = Replace( "/CatTree/CatTreeContent.aspx?ReportId={0}&DatabaseId={1}&CategoryId={2}&ActorType={3}", "{0}", ResultArray(index, 3) )
								            TankUrl = Replace( TankUrl, "{1}", ResultArray(index, 4) )
								            TankUrl = Replace( TankUrl, "{2}", ResultArray(index, 5) )
								            TankUrl = Replace( TankUrl, "{3}", ResultArray(index, 16) )
								            
								            TankUrl = "/category/categorycontent.aspx?ReportId={0}&CategoryId={2}&ActorType={3}&kpi=0"
								            TankUrl = Replace( TankUrl, "{0}", ResultArray(index, 18))
								            TankUrl = Replace( TankUrl, "{2}", ResultArray(index,  5))
								            TankUrl = Replace( TankUrl, "{3}", ResultArray(index, 16)) 
								            
								            Response.Write "<a href=""" & TankUrl & """ target=""_blank"">" & ResultArray(index, 7) & "</a>"
            								
					                case "DB023"
					                        if left(ResultArray(index, 11),1) = "/" then
					                            Response.Write "<a href=""/techCD" & ResultArray(index, 11) & """ target=""_blank"">" & ResultArray(index, 7) & "</a>"
					                        else
					                            Response.Write "<a href=""/techCD/" & ResultArray(index, 11) & """ target=""_blank"">" & ResultArray(index, 7) & "</a>"
					                        end if
					             end select							
								
							ElseIf ResultArray(index, 2) = "1" Then	'---站內單元---
											
								If ResultArray(index, 10) = "1" Then
									
									SiteUrl = Replace( session("myURL") & "ct.asp?xItem={0}&ctNode={1}", "{0}", ResultArray(index, 3) )
									SiteUrl = Replace( SiteUrl, "{1}", ResultArray(index, 13) )
									Response.Write "<a href=""" & SiteUrl & """ target=""_blank"">" & ResultArray(index, 7) & "</a>"
									
								ElseIf ResultArray(index, 10) = "2" Then
									
									Response.Write "<a href=""" & ResultArray(index, 11) & """ target=""_blank"">" & ResultArray(index, 7) & "</a>"
									
								ElseIf ResultArray(index, 10) = "3" Then
									
									Response.Write "<a href=""" & session("myURL") & "public/Data/" & ResultArray(index, 12) & """ target=""_blank"">" & ResultArray(index, 7) & "</a>"	
								
								End If
								
							ElseIf ResultArray(index, 2) = "2" Then	'---主題館---
																
								If ResultArray(index, 10) = "1" Then
									
									TopicUrl = Replace( session("myURL") & "subject/ct.asp?xItem={0}&ctNode={1}&mp={2}", "{0}", ResultArray(index, 3) )
									TopicUrl = Replace( TopicUrl, "{1}", ResultArray(index, 13) )
									TopicUrl = Replace( TopicUrl, "{2}", ResultArray(index, 15) )
									Response.Write "<a href=""" & TopicUrl & """ target=""_blank"">" & ResultArray(index, 7) & "</a>"	
									
								ElseIf ResultArray(index, 10) = "2" Then
									
									Response.Write "<a href=""" & ResultArray(index, 11) & """ target=""_blank"">" & ResultArray(index, 7) & "</a>"	
									
								ElseIf ResultArray(index, 10) = "3" Then
									
									Response.Write "<a href=""" & session("myURL") & "subject/public/Data/" & ResultArray(index, 12) & """ target=""_blank"">" & ResultArray(index, 7) & "</a>"	
									
								End If
								
							ElseIf ResultArray(index, 2) = "3" Then	'---知識家---
								HomeUrl = Replace( session("myURL") & "knowledge/knowledge_cp.aspx?ArticleId={0}&ArticleType=A&CategoryId=E", "{0}", ResultArray(index, 3) )								
								Response.Write "<a href=""" & HomeUrl & """ target=""_blank"">" & ResultArray(index, 7) & "</a>"
							

							ElseIf ResultArray(index, 2) = "4" Then	'---農業小百科---
                                Response.Write("<a href='/Pedia/PediaContent.aspx?AId=" & ResultArray(index, 3) & "' target='_blank'>" & ResultArray(index, 7) & "</a>")
							End If
						%>						
						<span class="relper">關聯性: 
						<%
							If Order = "0" Then
								if(Round(ResultArray(index, 1) / MaxRelated * 100, 2)>100) then
									Response.Write("100%")
								else
									Response.Write Round(ResultArray(index, 1) / MaxRelated * 100, 2) & "%"
								end if								
							Else
								if(ResultArray(index, 1) > 100) then
									Response.Write("100%")
								else
									Response.Write ResultArray(index, 1) & "%"
								end if
							End If
						%>
						</span>
						<div class="content">
							<%							
								If Instr(ResultArray(index, 14), "<") = 1 Then
								ElseIf ResultArray(index, 14) = "" Then
								Else
									'Response.Write Replace(Replace(ResultArray(index, 14), "<br>", ""), "<BR>", "") & "..."								
								End If
							%>							
						</div>
					</td>
					<td width="10px">[<%=ResultArray(index, 9)%>]</td>
					<td>
					<%
					    select case ResultArray(index, 2)
					        case "0"
					            select case ResultArray(index, 4)
					                case KnowledgeTankDatabaseId
					                    Response.Write "知識庫"
					                case "DB023"
					                    Response.Write "技術光碟"
					             end select
					        case "1"
					           
					            select case ResultArray(index, 4)
					                case "41"
					                    Response.Write "影音專區"
					                case else
					                     Response.Write "站內單元"
					             end select					            
					        case "2"
					            Response.Write "主題館"
					        case "3"
					            Response.Write "知識家"
					        case "4"
					            Response.Write "農業小百科"
					     end select						
					%>	
					</td>
					<td width="10px" align="right"><%=ResultArray(index, 8) %></td>
				</tr>
			<% End If %>
		<% Next%>
		
		</table>
	</div>
	<div class="Page">		
		第<span class="Number">
		<%
			If HyftdTotalRecordCount-NullRecordCount = 0 Then
				Response.Write "0"
			Else
				Response.Write PageNumber
			End If
		%>/<%=TotalPage%></span>頁，共<span class="Number"><%=HyftdTotalRecordCount-NullRecordCount%></span>筆，
		<% If CInt(PageNumber) > 1 Then %>
			<a href="javascript:PreviousPageOnClick(<%=PageNumber - 1%>);"><img src="/xslgip/style1/images3/arrow_left.gif" alt="左方箭頭" hspace="5" />上一頁</a>，
		<% End If %>
		
		
		跳至第
		<select id="PageNumber_down" name="PageNumber_down"  onchange="MakePageNumberSame()" >
		<%		
			For index = 1 To TotalPage
				If CInt(PageNumber) = index Then
					Response.Write "<option value=""" & index & """ selected=""selected"">" & index & "</option>"
				Else
					Response.Write "<option value=""" & index & """>" & index & "</option>"
				End If
			Next
			
			
			
		%>
		</select>
		<label for="select">頁</label>，
		
		
		<% If CInt(PageNumber) < TotalPage Then %>
			<a href="javascript:NextPageOnClick(<%=PageNumber + 1%>);">下一頁<img src="/xslgip/style1/images3/arrow_right.gif" alt="右方箭頭" hspace="5" /></a>，
		<% End If %>
		
		每頁
		<select id="PageSize_down" name="PageSize_down" onchange="MakePageSizeSame()" >
			<option value="10" <% If PageSize = 10 Then %> selected="selected" <% End If %> >10</option>
			<option value="30" <% If PageSize = 30 Then %> selected="selected" <% End If %> >30</option>
			<option value="50" <% If PageSize = 50 Then %> selected="selected" <% End If %> >50</option>
		</select>
		<label for="select">筆</label>資料
		
	</div>
</td>
</tr>
</table>
</div>
</form>	
<%


	if HyftdTotalRecordCount>0 then    	
    	set http = server.createObject("microsoft.xmlhttp")	
        For i = 0 To UBound(KeywordArray)
            qStr= server.URLEncode("Keyword=" & KeywordArray(i))	        
	        http.open "POST", session("myURL") & "services/tagCloud.aspx?cmd=addtag&tag=" & qStr
	        http.send
	    next
	end if
	
	conn.close
	Set conn = Nothing	
	
	KmConn.close
	Set KmConn = Nothing	
%>