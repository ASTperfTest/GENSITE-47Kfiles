<%@ CodePage = 65001 %>
<%Response.ContentType="text/html; charset=utf-8"%>
<!-- #include file = "exam_util.inc.asp" -->
<%
'tune_id=request.QueryString("tuneID")
ExamCount=request.QueryString("eCount")
Examtype=request.QueryString("examtype")
xSeason=request.QueryString("xSeason")
xCity=request.QueryString("xCity")

'IcuitemList = GetExamItemList(ExamCount,Examtype,tune_id)  
IcuitemList = GetExamItemList(ExamCount,Examtype,xSeason,xCity)  
'for i=0 to UBOUND(IcuitemList)
for i=0 to 0
    Set oTopic=GetTopicByCuitem(IcuitemList(i))	 
		eId=oTopic.GetField("et_Id")
		stitle=oTopic.GetField("ct_stitle")
	 	iExamType=oTopic.GetField("ct_examtype")
	    Answer=oTopic.GetField("et_correct")
	    URLtemp=URLtemp & "&title" & i+1 & "=" & stitle
		'取得目前題目選項 
				IF iExamType<>"1"  then ' <>是非 
					set options=GetAllOption(eId )	
						for x=0 to options.count-1
								xOption=options(x).GetField("OptTitle")					  
								URLtemp=URLtemp & "&C" & i+1 & x+1 & "=" & xOption		 		
						next 				
				End IF				
				URLtemp=URLtemp & "&A" & i+1 & "=" & Answer	
Next
response.write URLtemp
%>
 
