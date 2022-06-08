<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include file = "exam_class.asp" -->

<%
'隨機抽題,取得題目主題ID
Function GetExamItemList(oCount,oTypeid,oSeason,oCity)  '題數,題型,季節,區域城市
	Dim arrayList  'arrayList()不要用括號	
'	Dim arr_all
	Dim arr_temp 
	Dim oCuDtGeneric
	Dim sqlDtGeneric, rsDtGeneric
	Dim strParm	
 
   Select Case (oSeason)
   Case "A":
       '春 3,4,5
	   SQLxabstract="(gx.xabstract like '%3%') or (gx.xabstract like '%4%') or (gx.xabstract like '%5%')"
   Case "B":
       '夏6,7,8
       SQLxabstract="(gx.xabstract like '%6%') or (gx.xabstract like '%7%') or (gx.xabstract like '%8%')" 
   Case "C":
       '秋9,10,11 
          SQLxabstract="(gx.xabstract like '%9%') or (gx.xabstract like '%10%') or (gx.xabstract like '%11%')" 
   Case "D":
       '冬,12,1,2
	      SQLxabstract="(gx.xabstract like '%12%') or (gx.xabstract like '%1%') or (gx.xabstract like '%2%')" 
   Case else:
		'春 3,4,5 2011/05/06 sam_chen 不給default會有sqlserver db error
	   SQLxabstract="(gx.xabstract like '%3%') or (gx.xabstract like '%4%') or (gx.xabstract like '%5%')"
   End Select
   
	 sqlDtGeneric = " select  Top " & oCount & " gx.icuitem, gx.stitle,gx.topCat,gx.xabstract,gx.xbody, tx.tune_id, tx.et_id ,tx.et_correct " & _
 	                " from CuDtGeneric gx INNER JOIN hakkalangExamTopic tx ON gx.icuitem = tx.gicuitem " & _
				    " where  (gx.ictunit<>'633') And (gx.fctupublic='Y') And (" & SQLxabstract & ")" & _
 				    " And (gx.xbody like '%" & oCity & "%')"

	'sqlDtGeneric = " select  gx.icuitem, gx.stitle,gx.topCat,gx.xabstract,gx.xbody, tx.tune_id, tx.et_id ,tx.et_correct " & _
 	               '" from CuDtGeneric gx INNER JOIN hakkalangExamTopic tx ON gx.icuitem = tx.gicuitem " & _
				   '" where  (gx.ictunit<>'440') And (gx.fctupublic='Y') " 
				   
	if  oTypeid<>"0" then sqlDtGeneric=sqlDtGeneric & " AND (gx.topCat = " &  pkstr(oTypeid,"") & ")"
 		sqlDtGeneric=sqlDtGeneric & " order by  NEWID() desc"
 	
    'response.Write sqlDtGeneric & "<br>"	
  	Set rsExamList = Server.CreateObject("ADODB.RecordSet")
	rsExamList.Open sqlDtGeneric, Conn, 1, 1

	dim filterOK  
 	If Not rsExamList.EOF Then
	icount=0
    While Not rsExamList.EOF
		
		 Season_filterOK="N"
		 City_filterOK="N"
		 arrSeason=""
		 arrCity=""
         arrSeason=split(rsExamList("xabstract"),",")
		 arrCity=split(rsExamList("xbody"),",")
         		  
		 
		 for s=0 to UBOUND(arrSeason) 
		  
           Select Case int(arrSeason(s))
			   Case 3,4,5: '春
			        IF oSeason="A" THEN Season_filterOK="Y"	 	
					'EXIT FOR
			   Case 6,7,8: '夏			        
			        IF oSeason="B" THEN Season_filterOK="Y"
					'EXIT FOR 
			   Case 9,10,11:'秋			       
			        IF oSeason="C" THEN Season_filterOK="Y"
					'EXIT FOR
			   Case 12,1,2:'冬
			        IF oSeason="D" THEN Season_filterOK="Y"		
                    'EXIT FOR	   
               Case else:					
           End Select		 
         next 
		  'response.write "月份:" & rsExamList("xabstract") & "<br>"
		  'response.write "縣市:" & rsExamList("xbody") & "<br>" 
		 'response.write Season_filterOK
		   for c=0 to UBOUND(arrCity) 
		       'RESPONSE.WRITE int(arrCity(c)) &  ":" & int(oCity) & "<BR/>"
		       IF int(arrCity(c))=int(oCity) then City_filterOK="Y"
		   next 
         'response.write City_filterOK
         
		  'response.write rsExamList("icuitem")
		 
          IF (Season_filterOK="Y"  AND City_filterOK="Y") THEN 
			  icount=icount+1		
				if sid="" then
					sid=rsExamList.Fields("icuitem").Value
				else
					sid= sid & "," &  rsExamList.Fields("icuitem").Value
				end if
		  end if
		  
			' ReDim arrayList(icount-1)		
			 ' For i = 0 To UBound(arrayList)
					' arrayList(i) = rsExamList.Fields("icuitem").Value
			 ' Next
			 
		rsExamList.MoveNext		             
	WEND
				 
        		
	Else
		Err.Raise vbObjectError + 1, "", "抽題失敗"
		 
	End If
	        '  RESPONSE.WRITE  "A" & SID & ":"
			 
			     arrayList=split(sid,",")
			   
			 
	
	rsExamList.Close
	Set rsExamList = Nothing
	
	 IF SID="" THEN  
	     GetExamItemList=""
	 else
	       GetExamItemList=arrayList
	 end if
End Function


Function GetExamItemList_old(oCount,oTypeid,oTuneid)  '題數,題型,腔調
	Dim arrayList  'arrayList()不要用括號
	
	Dim oCuDtGeneric
	Dim sqlDtGeneric, rsDtGeneric
 
	sqlDtGeneric = "select  Top " & oCount & " gx.icuitem, gx.stitle,gx.topCat, tx.tune_id, tx.et_id ,tx.et_correct " & _
 	                       "from CuDtGeneric gx INNER JOIN hakkalangExamTopic tx ON gx.icuitem = tx.gicuitem " & _
								  "where  (gx.ictunit<>'570') and  (tx.tune_id = " & pkstr(oTuneid,"") & ") And (gx.fctupublic='Y') " 
		  
	if  oTypeid<>"0" then sqlDtGeneric=sqlDtGeneric & " AND (gx.topCat = " &  pkstr(oTypeid,"") & ")"
 		sqlDtGeneric=sqlDtGeneric & " order by  NEWID() desc"
 	 
  	Set rsExamList = Server.CreateObject("ADODB.RecordSet")
	rsExamList.Open sqlDtGeneric, Conn, 1, 1

	'  response.Write sqlDtGeneric & "<br>"
 	If Not rsExamList.EOF Then
		'ReDim arrayList(rsExamList.RecordCount - 1,3 )	 '重設array把變動的放後面	
		'ReDim arrayList(7, rsExamList.RecordCount - 1)		
		ReDim arrayList(rsExamList.RecordCount - 1)		
			 For i = 0 To UBound(arrayList)
					arrayList(i) = rsExamList.Fields("icuitem").Value
						 					
					 rsExamList.MoveNext		             
					
			Next  
	Else
		Err.Raise vbObjectError + 1, "", "抽題失敗"
	End If
	
	rsExamList.Close
	Set rsExamList = Nothing
	GetExamItemList=arrayList
	
End Function


 

Function GethazardItemList(oCount)  '題數,題型,腔調
	Dim arrayList  'arrayList()不要用括號
	
	Dim oCuDtGeneric
	Dim sqlDtGeneric, rsDtGeneric
 
	sqlDtGeneric = "select  Top " & oCount & " gx.icuitem, gx.stitle,gx.topCat, tx.tune_id, tx.et_id ,tx.et_correct " & _
 	                       "from CuDtGeneric gx INNER JOIN hakkalangExamTopic tx ON gx.icuitem = tx.gicuitem " & _
								  "where gx.ictunit='633' And (gx.fctupublic='Y') order by  NEWID() desc" 
  

  	
	Set rsExamList = Server.CreateObject("ADODB.RecordSet")
	rsExamList.Open sqlDtGeneric, Conn, 1, 1

	'  response.Write sqlDtGeneric & "<br>"
 	If Not rsExamList.EOF Then
		'ReDim arrayList(rsExamList.RecordCount - 1,3 )	 '重設array把變動的放後面	
		'ReDim arrayList(7, rsExamList.RecordCount - 1)		
		ReDim arrayList(rsExamList.RecordCount - 1)		
			 For i = 0 To UBound(arrayList)
					arrayList(i) = rsExamList.Fields("icuitem").Value
						 					
					 rsExamList.MoveNext		             
					
			Next  
	Else
		Err.Raise vbObjectError + 1, "", "抽題失敗"
	End If
	
	rsExamList.Close
	Set rsExamList = Nothing
	GethazardItemList=arrayList
	
End Function


'取題目 
Function GetTopicByCuitem(CuItemId)
    Dim oExamTopic
	Dim sqlTopic, rsTopic
	
	sqlTopic="select * from CuDtGeneric g, hakkalangExamTopic t where g.icuitem=t.gicuitem and  g.icuitem=" & CuItemId 	
	Set rsTopic = Server.CreateObject("ADODB.RecordSet")
	rsTopic.Open sqlTopic, Conn, 1, 1
'response.write sqlTopic
	If Not rsTopic.EOF Then
		Set oExamTopic = New ExamTopic
		oExamTopic.SetField "ct_id", rsTopic.Fields("icuitem").Value
		oExamTopic.SetField "et_Id", rsTopic.Fields("et_id").Value
		oExamTopic.SetField "ct_stitle", rsTopic.Fields("stitle").Value
		oExamTopic.SetField "ct_examtype", rsTopic.Fields("topCat").Value
		oExamTopic.SetField "et_tuneId", rsTopic.Fields("tune_id").Value
		oExamTopic.SetField "et_correct", rsTopic.Fields("et_correct").Value

	Else
		Err.Raise vbObjectError + 1, "", "無法取得腔調題目"
	End If

	rsTopic.Close
	Set rsTopic = Nothing	
	Set GetTopicByCuitem = oExamTopic	
	
	
End Function

' 取得題目所有選項
Function GetAllOption(iTopicId)
	Dim dExamOptions
	Dim oExamOption
	Dim sqlAllOptions, rsAllOption
	
	Set dExamOptions = Server.CreateObject("Scripting.Dictionary")
	
	sqlAllOptions = "select * from hakkalangExamOption where et_id = " & iTopicId & " order by eo_sort asc"
	Set rsAllOption = Server.CreateObject("ADODB.RecordSet")
	rsAllOption.Open sqlAllOptions, Conn, 1, 1
	'response.write sqlAllOptions
 
	If Not rsAllOption.EOF Then
		While Not rsAllOption.EOF
			Set oExamOption = New ExamOption
			oExamOption.SetField "Optetid", rsAllOption.Fields("et_id").Value
			oExamOption.SetField "OptTitle", rsAllOption.Fields("eo_title").Value
			oExamOption.SetField "OptAnswer", rsAllOption.Fields("eo_answer").Value
			oExamOption.SetField "OptSort", rsAllOption.Fields("eo_sort").Value			
			dExamOptions.Add dExamOptions.Count, oExamOption
			rsAllOption.MoveNext
		Wend
	Else
		Err.Raise vbObjectError + 1, "", "無法取得題目選項"
	End If
	
	rsAllOption.Close
	Set rsAllOption = Nothing
	
	Set GetAllOption = dExamOptions
End Function

' 轉換配合題甲乙丙
Function Changli(x)
   Dim litype(7)
   Select Case x
			Case 0:
				Changli="甲."
			Case 1:
				Changli="乙."
			Case 2:
				Changli="丙."
			Case 3:
				Changli="丁."
			Case 4:
				Changli="戊."
			Case 5:
				Changli="己."
			Case 6:
				Changli="庚."
	End Select				
End Function

%>