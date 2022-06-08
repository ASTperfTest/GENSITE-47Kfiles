<!-- #include file = "exam.class.asp" -->
<%
'取得題目主題
Function GetCuDtGeneric(oConnection, iCuItem)
	Dim oCuDtGeneric
	Dim sqlDtGeneric, rsDtGeneric

	sqlDtGeneric = "select * from CuDtGeneric where icuitem = " & iCuItem
	Set rsDtGeneric = Server.CreateObject("ADODB.RecordSet")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rsDtGeneric.Open sqlDtGeneric, oConnection, 1, 1
Set rsDtGeneric =  oConnection.execute(sqlDtGeneric)

'----------HyWeb GIP DB CONNECTION PATCH----------

	
	If Not rsDtGeneric.EOF Then
		Set oCuDtGeneric = New CuDtGeneric
		oCuDtGeneric.SetField "Id", rsDtGeneric.Fields("icuitem").Value
		oCuDtGeneric.SetField "Title", rsDtGeneric.Fields("stitle").Value
		oCuDtGeneric.SetField "TopCat", rsDtGeneric.Fields("topCat").Value
	Else
		Err.Raise vbObjectError + 1, "", "主題不存在"
	End If
	
	rsDtGeneric.Close
	Set rsDtGeneric = Nothing
	
	Set GetCuDtGeneric = oCuDtGeneric
End Function

'取得腔調
' 0:全部  1:交集  2:差集
Function GetTuneList(oConnection, iCuItem, iFilterType)
	Dim arrayList()
	
	sqlTuneList = "select c.codeMetaId, c.mcode, c.mvalue, c.msortValue, t.et_id " & _
	              "from CodeMain c " & _
	              "left join hakkaExamTopic t on t.tune_id = c.mcode and t.icuitem = " & iCuItem & "" & _
	              "where c.codeMetaId = 'hakkalangtone' " & _
	              "order by c.msortValue asc"
	Set rsTuneList = Server.CreateObject("ADODB.RecordSet")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rsTuneList.Open sqlTuneList, oConnection, 1, 1
Set rsTuneList =  oConnection.execute(sqlTuneList)

'----------HyWeb GIP DB CONNECTION PATCH----------

	
	If Not rsTuneList.EOF Then
		ReDim arrayList(1, rsTuneList.RecordCount - 1)
		iTuneList = 0
		While Not rsTuneList.EOF
			
			If iFilterType = 0 Or (iFilterType = 1 And Not IsNull(rsTuneList.Fields("et_id").Value)) Or (iFilterType = 2 And IsNull(rsTuneList.Fields("et_id").Value)) Then
				arrayList(0, iTuneList) = rsTuneList("mcode")
				arrayList(1, iTuneList) = rsTuneList("mvalue")
				iTuneList = iTuneList + 1
			End If
			
			rsTuneList.MoveNext
		Wend
		
		If iTuneList <> 0 Then
			ReDim Preserve arrayList(1, iTuneList - 1)
		Else
			Err.Raise vbObjectError + 1, "", "無法取得腔調"
		End If
	Else
		Err.Raise vbObjectError + 1, "", "無法取得腔調"
	End If
	
	rsTuneList.Close
	Set rsTuneList = Nothing
	
	GetTuneList = arrayList
End Function

'取得題目類型
Function GetExamType(oConnection)
	Dim arrayList()
	
	sqlExamTypeList = "select codeMetaId, mcode, mvalue, msortValue " & _
	              "from CodeMain " & _
	              "where codeMetaId = 'exam_topcat' " & _
	              "order by msortValue asc"
	Set rsExamTypeList = Server.CreateObject("ADODB.RecordSet")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rsExamTypeList.Open sqlExamTypeList, oConnection, 1, 1
Set rsExamTypeList =  oConnection.execute(sqlExamTypeList)

'----------HyWeb GIP DB CONNECTION PATCH----------

	
	If Not rsExamTypeList.EOF Then
		ReDim arrayList(1, rsTuneList.RecordCount - 1)
		For iExamTypeList = 0 To UBound(arrayList)
			arrayList(0, iTuneList) = rsExamTypeList("mcode")
			arrayList(1, iTuneList) = rsExamTypeList("mvalue")
			rsExamTypeList.MoveNext
		Next
	Else
		Err.Raise vbObjectError + 1, "", "無法取得題目類型"
	End If
	
	rsExamTypeList.Close
	Set rsExamTypeList = Nothing
	
	GetExamType = arrayList
End Function

' 取得腔調題目
Function GetTopic(oConnection, iTopicId)
	Dim oExamTopic
	Dim sqlTopic, rsTopic
	
	sqlTopic = "select * from hakkalangExamTopic where et_id = " & iTopicId & ""
	Set rsTopic = Server.CreateObject("ADODB.RecordSet")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rsTopic.Open sqlTopic, oConnection, 1, 1
Set rsTopic =  oConnection.execute(sqlTopic)

'----------HyWeb GIP DB CONNECTION PATCH----------


	If Not rsTopic.EOF Then
		Set oExamTopic = New ExamTopic
		oExamTopic.SetField "Id", rsTopic.Fields("et_id").Value
		oExamTopic.SetField "CuItemId", rsTopic.Fields("icuitem").Value
		oExamTopic.SetField "TuneId", rsTopic.Fields("tune_id").Value
		oExamTopic.SetField "Correct", rsTopic.Fields("et_correct").Value
	Else
		Err.Raise vbObjectError + 1, "", "無法取得腔調題目"
	End If

	rsTopic.Close
	Set rsTopic = Nothing
	
	Set GetTopic = oExamTopic
End Function

' 取得腔調題目 by gicuitem
Function GetTopicByCuitem(oConnection, iCuItem)
	Dim oExamTopic
	Dim sqlTopic, rsTopic
	
	sqlTopic = "select * from hakkalangExamTopic where gicuitem = " & iCuItem & ""
	Set rsTopic = Server.CreateObject("ADODB.RecordSet")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rsTopic.Open sqlTopic, oConnection, 1, 1
Set rsTopic =  oConnection.execute(sqlTopic)

'----------HyWeb GIP DB CONNECTION PATCH----------


	If Not rsTopic.EOF Then
		Set oExamTopic = New ExamTopic
		oExamTopic.SetField "Id", rsTopic.Fields("et_id").Value
		oExamTopic.SetField "CuItemId", rsTopic.Fields("gicuitem").Value
		oExamTopic.SetField "TuneId", rsTopic.Fields("tune_id").Value
		oExamTopic.SetField "Correct", rsTopic.Fields("et_correct").Value
	Else
		Err.Raise vbObjectError + 1, "", "無法取得腔調題目"
	End If

	rsTopic.Close
	Set rsTopic = Nothing
	
	Set GetTopicByCuitem = oExamTopic
End Function

' 新增腔調題目
Function CreateTopic(oConnection, oExamTopic)
	Dim sqlTopic, rsTopic
	
	oExamTopic.Validate
	
	sqlTopic = "hakkalangExamTopic"
	Set rsTopic = Server.CreateObject("ADODB.RecordSet")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rsTopic.Open sqlTopic, oConnection, 1, 3
Set rsTopic =  oConnection.execute(sqlTopic)

'----------HyWeb GIP DB CONNECTION PATCH----------

	
	rsTopic.AddNew
	rsTopic("gicuitem") = oExamTopic.GetField("CuItemId")
	rsTopic("tune_id") = oExamTopic.GetField("TuneId")
	rsTopic("et_correct") = oExamTopic.GetField("Correct")
	rsTopic.Update
	
	oExamTopic.SetField "Id", rsTopic.Fields("et_id").Value
	
	rsTopic.Close
	Set rsTopic = Nothing
End Function


' 更新腔調題目
Function UpdateTopic(oConnection, oExamTopic)
	Dim sqlTopic, rsTopic
	
	oExamTopic.Validate
	
	sqlTopic = "select * from hakkalangExamTopic where et_id = " & oExamTopic.GetField("Id") & ""
	Set rsTopic = Server.CreateObject("ADODB.RecordSet")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rsTopic.Open sqlTopic, oConnection, 1, 3
Set rsTopic =  oConnection.execute(sqlTopic)

'----------HyWeb GIP DB CONNECTION PATCH----------

	
	If Not rsTopic.EOF Then
		'rsTopic("gicuitem") = oExamTopic.GetField("CuItemId")
		'rsTopic("tune_id") = oExamTopic.GetField("TuneId")
		rsTopic("et_correct") = oExamTopic.GetField("Correct")
		rsTopic.Update
	Else
		Err.Raise vbObjectError + 1, "", "題目不存在"
	End If
	
	rsTopic.Close
	Set rsTopic = Nothing
End Function

' 刪除腔調題目
Function DelTopic(oConnection, oExamTopic)
	Dim sqlTopic
	
	sqlTopic = "delete from hakkalangExamTopic where et_id = " & oExamTopic.GetField("Id")
	oConnection.Execute(sqlTopic)
End Function

' 取得題目選項
Function GetOption(oConnection, iOptionId, iOptionSort)
	Dim oExamOption
	Dim sqlOption, rsOption
	
	sqlOption = "select * from hakkalangExamOption where eo_id = " & iOptionId & " and eo_sort = " & iOptionSort & " order by eo_sort asc"
	Set rsOption = Server.CreateObject("ADODB.RecordSet")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rsOption.Open sqlOption, oConnection, 1, 1
Set rsOption =  oConnection.execute(sqlOption)

'----------HyWeb GIP DB CONNECTION PATCH----------

	
	If Not rsOption.EOF Then
		Set oExamOption = New ExamOption
		oExamOption.SetField "TopicId", rsOption.Fields("et_id").Value
		oExamOption.SetField "Title", rsOption.Fields("eo_title").Value
		oExamOption.SetField "Answer", rsOption.Fields("eo_answer").Value
		oExamOption.SetField "Sort", rsOption.Fields("eo_sort").Value
	Else
		Err.Raise vbObjectError + 1, "", "無法取得題目選項"
	End If
	
	rsOption.Close
	Set rsOption = Nothing
	
	Set GetOption = oExamOption
End Function

' 取得題目所有選項
Function GetAllOption(oConnection, iTopicId)
	Dim dExamOptions
	Dim oExamOption
	Dim sqlAllOptions, rsAllOption
	
	Set dExamOptions = Server.CreateObject("Scripting.Dictionary")
	
	sqlAllOptions = "select * from hakkalangExamOption where et_id = " & iTopicId & " order by eo_sort asc"
	Set rsAllOption = Server.CreateObject("ADODB.RecordSet")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rsAllOption.Open sqlAllOptions, oConnection, 1, 1
Set rsAllOption =  oConnection.execute(sqlAllOptions)

'----------HyWeb GIP DB CONNECTION PATCH----------

	
	If Not rsAllOption.EOF Then
		While Not rsAllOption.EOF
			Set oExamOption = New ExamOption
			oExamOption.SetField "TopicId", rsAllOption.Fields("et_id").Value
			oExamOption.SetField "Title", rsAllOption.Fields("eo_title").Value
			oExamOption.SetField "Answer", rsAllOption.Fields("eo_answer").Value
			oExamOption.SetField "Sort", rsAllOption.Fields("eo_sort").Value
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

' 新增題目選項
Function CreateOption(oConnection, oExamOption)
	Dim sqlOption, rsOption
	
	oExamOption.Validate
	
	sqlOption = "hakkalangExamOption"
	Set rsOption = Server.CreateObject("ADODB.RecordSet")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rsOption.Open sqlOption, oConnection, 1, 3
Set rsOption =  oConnection.execute(sqlOption)

'----------HyWeb GIP DB CONNECTION PATCH----------

	
	rsOption.AddNew
	rsOption("et_id") = oExamOption.GetField("TopicId")
	rsOption("eo_title") = oExamOption.GetField("Title")
	rsOption("eo_answer") = oExamOption.GetField("Answer")
	rsOption("eo_sort") = oExamOption.GetField("Sort")
	rsOption.Update
	
	rsOption.Close
	Set rsOption = Nothing
End Function

' 更新題目選項
Function UpdateOption(oConnection, oExamOption)
	Dim sqlOption, rsOption
	
	oExamOption.Validate
	
	sqlOption = "select * from hakkalangExamOption where et_id = " & oExamOption.GetField("TopicId") & " and eo_sort = " & oExamOption.GetField("Sort") & ""
	Response.Write sqlOption
	Set rsOption = Server.CreateObject("ADODB.RecordSet")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rsOption.Open sqlOption, oConnection, 1, 3
Set rsOption =  oConnection.execute(sqlOption)

'----------HyWeb GIP DB CONNECTION PATCH----------

	
	If Not rsOption.EOF Then
		rsOption("et_id") = oExamOption.GetField("TopicId")
		rsOption("eo_title") = oExamOption.GetField("Title")
		rsOption("eo_answer") = oExamOption.GetField("Answer")
		rsOption("eo_sort") = oExamOption.GetField("Sort")
		rsOption.Update
	Else
		Err.Raise vbObjectError + 1, "", "選項不存在"
	End If
	
	rsOption.Close
	Set rsOption = Nothing
End Function

' 刪除題目選項
Function DelOption(oConnection, oExamOption)
	Dim sqlOption
	
	sqlOption = "delete from hakkalangExamOption where et_id = " & oExamOption.GetField("TopicId") & " and eo_sort = " & oExamOption.GetField("Sort") & ""
	oConnection.Execute(sqlOption)
End Function


' 刪除題目所有選項
Function DelAllOption(oConnection, oExamTopic)
	Dim sqlOption
	
	sqlOption = "delete from hakkalangExamOption where et_id = " & oExamTopic.GetField("Id")
	oConnection.Execute(sqlOption)
End Function
%>
