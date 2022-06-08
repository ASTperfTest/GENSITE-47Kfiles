<script language="jscript" runat="server"> 
function GetTime() 
{ 
	var d = new Date(); 
	return d.getTime(); 
} 
</script> 
<%
	'-------------------------------------------------------------------------------
	'---Hyftd-----------------------------------------------------------------------	
	Dim HyftdObj 										'---Hyftd物件---
	Dim HyftdServer 								'---Server Address---
	Dim HyftdPort 									'---Server Port---
	Dim HyftdUserId									'---User Id---
	Dim HyftdUserPassword						'---User Password---
	Dim HyftdGroupName 							'---Group Name---
	Dim HyftdConnId 								'---建立連線的Id---
	Dim HyftdSortQueryId 						'---查詢的Id(sortquery)---
	Dim HyftdQueryExId							'---查詢的Id(query_ex)---
	Dim HyftdQueryGroupingId 				'---查詢的Id(query_grouping)---
	Dim HyftdAddQueryId 						'---加入查詢的Id---
	Dim HyftdAssociateDataName 			'---AssociateData Name---
  Dim HyftdTotalRecordCount				'---總查詢筆數---
  Dim HyftdCurrentRecordCount 		'---目前回傳的筆數---
  Dim HyftdMaxQueryCount 					'---查詢筆數的上限---
  Dim HyftdQueryType 							'---查詢方式,0~4---
  Dim HyftdSortName 							'---做排序用的某個欄位名稱---
  Dim HyftdSortType 							'---遞增,遞減---
  Dim HyftdDataId 								'---某筆結果的Id(sysid)---
  Dim HyftdDataRelatedNumber 			'---關聯度---
  Dim HyftdDataLength 						'---某筆結果的AssociateData的長度---
  Dim HyftdPhonetic 							'---查詢時加入同音---
  Dim HyftdAuthorize 							'---查詢時加入同義---
  Dim HyftdStartTime 							'---查詢開始時間---
  Dim HyftdEndTime 								'---查詢結束時間---
  Dim HyftdTimeDiff 							'---查詢所花費時間---
  Dim HyftdRecordFrom 						'---查詢時所要抓取的開始筆數---
  Dim HyftdRecordTo 							'---查詢時所要抓取的結束筆數---
  Dim HyftdX 											'---呼叫Function時的指定值---
  Dim HyftdDebug 									'---Debug Mode---
	Dim HyftdGrouping() 						'---存放後分類的陣列---
	Dim HyftdGroupingCount 					'---後分類的數目---
	Dim HyftdGroupingName 					'---後分類名稱---
	Dim HyftdGroupingType 					'---後分類名稱排序---
	Dim HyftdGrouping1() 						'---存放後分類的陣列---
	Dim HyftdGroupingCount1					'---後分類的數目---
	Dim HyftdGroupingName1 					'---後分類名稱---
	Dim HyftdGroupingType1 					'---後分類名稱排序---
	'-------------------------------------------------------------------------------
	'---Function Hyftd Initialization-----------------------------------------------
	
	Sub Hyftd_Initial_Parameter( PageNumber, PageSize, Debug, Order, Sort, Phonetic, Authorize )
		
		HyftdServer = Application("HyftdServer")
  	HyftdPort = Application("HyftdPort")
  	HyftdUserId	= ""
		HyftdUserPassword = ""
	  HyftdGroupName = Application("HyftdGroupName")
	  HyftdAssociateDataName = Application("HyftdAssociateDataName")
	  HyftdConnId = 0
	  HyftdQueryId = 0
	  HyftdSortQueryId = 0
		HyftdQueryExId = 0
		HyftdQueryGroupingId = 0
	  HyftdAddQueryId = 0	  
	  HyftdTotalRecordCount = 0
	  HyftdCurrentRecordCount = 0 
	  HyftdMaxQueryCount = 10000
	  HyftdSortName = "" 
	  HyftdSortType = 0
	  HyftdDataId = 0
	  HyftdDataRelatedNumber = 0
	  HyftdDataLength = 0
	  HyftdPhonetic = CInt(Phonetic)
	  HyftdAuthorize = CInt(Authorize)
	  HyftdStartTime = GetTime()
  	HyftdEndTime = 0
	  HyftdTimeDiff = 0
	  HyftdGroupingCount = 0
	  HyftdGroupingName = "onlinedate"
		HyftdGroupingType = 1
		HyftdGroupingCount1 = 0
	  HyftdGroupingName1 = ""
		HyftdGroupingType1 = 1
		
		If PageNumber > 1 then
			HyftdRecordFrom = (PageNumber - 1) * PageSize
			HyftdRecordTo = (PageNumber - 1) * PageSize + PageSize
		Else
			HyftdRecordFrom = 0
			HyftdRecordTo = PageSize
		End if				
		'-------------------------------------------------------------------------------
		'---Query Type------------------------------------------------------------------
  	If Order = "0" Then		'---相關度---
	  	HyftdQueryType = 3	'---全文檢索(相關度分數出現次數*100)---
	  Else									'---其它---
    	HyftdQueryType = 0	'---QueryType=0,才能sortquery---
  	End If  	
  	'-------------------------------------------------------------------------------
		'---Hyftd Sort Name-------------------------------------------------------------
		If Order = "0" Then							'---關聯度---
	  	HyftdSortName = ""
	  ElseIf Order = "1" Then					'---標題---
    	HyftdSortName = "subject"
    ElseIf Order = "2" Then					'---發佈日期---
    	HyftdSortName = "onlinedate"
    ElseIf Order = "3" Then					'---資料來源---
    	HyftdSortName = "siteid"
    ElseIf Order = "4" Then					'---點閱次數---
    	HyftdSortName = "clickcount" 
  	End If
  	'-------------------------------------------------------------------------------
		'---Hyftd Sort Type-------------------------------------------------------------
		If Sort = "0" Then			'---遞減---
			HyftdSortType = 1
		ElseIf Sort = "1" Then	'---遞增---
			HyftdSortType = 0
		End If
			
	End Sub

	'-------------------------------------------------------------------------------
	'---Function Hyftd Initialization-----------------------------------------------
		
	Sub Hyftd_Set_Encoding(ft, Encoding)
	
		ft.hyft_setencoding(Encoding)
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Set Encoding : " & ft.errmsg
		  Response.End
		End If
		
	End Sub
	
	'-------------------------------------------------------------------------------
	'---Function Hyftd Connection---------------------------------------------------
		
	Function Hyftd_Connection( ft, ServerIP, Port, GroupName, UserId, UserPassword )	
		HyftdX = ft.hyft_connect( ServerIP, Port, GroupName, UserId, UserPassword )
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Connection : " & ft.errmsg
		  Response.End
		End If	
		Hyftd_Connection = ft.ret_val			 
		
	End Function

	'-------------------------------------------------------------------------------
	'---Function Hyftd Initial Query------------------------------------------------
	
	Sub Hyftd_Initial_Query( ft, ConnId )
	
		HyftdX = ft.hyft_initquery( ConnId )	
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Initial Query : " & ft.errmsg
		  Response.End
		End If
		
	End Sub

	'-------------------------------------------------------------------------------
	'---Function Hyftd Add Query----------------------------------------------------
	
	Sub Hyftd_Add_Query( ft, ConnId, ColumnName, QueryString1, QueryString2, QueryType, Phonetic, Authorize )
	
		HyftdX = ft.hyft_addquery( ConnId, ColumnName, QueryString1, QueryString2, QueryType, Phonetic, Authorize )
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Add Query : " & ft.errmsg
		  Response.End
		End If
		
	End Sub
	
	Sub Hyftd_Add_And( ft, ConnId )
		
		HyftdX = ft.hyft_addand(ConnId)
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Add And : " & ft.errmsg
		  Response.End
		End If
		
	End Sub

	Sub Hyftd_Add_Or( ft, ConnId )
		
		HyftdX = ft.hyft_addor(ConnId)
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Add Or : " & ft.errmsg
		  Response.End
		End If
		
	End Sub
	
	Sub Hyftd_Add_Not( ft, ConnId )
		
		HyftdX = ft.hyft_addnot(ConnId)
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Add Not : " & ft.errmsg
		  Response.End
		End If
		
	End Sub

	Function Hyftd_Query_Ex( ft, ConnId, SortName, SortType, MaxQueryCount, RecordFrom, RecordTo, GroupName, GroupType)		
	
		HyftdX = ft.hyft_query_ex( ConnId, SortName, SortType, MaxQueryCount, RecordFrom, RecordTo, GroupName, GroupType )		
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Query Ex : " & ft.errmsg
		  Response.End
		End If
		Hyftd_Query_Ex = ft.ret_val
		
	End Function 
	
	Function Hyftd_Sort_Query( ft, ConnId, SortName, SortType, RecordFrom, RecordTo )	
	
		HyftdX = ft.hyft_sortquery( ConnId, SortName, SortType, RecordFrom, RecordTo )	
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Sort Query : " & ft.errmsg
		  Response.End
		End If
		Hyftd_Sort_Query = ft.ret_val
		
	End Function
	
	Function Hyftd_Total_Sysid( ft, ConnId, QueryId )
	
		HyftdX = ft.hyft_total_sysid( ConnId, QueryId )
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Total Sysid : " & ft.errmsg
		  Response.End
		End If
		Hyftd_Total_Sysid = ft.ret_val
		
	End Function

	Function Hyftd_Num_Sysid( ft, ConnId, QueryId )
	
		HyftdX = ft.hyft_num_sysid( ConnId, QueryId )
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Num Sysid : " & ft.errmsg
		  Response.End
		End If
		Hyftd_Num_Sysid = ft.ret_val
		
	End Function
	
	Function Hyftd_Fetch_Sysid( ft, ConnId, QueryId, i )
		
		HyftdX = ft.hyft_fetch_sysid( ConnId, QueryId, i )
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Fetch Sysid : " & ft.errmsg
		  Response.End
		End If
		Hyftd_Fetch_Sysid = ft.sysid
		
	End Function
	
	Function Hyftd_Get_Related_Number( ft )
	
		Hyftd_Get_Related_Number = ft.ret_val
		
	End Function
	
	Function Hyftd_One_Assdata_Len( ft, ConnId, AssociateDataName, DataId )
	
		HyftdX = ft.hyft_one_assdata_len( ConnId, AssociateDataName, DataId )
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "One Assdata Len : " & ft.errmsg
		  Response.End
		End If
		Hyftd_One_Assdata_Len = ft.ret_val
		
	End Function
	
	Function Hyftd_One_Assdata( ft, ConnId, DataLength )
	
		HyftdX = ft.hyft_one_assdata( ConnId, DataLength )
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "One Assdata : " & ft.errmsg
		  Response.End
		End If
		Hyftd_One_Assdata = ft.data
		
	End Function
	
	Function Hyftd_Query_Grouping( ft, ConnId, GroupingName, GroupingLen, GroupingType )
	
		If GroupingLen <> "0" Then
			HyftdX = ft.hyft_query_grouping( ConnId, GroupingName & ":" & GroupingLen, GroupingType )
		Else
			HyftdX = ft.hyft_query_grouping( ConnId, GroupingName, GroupingType )
		End If
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Query Grouping : " & ft.errmsg
		  Response.End
		End If
		Hyftd_Query_Grouping = ft.ret_val
		
	End Function
	
	Function Hyftd_Num_Grouping( ft, ConnId, QueryId )
		
		HyftdX = ft.hyft_num_grouping( ConnId, QueryId )
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Num Grouping : " & ft.errmsg
		  Response.End
		End If
		Hyftd_Num_Grouping = ft.ret_val
		
	End Function
	
	Function Hyftd_Fetch_Grouping( ft, ConnId, QueryId, i )
	
		HyftdX = ft.hyft_fetch_grouping( ConnId, QueryId, i )
		If ft.ret_val = -1 Then 'And HyftdDebug = "true" Then
			Response.Write "Fetch Grouping : " & ft.errmsg
		  Response.End
		End If
		Hyftd_Fetch_Grouping = ft.ret_val
	
	End Function
	
	Function Hyftd_Get_Grouping_Name( ft )
	 
		Hyftd_Get_Grouping_Name = ft.data
		
	End Function

	Function Hyftd_GetKeys( ft, ConnId, QueryString )
	
		HyftdX = ft.hyft_getkeys( ConnId, QueryString )
		If ft.ret_val < 0 Then 
			Response.Write "Get Keys : " & ft.errmsg
		  Response.End
		End If
		key_id = ft.ret_val
		HyftdX = ft.hyft_num_keys( ConnId, key_id )
		if ft.ret_val < 0 then
			Response.Write "Num Keys : " & ft.errmsg
		  Response.End
		end if
		key_num = ft.ret_val
		str = ""
		for i = 0 to key_num - 1
			HyftdX = ft.hyft_keylen( ConnId, key_id, i )
			if ft.ret_val < 0 then
				Response.Write "keys len : " & ft.errmsg
				Response.End
			end if
			l = ft.ret_val
			HyftdX = ft.hyft_onekey( ConnId, key_id, i ,l)
			if ft.ret_val < 0 then
				Response.Write "onekey : " & ft.errmsg
				Response.End
			end if
			str = str & ft.key & ";"
		next
		HyftdX = ft.hyft_free_key( ConnId, key_id )
		Hyftd_GetKeys = str
		
	End Function
	
	Function GetFieldIndexString(ft, ConnId, QueryId, fieldName)

	 	HyftdX = ft.hyft_get_fieldindexstring(ConnId, QueryId, fieldName)
		if ft.ret_val < 0 then
			response.write "hyft_get_fieldindexstring: " & ft.errmsg
			response.end
		end if
		GetFieldIndexString ft.data
	End Function

	
	Function CompareDate( ty, tm, td )
		
		dateNow = Date()
		yearNow = DatePart("yyyy", dateNow)
		monthNow = DatePart("m", dateNow)
		dayNow = DatePart("d", dateNow)
		Dim result
		If CInt(ty) > yearNow Then
			ty = CStr(yearNow)
			tm = CStr(monthNow)
			td = CStr(dayNow)
		ElseIf CInt(ty) = yearNow Then
			If CInt(tm) > monthNow Then
				tm = CStr(monthNow)
				td = CStr(dayNow)
			ElseIf CInt(tm) = monthNow Then
				If CInt(td) > dayNow Then
					td = CStr(dayNow)
				End If				
			End If
		End If
		If Len(tm) = 1 Then
			tm = "0" & tm
		End If
		If Len(td) = 1 Then
			td = "0" & td
		End If
		result = ty & tm & td
		CompareDate = result
	
	End Function
	
%>