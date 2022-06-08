<%@ CodePage = 65001 %>
<%
	'-------------------------------------------------------------------------------
	'---Function Hyftd Initialization-----------------------------------------------	
	Set HyftdObj = Server.CreateObject("hysdk.hyft.1")
 	HyftdObj.hyft_setencoding("big5")
	If HyftdObj.ret_val = -1 Then
		Response.Write "Set Encoding : " & HyftdObj.errmsg
		Response.End
	End If	
	HyftdX = HyftdObj.hyft_connect( "127.0.0.1", 2816, "Intra", "", "" )
	If HyftdObj.ret_val = -1 Then 
		Response.Write "Connection : " & HyftdObj.errmsg
	  Response.End
	End If	
	Hyftd_Connection = HyftdObj.ret_val	

	
	HyftdX = HyftdObj.hyft_initquery( Hyftd_Connection )	
	If HyftdObj.ret_val = -1 Then 
		Response.Write "Initial Query : " & HyftdObj.errmsg
	  Response.End
	End If
	
	HyftdX = HyftdObj.hyft_addquery( Hyftd_Connection, "associatedata", "水稻", "", 0, 0, 0 )
	If HyftdObj.ret_val = -1 Then 
		Response.Write "Add Query : " & HyftdObj.errmsg
	  Response.End
	End If
	'HyftdX = HyftdObj.hyft_addquery( Hyftd_Connection, "actorinfoid", "001", "", 0, 0, 0 )
	'If HyftdObj.ret_val = -1 Then 
	'	Response.Write "Add Query : " & HyftdObj.errmsg
	'  Response.End
	'End If	
	'HyftdX = HyftdObj.hyft_addand(Hyftd_Connection)
	'If HyftdObj.ret_val = -1 Then
	'	Response.Write "Add And : " & HyftdObj.errmsg
	'  Response.End
	'End If
	
	HyftdX = HyftdObj.hyft_sortquery( Hyftd_Connection, "subject", 0, 0, 500 )	
	If HyftdObj.ret_val = -1 Then 
		Response.Write "Sort Query : " & HyftdObj.errmsg
	  Response.End
	End If
	HyftdSortQueryId= HyftdObj.ret_val		
	'-------------------------------------------------------------------------------
	'---Get Total Record Count------------------------------------------------------	
	HyftdX = HyftdObj.hyft_total_sysid( Hyftd_Connection, HyftdSortQueryId )
	If HyftdObj.ret_val = -1 Then
		Response.Write "Total Sysid : " & HyftdObj.errmsg
	  Response.End
	End If
	HyftdTotalRecordCount = HyftdObj.ret_val
	response.write "HyftdTotalRecordCount : " & HyftdTotalRecordCount & "<hr>"
	'-------------------------------------------------------------------------------
	'---Get Current Record Count----------------------------------------------------	
	HyftdX = HyftdObj.hyft_num_sysid( Hyftd_Connection, HyftdSortQueryId )
	If HyftdObj.ret_val = -1 Then 
		Response.Write "Num Sysid : " & HyftdObj.errmsg
	  Response.End
	End If
	HyftdCurrentRecordCount = HyftdObj.ret_val
	
	If HyftdCurrentRecordCount > 0 Then		
		Dim index					
		For index = 0 To HyftdCurrentRecordCount - 1													
			'---------------------------------------------------------------------------
			'---Fetch Current Sysid-----------------------------------------------------			
			HyftdX = HyftdObj.hyft_fetch_sysid( Hyftd_Connection, HyftdSortQueryId, index )
			If HyftdObj.ret_val = -1 Then 
				Response.Write "Fetch Sysid : " & HyftdObj.errmsg
		  	Response.End
			End If
			HyftdDataId = HyftdObj.sysid
			response.write "HyftdDataId : " & HyftdDataId & " ~ "
      '---------------------------------------------------------------------------
      '---Fetch Associated Data---------------------------------------------------                              
      HyftdX = HyftdObj.hyft_one_assdata_len( Hyftd_Connection, "brief", HyftdDataId )
			If HyftdObj.ret_val = -1 Then 
				Response.Write "One Assdata Len : " & HyftdObj.errmsg
		  	Response.End
			End If
			HyftdAssDataLen = HyftdObj.ret_val
			HyftdX = HyftdObj.hyft_one_assdata( Hyftd_Connection, HyftdAssDataLen )
			If HyftdObj.ret_val = -1 Then 
				Response.Write "One Assdata : " & HyftdObj.errmsg
		  	Response.End
			End If
			Hyftd_One_Assdata = HyftdObj.data					
      response.write "HyftdAssData : " & Hyftd_One_Assdata & "<hr>"
	Next					
	End If  	
	'---get grouping---	
	HyftdX = HyftdObj.hyft_query_grouping( Hyftd_Connection, "actorinfoid", 0 )
	If HyftdObj.ret_val = -1 Then 
		Response.Write "Query Grouping : " & HyftdObj.errmsg
	  Response.End
	End If
	HyftdQueryGroupingId = HyftdObj.ret_val
	
	'---get how many group---	
	HyftdX = HyftdObj.hyft_num_grouping( Hyftd_Connection, HyftdQueryGroupingId )
	If HyftdObj.ret_val = -1 Then
		Response.Write "Num Grouping : " & HyftdObj.errmsg
	  Response.End
	End If
	HyftdGroupingCount1 = HyftdObj.ret_val
	If HyftdGroupingCount1 > 0 Then
			For index = 0 To HyftdGroupingCount1 - 1
				'---get the name and number of grouping---
				HyftdX = HyftdObj.hyft_fetch_grouping( Hyftd_Connection, HyftdQueryGroupingId, index )
				If HyftdObj.ret_val = -1 Then 
					Response.Write "Fetch Grouping : " & HyftdObj.errmsg
		  		Response.End
				End If
				Hyftd_Fetch_Grouping = HyftdObj.ret_val					
				response.write "Hyftd_Fetch_Grouping : " &  Hyftd_Fetch_Grouping & " ~ "				
				response.write "Hyftd_Get_Grouping_Name : " &  HyftdObj.data	& "<hr>"
			Next
	End If
	HyftdX = HyftdObj.hyft_close(Hyftd_Connection)    
	HyftdObj = Empty
 	response.write err.description & " at: " & err.number		
%>