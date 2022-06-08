<% Response.Expires = 0
	DataType = Request.QueryString("DataType")
	Language = Request.QueryString("Language")

	Select Case DataType
	 Case "M01"
	 '內部公告
	  HTProgCode = "Pn01M01"
	  CatDecide = "N"
	  DateDecide = "Y"
	  EMailDecide = "Y"
	  Subject = "主旨"
	  Content = "內容"
	  Extend_1 = Null
	  Extend_2 = Null
	  Extend_3 = Null
	 Case "M02"
	 '外部公告
	  HTProgCode = "Pn01M02"
	  CatDecide = "N"
	  DateDecide = "Y"
	  EMailDecide = "N"
	  Subject = "主旨"
	  Content = "內容"
	  Extend_1 = Null
	  Extend_2 = Null
	  Extend_3 = Null
	End Select

%>
