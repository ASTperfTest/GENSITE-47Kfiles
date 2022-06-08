<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
	CatID = Request.Form("CatID") 
	xType = Request.QueryString("xType")
	CatName = Request.Form("CatName")
	
	If xType = "Add" Then
       SQL = "SELECT * FROM DataCat Where Language = N'"& Language &"' And DataType = N'"& DataType &"' And CatName = "& NoHTMLCode(CatName) 
       set rs = conn.execute(SQL) 
        If Not rs.EOF Then %>
        <script language=VBS>
          alert "類別名稱重覆，請重新新增！"
		  window.location.href = "CatList.asp?Language=<%=Language%>&DataType=<%=DataType%>"         
        </script>
<%      Else
	    SQL1 = "Select * From DataCat Where Language = N'"& Language &"' And DataType = N'"& DataType &"' Order By CatShowOrder"
	    set rs1 = conn.Execute(SQL1)
		If not rs1.EOF Then
	      ItemOrder = 1
		  do while not rs1.eof
		    ItemOrder = ItemOrder + 1
		    sql2 = "update DataCat set CatShowOrder =" & ItemOrder & " where CatID =" & rs1("CatID")
		    set rs2 = conn.execute(sql2)
		  rs1.movenext
		  loop
		End if
	     SQL = "INSERT INTO DataCat (Language, DataType, CatName, EditUserID, EditDate,CatShowOrder) VALUES (N'"& Language &"',N'"& DataType &"',"& NoHTMLCode(CatName) &",N'"& Session("UserID") &"',N'"& Date() &"',1)"
	     conn.execute(SQL)
	     msg = "新增完成！" 
        End If
	ElseIf xType = "Edit" Then
       SQL = "SELECT * FROM DataCat Where Language = N'"& Language &"' And DataType = N'"& DataType &"' And CatName = "& NoHTMLCode(CatName) &" And CatID <> "& CatID
       set rs = conn.execute(SQL) 
        If Not rs.EOF Then %>
        <script language=VBS>
          alert "類別名稱重覆，請重新編修！"
		  window.location.href = "CatList.asp?Language=<%=Language%>&DataType=<%=DataType%>"         
        </script>
<%      Else
			SQL = "Update DataCat Set CatName = "& NoHTMLCode(CatName) &" Where CatID ="& CatID
			set rs = conn.Execute(SQL)
			msg = "編修完成！"
        End If
	ElseIf xType = "Del" Then
		SQL = "Delete from DataCat Where CatID ="& CatID
		set rs = conn.Execute(SQL)

	    SQL1 = "Select * From DataCat Where Language = N'"& Language &"' And DataType = N'"& DataType &"' Order By CatShowOrder"
	    set rs1 = conn.Execute(SQL1)
		If not rs1.EOF Then
	      ItemOrder = 0
		  do while not rs1.eof
		    ItemOrder = ItemOrder + 1
		    sql2 = "update DataCat set CatShowOrder =" & ItemOrder & " where CatID =" & rs1("CatID")
		    set rs2 = conn.execute(sql2)
		  rs1.movenext
		  loop
		End if
		msg = "刪除完成！"
	ElseIf xType = "OrderBy" Then
	 CatIDArrar = Split(CatID,",")
	 NowShowOrder = 0
	  For xno = 0 to UBound(CatIDArrar)
	   NowShowOrder = NowShowOrder + 1
		sql2 = "Update DataCat set CatShowOrder =" & NowShowOrder & " where CatID ="& CatIDArrar(xno)
		set rs2 = conn.execute(sql2)
	  Next
	  msg = "排序完成！"
	End If

Function NoHTMLCode(DataCode)
  NewData = "" 
  If DataCode <> "" Then
   NewDate = Replace(DataCode,"'","''")
   NewDate = Replace(NewDate,"<","&lt;")
   NewDate = Replace(NewDate,">","&gt;")
   NoHTMLCode = "'" & NewDate & "'"
  End If
End Function

%>
<script language=VBScript>
  alert("<%=msg%>")
  document.location.href="CatList.asp?Language=<%=Language%>&DataType=<%=DataType%>"
</script>

