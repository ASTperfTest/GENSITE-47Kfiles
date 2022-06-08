<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
    <% SQL = "SELECT * FROM DataCat Where Language = '"& Language &"' And DataType = '"& DataType &"' And CatName = "& NoHTMLCode(Request.QueryString("CatName")) 
       set rs = conn.execute(SQL) 
        If Not rs.EOF Then %>
        <script language=VBS>
          alert "類別名稱重覆，請重新新增！"
		  window.parent.mainFrame.document.all.xin_CatID.options(0).selected = True        
		  window.location.href = "../leftmenu2.htm"         
        </script>
<%    response.end
     End if

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

    SQL = "INSERT INTO DataCat (Language, DataType, CatName, EditUserID, EditDate, CatShowOrder) VALUES (N'"& Language &"',N'"& DataType &"',"& NoHTMLCode(Request.QueryString("CatName")) &",N'"& Session("UserID") &"',N'"& Date() &"',1)"
    conn.execute(SQL) 
    
	SQL = "SELECT CatID FROM DataCat Where Language = N'"& Language &"' And DataType = N'"& DataType &"' And CatName = "& NoHTMLCode(Request.QueryString("CatName"))
	set rs = conn.execute(SQL) %>

<script language="VBscript">
	window.parent.mainFrame.document.all("xID").value=<%=rs("CatID")%>         
	window.parent.mainFrame.document.all("AddName").value="<%=Request.QueryString("CatName")%>"         
	window.parent.mainFrame.SelectAdd 
	window.location.href = "../leftmenu2.htm"         
</script>
<%
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

