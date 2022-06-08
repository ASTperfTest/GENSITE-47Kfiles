<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
	 CatID = Request.QueryString("CatID")
	 If CatID <> 0 Then
	  SQL = "Select * from CatTreeNode Where ctNodeId = "& CatID
	  Set RS = Conn.execute(SQL)
	  CatName = rs("CatName")
	  TitleName = "目錄名稱："
	  DelCk = Request.QueryString("DelCk")
		SQLCom = "Select * from CatTreeNode Where ctNodeKind='U' AND dataParent="& CatID &" Order by catShowOrder"
		Set RSCom = Conn.execute(SQLCom)
	 Else
	  CatName = "根目錄"
	  TitleName = ""
	  DelCk = "N"
		SQLCom = "Select * from CatTreeNode Where ctNodeKind='U' AND dataParent="& 0 &" Order by catShowOrder"
		Set RSCom = Conn.execute(SQLCom)
	 End IF

%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>檢視目錄</title>
<link rel="stylesheet" href="../inc/setstyle.css">
</head>
<body>
<center>
<form method="POST" name="Menu">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td class="FormName">目錄樹管理 - 檢視目錄</td>
    <td align="right">
    <% If CatID <> 0 Then %><input type="button" value="編修" class="cbutton" OnClick="FormSubmit()"><% End IF %>
    <input type="button" value="新增項目" class="cbutton" OnClick="AddForum()">
     <% If CatID<>0 Then %><input type="button" value="清單" class="cbutton" OnClick="VBS: document.location.href='CatList.asp?ItemID=<%=ItemID%>&CatID=<%=rs("DataParent")%>'"><% End IF %>
    </td>
  </tr>
  <tr>
    <td colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td colspan="2" class="Formtext"></td>
  </tr>
</table>
</form>
  <table border="0" class="CTable" cellspacing="1" cellpadding="3" width="500">
    <tr>
      <td class="TableHeadTD" align="right" width="30%"><%=TitleName%></td>
      <td class="TableBodyTD"><%=CatName%></td>
    </tr>
    <tr>
      <td class="TableHeadTD" align="right" valign="top">下屬單元：</td>
      <td class="TableBodyTD">
     <% If Not RSCom.EOF Then %>
        <table border="0" width="100%" cellspacing="0" cellpadding="3" class="TableBodyTD">
        <% Do while not RSCom.EOF %>
          <tr>
            <td width="100%"><%=RSCom("CatName")%></td>
          </tr>
        <% RSCom.movenext
	       Loop %>
        </table>
       <% Else %>
        無單元
       <% End IF %>
      </td>
    </tr>
  </table>
</center>
</body></html>
<script language="VBScript">
 Sub FormSubmit()
' 	window.location.href = "CatEdit.asp?ItemID=<%=ItemID%>&DelCk=<%=DelCk%>&CatID=<%=CatID%>&ConfigCk=Edit"
 	window.location.href = "CtUnitNodeEdit.asp?ctNodeId=<%=CatID%>&phase=edit"
 End Sub

 Sub xFormSubmit()
 	window.location.href = "ForumOrder.asp?ItemID=<%=ItemID%>&CatID=<%=CatID%>&ConfigCk=Order"
 End Sub

 function blen(xs)
 xl = len(xs)
 for i=1 to len(xs)
  if asc(mid(xs,i,1))<0 then xl = xl + 1
 next
 blen = xl
end function

	Sub AddForum()
	 window.location.href = "CtUnitNodeAdd.asp?ItemID=<%=ItemID%>&CatID=<%=CatID%>&ctNodeKind=U&phase=add"
	End Sub

</script>