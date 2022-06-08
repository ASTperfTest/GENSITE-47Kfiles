<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
	CatID = Request.QueryString("CatID")

	 SQL = "Select * from CatTreeNode Where CtNodeID = "& CatID
	 Set RS = Conn.execute(SQL)
	  If rs.EOF Then
	   DataTop = 0
	   DataParent = 0
	   DataLevel = 1
	  Else
	   DataTop = 0
	   DataParent = CatID
	   DataLevel = rs("DataLevel") + 1
	  End IF
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>新增分類</title>
<link rel="stylesheet" href="../inc/setstyle.css">
</head>
<body>
<center>
<form method="POST" name="Menu">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td class="FormName">分類樹管理 - 新增分類</td>
    <td align="right"><input type="button" value="清單" class="cbutton" OnClick="VBS: history.back"></td>
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
<form method="POST" action="CatConfig.asp?ItemID=<%=ItemID%>&ConfigCk=Add" name="AddCat">
<input type="hidden" name="DataTop" value="<%=DataTop%>">
<input type="hidden" name="DataParent" value="<%=DataParent%>">
<input type="hidden" name="DataLevel" value="<%=DataLevel%>">
  <table border="0" class="CTable" cellspacing="1" cellpadding="3" width="300" height="150">
    <tr>
      <td class="TableHeadTD">新增分類</td>
    </tr>
    <tr>
      <td class="TableBodyTD" align="left" height="135">分類名稱：<input type="text" name="CatName" size="30">
      <p align="left">
      顯示次序：<input type="text" name="CatShowOrder" size="3">
        <p align="right"><input type="button" value="確定" class="cbutton" OnClick="FormSubmit()"><input type="button" value="重填" OnClick="VBS: AddCat.reset" class="cbutton"></p>
      </td>
    </tr>
  </table>
</form>
</center>
</body></html>
<script language="VBScript">
 Sub FormSubmit()
  If trim(AddCat.CatName.value) = Empty Then
     MsgBox "請輸入分類名稱！", 16, "Sorry!"
     AddCat.CatName.focus
     Exit Sub
  ElseIf blen(AddCat.CatName.value) > 100 Then
     MsgBox "分類名稱字數過多！", 16, "Sorry!"
     AddCat.CatName.focus
     Exit Sub
  End IF
  AddCat.Submit
 End Sub

 function blen(xs)
 xl = len(xs)
 for i=1 to len(xs)
  if asc(mid(xs,i,1))<0 then xl = xl + 1
 next
 blen = xl
end function

</script>