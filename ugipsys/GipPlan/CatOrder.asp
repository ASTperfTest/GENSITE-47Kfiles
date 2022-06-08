<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "Pn03M04" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
	CatID = Request.QueryString("CatID")
	ItemCount = Request.QueryString("ItemCount")

     	 SQLCom = "Select Count(CatID) from Catalogue Where DataParent = "& CatID
	 	 Set RSCom = Conn.execute(SQLCom)
	 	  ItemCount = rscom(0)
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>分類排序</title>
<link rel="stylesheet" href="../inc/setstyle.css">
</head>
<body>
<center>
<form method="POST" name="Menu">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td class="FormName">討論區管理 - 分類排序</td>
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
         <form method="POST" action="CatConfig.asp?ItemID=<%=ItemID%>&CatID=<%=CatID%>&ConfigCk=Order" name="Orderreg">
		  <input type="hidden" name="DataParent" value="<%=CatID%>">
          <table border="0" cellspacing="0" cellpadding="3" class="c12-3">
           <tr><td align="center">
    <% 	 SQLCom = "Select * from Catalogue Where DataParent = "& CatID &" Order by CatShowOrder"
	 	 Set rs2 = Conn.execute(SQLCom)
	        If Not rs2.EOF Then %>
           <select size="<%=ItemCount%>" name="CatID" multiple>
	      <% Do while not rs2.eof %>
          <option value="<%=rs2("CatID")%>"><%=rs2("CatName")%></option>
	    <%	 rs2.movenext
	    	 loop %>
              </select>
     <% End If %>
              </td><td align="center" valign="middle">
    <img border="0" src="image/up01.gif" OnClick="VBS: ChgOrderUp 'CatID'" style="cursor:hand;">
    <p><img border="0" src="image/down01.gif" OnClick="VBS: ChgOrderDown 'CatID'" style="cursor:hand;">
              </td></tr>
           <tr><td align="right" colspan="2"><input type="button" value="確定" class="cbutton" OnClick="VBS: FormSubmit()">
              </td></tr>
          </table>
     	 </form>

</center>
</body></html>
<script language=VBScript>
 Sub ChgOrderUp(FormObjName)
 	SET xn = document.OrderReg(FormObjName)
 	OldOpt = xn.Options.Selectedindex

  If OldOpt <> -1 Then
   If OldOpt+1 > 1 Then
 	OldText = xn.Options(OldOpt).text
 	OldValue = xn.Options(OldOpt).value
 	OldText2 = xn.Options(OldOpt-1).text
 	OldValue2 = xn.Options(OldOpt-1).value

 	 xn.Options(OldOpt).text = xn.Options(OldOpt-1).text
     xn.Options(OldOpt).value = xn.Options(OldOpt-1).value
 	 xn.Options(OldOpt-1).text = OldText
 	 xn.Options(OldOpt-1).value = OldValue
 	 xn.Options(OldOpt-1).selected = True
 	 xn.Options(OldOpt).selected = False
   Else
 	 alert ("已經是最上方了！")
   End If
  Else
 	 alert ("請先選擇..")
  End IF
 End Sub

 Sub ChgOrderDown(FormObjName)
 	SET xn = document.OrderReg(FormObjName)
 	rno = xn.Options.length
 	OldOpt = xn.Options.Selectedindex

  If OldOpt <> -1 Then
 	If OldOpt+1 < rno Then
 	 OldText = xn.Options(OldOpt).text
 	 OldValue = xn.Options(OldOpt).value
 	 OldText2 = xn.Options(OldOpt+1).text
 	 OldValue2 = xn.Options(OldOpt+1).value

      xn.Options(OldOpt).text = xn.Options(OldOpt+1).text
 	  xn.Options(OldOpt).value = xn.Options(OldOpt+1).value
 	  xn.Options(OldOpt+1).text = OldText
 	  xn.Options(OldOpt+1).value = OldValue
 	  xn.Options(OldOpt+1).selected = True
 	  xn.Options(OldOpt).selected = False
 	Else
 	 alert ("已經是最下方了！")
 	End If
  Else
 	 alert ("請先選擇..")
  End If
 End Sub

	Sub Document_OnKeyDown()
	 If window.event.KeyCode = 16 or window.event.KeyCode = 17 Then
	  document.all.OrderReg.CatID.multiple = False
	 End IF
	End Sub

	Sub Document_OnKeyUp()
	  document.all.OrderReg.CatID.multiple = True
	End Sub

	Sub FormSubmit()
	 Optpos = OrderReg.CatID.options.length
	 for OptCount = 0 to Optpos - 1
	   document.all.OrderReg.CatID.options(OptCount).selected = True
	 next
	 OrderReg.submit
	End Sub
</script>
