<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
	ItemCount = 0

	SQL = "SELECT * FROM DataCat Where Language = N'"& Language &"' And DataType = N'"& DataType &"' Order By CatShowOrder"
    set rs = conn.execute(SQL)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<title>分類編修</title>
</head>
<body>
<center>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=Title%>&nbsp;<font size=2>【分類編修】</font></td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
          <% if (HTProgRight and 8)=8 And Not rs.EOF then %><SPAN OnClick="FormOrderOpen()" style="cursor:hand;">分類排序</SPAN><% End IF %>
          <% if (HTProgRight and 4)=4 then %>　<SPAN OnClick="FormAddOpen()" style="cursor:hand;">新增分類</SPAN><% End IF %>
          <% if (HTProgRight and 2)=2 then %>　<a href="BulletinList.asp?Language=<%=Language%>&DataType=<%=DataType%>">清單</a><% End IF %>
		</td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15">編修分類，請點選分類名稱。</td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    

<br>
          <% If rs.EOF Then %>
          <font color="#FF0000" size="2">
            查無分類資料</font>
		  <% Else %>
            <table border="0" width="95%" cellspacing="1" cellpadding="3" class="bluetable">
              <tr align="center" class="lightbluetable">
                <td width="85%">分　類　名　稱</td>
                <td width="15%"></td>
              </tr>
             <% Do While Not Rs.EOF
              	 ItemCount = ItemCount + 1
                 SQLCom = "Select count(*) From DataUnit Where CatID = "& rs("CatID")
				 set rscom = conn.Execute(SQLCom) %>
              <tr class="whitetablebg">
                <td width="85%" style="cursor:hand;" OnClick="VBS: FormEditOpen <%=rs("CatID")%>,'<%=rs("CatName")%>',<%=rscom(0)%>"><%=rs("CatName")%></td>
                <td width="15%" align="center"><font color="#FF0000"><%=rscom(0)%></font></td>
              </tr>
              <% rs.MoveNext
                 Loop %>
            </table>
          <% End If %>
          
    </td>
  </tr>  
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      Best viewed with Internet Explorer 5.5+ , Screen Size set to 800x600.</td>                                            
  </tr> 
</table> 

</center>

<table border="0" width="100%" cellspacing="1" cellpadding="3" id="SetView" class="bluetable" style="position:absolute;top:180px;left:255px; width: 300px; height: 146px; visibility: hidden">
  <tr class="lightbluetable">
    <td>
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr class="lightbluetable">
          <td ID="EditTitle"></td>
          <td align="right"><input type="button" value="Ｘ" class="cbutton" onClick="VBS: SetView.style.visibility = 'hidden'"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr class="whitetablebg">
    <td align="center">

         <form method="POST" action="" name="reg">
          <input type="hidden" name="CatID" value="">
          <table border="0" cellspacing="0" cellpadding="0" class="whitetablebg">
           <tr><td align="center">名稱: &nbsp;<input type="text" name="CatName" size="20" value=""></td></tr>  
          </table><br>
          <table border="0" width="95%" cellspacing="0" cellpadding="0">
           <tr><td align="right"><% if (HTProgRight and 4)=4 then %><input type="button" value="確定" class="cbutton" onClick="Upedit()"><% End If %><% if (HTProgRight and 16)=16 then %><input type="button" name="EditDelBT" value="刪除" class="cbutton" onClick="Del()"><% End If %><input type="button" value="取消" class="cbutton" onClick="WindowClose()"></td></tr>
          </table>
     	 </form>

    </td>
  </tr>
</table>


<table border="0" width="100%" cellspacing="1" cellpadding="3" id="SetOrderView" class="bluetable" style="position:absolute;top:180px;left:255px; width: 300px; height: 146px; visibility: hidden">
  <tr class="lightbluetable">
    <td>
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr class="lightbluetable">
          <td>排序</td>
          <td align="right"><input type="button" value="Ｘ" class="cbutton" onClick="WindowClose()"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr class="whitetablebg">
    <td align="center">

         <form method="POST" action="CatConfig.asp?Language=<%=Language%>&DataType=<%=DataType%>&xType=OrderBy" name="Orderreg">
          <table border="0" cellspacing="0" cellpadding="3" class="whitetablebg">
           <tr><td align="center">
    <% SQLCom = "SELECT CatID, CatName FROM DataCat Where Language = N'"& Language &"' And DataType = N'"& DataType &"' Order By CatShowOrder"
       set rs2 = conn.execute(SQLCom)
        If Not rs2.EOF Then %>
           <select size="<%=ItemCount%>" name="CatID" multiple>
      <% Do while not rs2.eof %>
          <option value="<%=rs2("CatID")%>"><%=rs2("CatName")%></option>
    <%	 rs2.movenext
    	 loop %>
              </select>
     <% End If %>
              </td>
              <td align="center" valign="middle"><input type="button" value="↑" class="cbutton" OnClick="VBS: ChgOrderUp 'CatID'">
										      <p><input type="button" value="↓" class="cbutton" OnClick="VBS: ChgOrderDown 'CatID'">
              </td></tr>
          </table><br>
          <table border="0" width="95%" cellspacing="0" cellpadding="0">
           <tr><td align="right"><input type="button" value="確定" class="cbutton" OnClick="VBS: FormSubmit()"></td></tr>
          </table>
     	 </form>

    </td>
  </tr>
</table>

</body>
</html>
<script language=VBScript>

xType = ""

Sub FormOrderOpen()
  DivTop = document.body.scrolltop + window.event.Y + 100
  document.all.SetOrderView.style.top = Divtop
  xType = "OrderBy"
  document.all.SetView.style.visibility = "hidden"
  reg.reset
  document.all.SetOrderView.style.visibility = ""
<% if (HTProgRight and 16)=16 then %>
  document.all("EditDelBT").style.display="none"
<% End IF %>
End Sub

Sub FormEditOpen(d,n,v)
  DivTop = document.body.scrolltop + window.event.Y + 50
  document.all.SetView.style.top = Divtop
  document.all.SetView.style.visibility = ""
  document.all.SetOrderView.style.visibility = "hidden"
  reg.reset
  document.reg.CatID.value = d
  document.reg.CatName.value = n
  xType = "Edit"
  document.all.EditTitle.InnerText = "編修分類"
<% if (HTProgRight and 16)=16 then %>
     if v = 0 then
      document.all("EditDelBT").style.display=""
     Else
      document.all("EditDelBT").style.display="none"
     end IF
<% End IF %>
End Sub

Sub FormAddOpen()
  DivTop = document.body.scrolltop + window.event.Y + 100
  document.all.SetView.style.top = Divtop
  xType = "Add"
  document.all.SetView.style.visibility = ""
  document.all.SetOrderView.style.visibility = "hidden"
<% if (HTProgRight and 16)=16 then %>
  document.all("EditDelBT").style.display="none"
<% End IF %>
  document.all.EditTitle.InnerText = "新增分類"
  reg.reset
End Sub

Sub WindowClose()
  document.all.SetView.style.visibility = "hidden"
  document.all.SetOrderView.style.visibility = "hidden"
<% if (HTProgRight and 16)=16 then %>
  document.all("EditDelBT").style.display="none"
<% End IF %>
  xType = ""
  reg.reset
End Sub

Sub upedit()
  msg1 = "請輸入分類名稱！"
  If trim(reg.CatName.value) = Empty Then
     MsgBox msg1, 64, "Sorry!"
     reg.CatName.focus
     Exit Sub
  End if

<% If Language = "C" Then %>
  msg = "抱歉!"& vbcrlf & vbcrlf &"類別名稱字數過多，請勿超過10個中文字。"& vbcrlf & vbcrlf &"請重新輸入 !"
<% ElseIf Language = "E" Then %>
  msg = "抱歉!"& vbcrlf & vbcrlf &"類別名稱字數過多，請勿超過20個字元。"& vbcrlf & vbcrlf &"請重新輸入 !"
<% End IF %>
  If blen(reg.CatName.value) > 20 Then
   MsgBox msg, 64, "Sorry!"
   reg.CatName.focus
   Exit Sub
  End IF

  reg.action="CatConfig.asp?Language=<%=Language%>&DataType=<%=DataType%>&xType="& xType
  reg.Submit
End Sub

Sub del()
 chky=msgbox("確定刪除?", 48+1, "刪除確認")
 if chky=vbok then
   xType = "Del"
   reg.action="CatConfig.asp?Language=<%=Language%>&DataType=<%=DataType%>&xType="& xType
   reg.Submit
 End If
End Sub

function blen(xs)
 xl = len(xs)
 for i=1 to len(xs)
  if asc(mid(xs,i,1))<0 then xl = xl + 1
 next
 blen = xl
end function

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
	 If xType = "OrderBy" And window.event.KeyCode = 16 or window.event.KeyCode = 17 Then
	  document.all.OrderReg.CatID.multiple = False
	 End IF
	End Sub

	Sub Document_OnKeyUp()
	  If xType = "OrderBy" Then document.all.OrderReg.CatID.multiple = True
	End Sub

	Sub FormSubmit()
	 Optpos = OrderReg.CatID.options.length
	 for OptCount = 0 to Optpos - 1
	   document.all.OrderReg.CatID.options(OptCount).selected = True
	 next
	 OrderReg.submit
	End Sub
</script>
