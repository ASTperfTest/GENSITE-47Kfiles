<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
	ItemCount = 0
	OrderByCk = ""
	SQLCom = "SELECT Subject, BeginDate, EndDate, UnitID "&_
			 "FROM DataUnit Where '"& Date() &"' between BeginDate and EndDate "&_
		     "And Language = N'"& Language &"' And DataType = N'"& DataType &"' Order By ShowOrder"
	SET RS = conn.execute(SQLCom)
	If Not rs.EOF Then OrderByCk = "Y"
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>資料清單 - 公佈時間內</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<center>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=Title%>&nbsp;<font size=2>【清單】</font></td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
          <% if (HTProgRight and 8)=8 And OrderByCk = "Y" then %><SPAN OnClick="FormOrderOpen()" style=cursor:hand>排序</SPAN><% End IF %>
          <% if (HTProgRight and 4)=4 then %>　<a href="BulletinAdd.asp?Language=<%=Language%>&DataType=<%=DataType%>">新增</a><% End IF %>
          <% if (HTProgRight and 2)=2 then %>　<a href="BulletinOldList.asp?Language=<%=Language%>&DataType=<%=DataType%>">進階查詢</a><% End IF %>
   　   </td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15">以下清單為目前有效資料。檢視資料請點選「<%=Subject%>」。</td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    

<br>
<% If rs.EOF Then %>
<br><br><font color="#FF0000" size="2">查無資料</font>
<% Else %>   
<table border="0" width="95%" cellspacing="1" cellpadding="2" class="bluetable">
  <tr class="lightbluetable">
    <td align="center" width="70%"><%=Subject%></td>
    <td align="center" width="30%">公佈期間</td>
  </tr>
  <% Do while not rs.eof
      ItemCount = ItemCount + 1 %>
  <tr class="whitetablebg">
    <td><a href="BulletinView.asp?Language=<%=Language%>&amp;DataType=<%=DataType%>&amp;UnitID=<%=rs("UnitID")%>"><%=rs("Subject")%></a></td>
    <td><%=rs("BeginDate")%> ~ <%=rs("EndDate")%></td>   
  </tr>
   <% rs.movenext
      loop %>
</table>
<% End IF %>
    </td>
  </tr>  
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      Best viewed with Internet Explorer 5.5+ , Screen Size set to 800x600.</td>                                          
  </tr> 
</table> 
</center>


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
         <form method="POST" action="UnitOrderConfig.asp?Language=<%=Language%>&DataType=<%=DataType%>&xType=UnitOrderBy" name="Orderreg">
          <table border="0" cellspacing="0" cellpadding="3" class="whitetablebg">
           <tr><td align="center">
    <% SQLCom = "SELECT Subject, BeginDate, EndDate, UnitID "&_
			 "FROM DataUnit Where '"& Date() &"' between BeginDate and EndDate "&_
		     "And Language = N'"& Language &"' And DataType = N'"& DataType &"' Order By ShowOrder"
       set rs2 = conn.execute(SQLCom)
        If Not rs2.EOF Then %>
           <select size="<%=ItemCount%>" name="UnitID" multiple>
      <% Do while not rs2.eof %>
          <option value="<%=rs2("UnitID")%>"><%=leftstr(rs2("Subject"))%></option>
    <%	 rs2.movenext
    	 loop %>
              </select>
     <% End If %>
              </td>
              <td align="center" valign="middle"><input type="button" value="↑" class="cbutton" OnClick="VBS: ChgOrderUp 'UnitID'">
										      <p><input type="button" value="↓" class="cbutton" OnClick="VBS: ChgOrderDown 'UnitID'">
              </td></tr>
          </table><br>
          <table border="0" width="95%" cellspacing="0" cellpadding="0">
           <tr><td align="right"><input type="button" value="確定" class="cbutton" OnClick="VBS: FormSubmit()"></td></tr>
          </table>
     	 </form>

    </td>
  </tr>
</table>
</body></html>
<%
function leftstr(s)
  l = 15
  If Language = "E" Then l = 30
  if len(s) > l then
    leftstr = left(s,l) & "..."
  else
    leftstr=s
  end if
end function

%>
<script language=VBScript>
Sub FormOrderOpen()
  DivTop = document.body.scrolltop + window.event.Y + 100
  document.all.SetOrderView.style.top = Divtop
  document.all.SetOrderView.style.visibility = ""
End Sub

Sub WindowClose()
  document.all.SetOrderView.style.visibility = "hidden"
  OrderReg.reset
End Sub

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
	  document.all.OrderReg.UnitID.multiple = False
	 End IF
	End Sub

	Sub Document_OnKeyUp()
	  document.all.OrderReg.UnitID.multiple = True
	End Sub

	Sub FormSubmit()
	 Optpos = OrderReg.UnitID.options.length
	 for OptCount = 0 to Optpos - 1
	   document.all.OrderReg.UnitID.options(OptCount).selected = True
	 next
	 OrderReg.submit
	End Sub
</script>
