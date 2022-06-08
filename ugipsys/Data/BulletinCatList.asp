<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
 CatSQL = Request.QueryString("CatSQL")
'===> 設定共有幾個欄位【請從 0 開始計算】
	rowsno = 1

'===> 設定說明內容
	ReadmeText = "※ 檢視請點選「"& Subject &"」"

'===> 設定 SQL
	DefOrderList = ""

  If  Request.Form("SelectKey") <> "" Then
    SetSQL = "SELECT DISTINCT DataUnit.UnitID, DataUnit.Subject, DataCat.CatName "&_
			 "FROM DataCat INNER JOIN "&_
      		 "DataUnit ON DataCat.CatID = DataUnit.CatID RIGHT OUTER JOIN "&_
      		 "DataContent ON DataUnit.UnitID = DataContent.UnitID Where DataUnit.DataType = N'"& DataType &"' And DataUnit.Language = N'"& Language &"'"
  Else
    SetSQL = "SELECT DataUnit.UnitID, DataUnit.Subject, DataCat.CatName "&_
			 "FROM DataCat INNER JOIN "&_
      		 "DataUnit ON DataCat.CatID = DataUnit.CatID "&_
      		 "Where DataUnit.DataType = N'"& DataType &"' And DataUnit.Language = N'"& Language &"'"
  End IF

         xNowCatID = Request.Form("CatID")
         If CatSQL <> "" Then  xNowCatID = CatSQL

     If xNowCatID <> "" And Request.Form("SelectKey") = "" Then
	  SetSQL = SetSQL & " And DataCat.CatID = "& xNowCatID &" Order By DataCat.CatShowOrder, DataUnit.ShowOrder"
     ElseIf xNowCatID <> "" And Request.Form("SelectKey") <> "" Then
	  SetSQL = SetSQL & " And DataCat.CatID = "& xNowCatID
	  SetSQL = SetSQL & " And (DataContent.Content Like N'%"& Request.Form("SelectKey") &"%' or DataUnit.Subject Like N'%"& Request.Form("SelectKey") &"%')"
	 ElseIf xNowCatID = "" And Request.Form("SelectKey") <> "" Then
	  SetSQL = SetSQL & " And (DataContent.Content Like N'%"& Request.Form("SelectKey") &"%' or DataUnit.Subject Like N'%"& Request.Form("SelectKey") &"%')"
	 ElseIf xNowCatID = "" And Request.Form("SelectKey") = "" Then
	  SetSQL = SetSQL & " Order By DataCat.CatShowOrder, DataUnit.ShowOrder"
     End IF


'===> 欄位排序設定
	OrderField = ""
	OrderName = ""


%>
<!--#Include file = "PageList_Head.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<title>資料清單 - 分類</title>
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
          <% if (HTProgRight and 8)=8 And xNowCatID <> "" and Request.Form("SelectKey") = "" And OrderByCk = "Y" then %><SPAN OnClick="FormOrderOpen()">排序</SPAN><% End IF %>
          <% if (HTProgRight and 4)=4 then %>　<a href="BulletinAdd.asp?Language=<%=Language%>&DataType=<%=DataType%>">新增</a><% End IF %>
          <% if (HTProgRight and 2)=2 then %>　<a href="CatList.asp?Language=<%=Language%>&DataType=<%=DataType%>">分類編修</a><% End IF %>
   　   </td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15">不設條件可查詢所有資料。檢視內容，請點選「<%=Subject%>」。選擇類別，關鍵字欄位空白查詢，可排序所選類別內資料。</td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    

<br>
<form method="POST" action="BulletinCatList.asp?Language=<%=Language%>&DataType=<%=DataType%>">
  <table border="0" cellspacing="0" cellpadding="3" class="c12-1">
    <tr>
      <td>類別：</td>
      <td><select size="1" name="CatID">
    <% SQLCom = "SELECT * FROM DataCat Where Language = N'"& Language &"' And DataType = N'"& DataType &"' Order By CatShowOrder"
       set rs2 = conn.execute(SQLCom)
        If Not rs2.EOF Then
         NowCatID = Request.Form("CatID")
          If Request.Form("CatID") = "" Then NowCatID = 0
          If CatSQL <> "" Then  NowCatID = CatSQL %>
          <option value=""></option>
      <% Do while not rs2.eof %>
          <option value="<%=rs2("CatID")%>" <% If rs2("CatID") = cint(NowCatID) Then %>selected<% End if %>><%=rs2("CatName")%></option>
    <%	 rs2.movenext
    	 loop
    	Else %>
          <option value="" style=color:red>無類別資料</option>
     <% End If %>
          </select></td>
      <td width="10">&nbsp;</td>
      <td>關鍵字：</td>
      <td><input type="text" name="SelectKey" size="15"></td>
      <td><input type="submit" value="查詢"></td>
    </tr>
  </table>
</form>
<form name="lform" method="POST">
<!--#Include file = "PageList_body.inc" -->
            <div id="sortTable">
              <table border="0" width="95%" cellpadding="3" cellspacing="1" RULES="GROUPS" FRAME="HSIDES" id="tbx" class="bluetable">
                <THEAD>
                <tr style="cursor:hand;">
                  <th id="tbfs0" class="lightbluetable" width="25%">類別&nbsp;<img ID="xtbfs0" border="0" src="../images/i_down.gif"></th>
                  <th id="tbfs1" class="lightbluetable" width="75%"><%=Subject%>&nbsp;<img ID="xtbfs1" border="0" src="../images/i_down.gif"></th>
                </tr>
                </THEAD>
             <% for gxo=1 to SetPageCon
			     if not RS.eof then %>
                <tr class="whitetablebg" id=tbr<%=gxo%>>
                  <td><%=rs("CatName")%></td>
                  <td><a href="BulletinView.asp?Language=<%=Language%>&amp;DataType=<%=DataType%>&amp;UnitID=<%=rs("UnitID")%>"><%=rs("Subject")%></a></td>
                </tr>
			 <%  RS.moveNext
			     End IF
			    Next %>
			  </table>
			</div>
</form>
    </td>
  </tr>  
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      Best viewed with Internet Explorer 5.5+ , Screen Size set to 800x600.</td>                                            
  </tr> 
</table> 
</center>
<% If xNowCatID <> "" and Request.Form("SelectKey") = "" Then %>
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

         <form method="POST" action="UnitOrderConfig.asp?Language=<%=Language%>&DataType=<%=DataType%>&xType=UnitOrderBy&CatSQL=<%=xNowCatID%>" name="Orderreg">
          <table border="0" cellspacing="0" cellpadding="3" class="c12-3">
           <tr><td align="center">
    <% SQLCom = "SELECT Subject, EditDate, UnitID "&_
			    "FROM DataUnit Where "&_
		        "CatID = "& xNowCatID &" And Language = N'"& Language &"' And DataType = N'"& DataType &"' Order By ShowOrder"
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
<% End If %>
</body></html>
<!--#Include file = "PageList_End.inc" -->
<%
If xNowCatID <> "" and Request.Form("SelectKey") = "" Then

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
<% End IF %>