<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<title>新增資料</title>
<script language="vbscript">
Sub datacheck()
<% If CatDecide = "Y" Then %>
  msg2 = "請選擇「類別」！"
  If BoardForm.xin_CatID.value = "" Then
     MsgBox msg2, 16, "Sorry!"
     BoardForm.xin_CatID.focus
     Exit Sub
  End if
<% End IF %>
  msg1 = "「<%=subject%>」欄位不得為空白！"
  If BoardForm.xfn_subject.value = Empty Then
     MsgBox msg1, 16, "Sorry!"
     BoardForm.xfn_subject.focus
     Exit Sub
  End if
<% If DateDecide = "Y" Then %>
  msg3 = "「公佈起始時間」欄位不得為空白！"
  msg4 = "「公佈結束時間」欄位不得為空白！"
  msg5 = "「公佈結束時間」早於「公佈起始時間」，請重新輸入 !"
  If BoardForm.xfn_BeginDate.value = Empty Then
     MsgBox msg3, 16, "Sorry!"
     BoardForm.xfn_BeginDate.focus
     Exit Sub
  ElseIf BoardForm.xfn_EndDate.value = Empty Then
     MsgBox msg4, 16, "Sorry!"
     BoardForm.xfn_EndDate.focus
     Exit Sub
  ElseIf CDate(BoardForm.xfn_BeginDate.value) > CDate(BoardForm.xfn_EndDate.value) Then
     MsgBox msg5, 16, "Sorry!"
     BoardForm.xfn_EndDate.focus
     Exit Sub
  End if
<% End If %>
 For fno=1 to document.BoardForm.length-1
  IF left(document.BoardForm(fno).name,7) = "Content" Then
    fono = mid(document.BoardForm(fno).name,8)
	  If document.BoardForm("Content"& fono).value = Empty Then
	     MsgBox "第 "& document.BoardForm("No"& fono).value &" 段<%=Content%>不得為空白！", 16, "Sorry!"
	     document.BoardForm("Content"& fono).focus
	     Exit Sub
	  End if

	  If document.BoardForm("ImageFile"& fono).value <> Empty Then
	   filename = lcase(mid(document.BoardForm("ImageFile"& fono).value, instr(document.BoardForm("ImageFile"& fono).value,".")))
	   If filename <> ".gif" And filename <> ".jpg" Then
	      MsgBox "第 "& document.BoardForm("No"& fono).value &" 段附圖格式本系統不接受，本系統附圖僅支援 GIF、JPG 格式！", 16, "Sorry!"
	      document.BoardForm("ImageFile"& fono).focus
	     Exit Sub
	   End If
	  End if
  End If
 Next

  BoardForm.Submit
End Sub
<% If DateDecide = "Y" Then %>
dim CanTarget
sub xbtdate(n,xdate)
	if xdate = "" then	xdate = date()
	document.all.calendar.way = xdate
	
'	document.all.calendar.xxx xdate
 If document.all.calendar.style.visibility="" Then
   document.all.calendar.style.visibility="hidden"
 Else
   document.all.calendar.style.visibility=""
 End If
 CanTarget=n
end sub

sub btdate(n)
 If document.all.calendar.style.visibility="" Then
   document.all.calendar.style.visibility="hidden"
 Else
   document.all.calendar.style.visibility=""
 End If
 CanTarget=n
end sub

sub calendar_onscriptletevent(n,o)
  document.all.calendar.style.visibility="hidden"
  select case CanTarget
     case 1
          document.all.xfn_Begindate.value=n
          document.all.xfn_Enddate.value=cdate(n) + 15
     case 2
          document.all.xfn_Enddate.value=n
  end select
end sub
<% End IF %>
</script>
</head>
<body>
<center>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=Title%>&nbsp;<font size=2>【新增】</font></td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
          <% if (HTProgRight and 2)=2 then %><a href="BulletinList.asp?Language=<%=Language%>&DataType=<%=DataType%>">資料查詢</a><% End IF %>
   　   </td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15">附圖寬度建議小於 200 圖素，且需為 GIF 或 JPG 格式。</td> 
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    

<form method="POST" name="BoardForm" ENCTYPE="MULTIPART/FORM-DATA" action="BulletinAddConfig.asp?Language=<%=Language%>&DataType=<%=DataType%>">
  <input type="hidden" name="EditDate" value="<%=date%>">
<% If DateDecide = "Y" Then %>
  <object data=../inc/calendar.htm id=calendar type=text/x-scriptlet width=245 height=160 style="position: absolute; top: 30; left: 351; visibility: hidden"></object>
<% End IF %>
  <table border="0" width="580" cellspacing="1" cellpadding="3" class="bluetable">
<% If CatDecide = "Y" Then %>
    <tr>
      <td width="100" align="right" class="lightbluetable">類別</td>
      <td width="480" class="whitetablebg"><select size="1" name="xin_CatID">
          <option value="">請選擇..</option>
    <% SQL = "SELECT * FROM DataCat Where Language = N'"& Language &"' And DataType = N'"& DataType &"' Order By CatShowOrder"
       set rs = conn.execute(SQL)
        If Not rs.EOF Then
         Do while not rs.eof %>
          <option value="<%=rs("CatID")%>"><%=rs("CatName")%></option>
    <%	 rs.movenext
    	 loop
    	End If %>
          <option value="Add" style="color:red">新增</option>
          </select></td>
    </tr>
<% End IF %>
    <tr>
      <td width="100" align="right" class="lightbluetable"><%=Subject%></td>
      <td width="480" class="whitetablebg"><input type="text" name="xfn_Subject" size="40"></td>
    </tr>
<% If DateDecide = "Y" Then %>
    <tr>
      <td width="100" align="right" class="lightbluetable">公佈時間</td>
      <td width="480" class="whitetablebg">
        <input type=text name=xfn_Begindate size=10 readonly value="" style="cursor:hand;" onclick="VBScript: btdate 1">
      ～<input type=text name=xfn_Enddate size=10 readonly value="" style="cursor:hand;" onclick="VBScript: btdate 2"></td>
    </tr>
<% End IF %>
<% If EMailDecide = "Y" Then %>
    <tr>
      <td width="100" align="right" class="lightbluetable">是否發佈</td>
      <td width="480" class="whitetablebg"><input type="radio" name="EMailCk" value="Y">是　<input type="radio" name="EMailCk" value="N" checked>否</td>
    </tr>
<% End IF %>
<% If Not IsNull(Extend_1) Then %>
    <tr>
      <td width="100" align="right" class="lightbluetable"><%=Extend_1%></td>
      <td width="480" class="whitetablebg">http://<input type="text" name="xfn_Extend_1" size="40"></td>
    </tr>
<% End IF %>
  </table>

  <DIV ID="BoardTB">
  <table ID="Table1" border="0" width="580" cellspacing="1" cellpadding="3" class="bluetable">
    <tr>
      <td width="10" align="center" class="lightbluetable">第<input type="text" name="No1" size="1" style="text-align:center;" readonly value="1" class=SEdit>段</td>
      <td width="90" align="right" valign="top" class="lightbluetable"><%=Content%></td>
      <td width="480" class="whitetablebg"><textarea rows="4" name="Content1" cols="60"></textarea>
      </td>
    </tr>
    <tr style="display='none'" id=ImageFileTr1>
      <td width="10" align="center" class="lightbluetable"></td>
      <td width="90" align="right" class="lightbluetable">附圖</td>
      <td width="480" id="ImageTd1" class="whitetablebg"><input type=file name="imagefile1" value="">　　<input type="radio" value="left" name="ImageWay1">
        左　<input type="radio" name="ImageWay1" value="right" checked>
        右</td>
    </tr>

    <tr id="toolbar1" style="display:'none'" class="whitetablebg">
      <td width="10" align="center"></td>
      <td width="90" align="center"></td>
      <td width="480">
        <table border="1" cellspacing="3" cellpadding="0" bgcolor="#D6D3CE" bordercolor="#D6D3CE" bordercolorlight="#808080" bordercolordark="#FFFFFF">
          <tr>
            <td><img border="0" src="../NewImages/image.gif" id=ToolBara1 alt="插入圖片" style="cursor:hand" width="23" height="22"><img border="0" src="../NewImages/image4.gif" id=ToolBarc1 style="display:'none';cursor:hand" alt="刪除附圖" width="23" height="22"><img border="0" src="../NewImages/image5.gif" id=ToolBard1 alt="在上方插入一個段落" style="cursor:hand" width="23" height="22"><img border="0" src="../NewImages/image6.gif" id=ToolBarf1 alt="在下方插入一個段落" style="cursor:hand" width="23" height="22"><img border="0" src="../NewImages/image2.gif" id=ToolBarg1 alt="刪除此段落" style="cursor:hand" width="23" height="22"></td>
          </tr>
        </table>
      </td>
    </tr>

  </table>
  </DIV>
  <p align="right">
  <% if (HTProgRight and 4)=4 then %><input type="button" value="確定" class="cbutton" OnClick="datacheck()"><input type="reset" value="取消" class="cbutton"><% End If %>
</form>
    </td>
  </tr>  
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      Best viewed with Internet Explorer 5.5+ , Screen Size set to 800x600.</td>                                          
  </tr> 
</table> 

</center>

<!-- 程式結束 -->
</body></html>
<% If CatDecide = "Y" Then %><!--#Include file = "AddSelect.inc" --><% End IF %>
<script language="vbscript">

mycount = 1
nowcount = 1

Sub AddKey(cktext,ojno)
 mycount = mycount + 1
 nowcount = nowcount + 1


		srcHtml = "<TABLE ID=TABLE"& mycount &" border=0 cellPadding=3 cellSpacing=1 class=bluetable width=580>"
		srcHtml = srcHtml & "<tr><td width=10 align=center class=lightbluetable>"
		srcHtml = srcHtml & "第<INPUT class=SEdit name=No"& mycount &" readonly size=1 style=TEXT-ALIGN:center; value="& mycount &">段</TD>"
		srcHtml = srcHtml & "<td width=90 align=right valign=top class=lightbluetable><%=Content%></td>"
		srcHtml = srcHtml & "<td width=480 class=whitetablebg><textarea rows=4 name=Content"& mycount &" cols=60></textarea>"
		srcHtml = srcHtml & "</td></tr><tr id=ImageFileTr"& mycount &" style=display='none'><td width=10 align=center class=lightbluetable></td>"
		srcHtml = srcHtml & "<td width=90 align=right class=lightbluetable>附圖</td><td width=480 id=ImageTd"& mycount &" class=whitetablebg>"
		srcHtml = srcHtml & "<input type=file name=imagefile"& mycount &">　　<input type=radio value=left name=ImageWay"& mycount &"> 左　<input type=radio name=ImageWay"& mycount &" value=right checked> 右</td></tr>"
		srcHtml = srcHtml & "<tr id=toolbar"& mycount &" style=display:'none' class=whitetablebg><td width=10 align=center></td><td width=90 align=center></td> <td width=480>"
		srcHtml = srcHtml & "<table border=1 cellspacing=3 cellpadding=0 bgcolor=#D6D3CE bordercolor=#D6D3CE bordercolorlight=#808080 bordercolordark=#FFFFFF>"
		srcHtml = srcHtml & "<tr><td><img border=0 src=../NewImages/image.gif id=ToolBara"& mycount &" alt=插入圖片 style=cursor:hand><img border=0 src=../NewImages/image4.gif id=ToolBarc"& mycount &" style=display:'none';cursor:hand alt=刪除附圖><img border=0 src=../NewImages/image5.gif id=ToolBard"& mycount &" style=cursor:hand alt=在上方插入一個段落><img border=0 src=../NewImages/image6.gif id=ToolBarf"& mycount &" alt=在下方插入一個段落 style=cursor:hand><img border=0 src=../NewImages/image2.gif id=ToolBarg"& mycount &" alt=刪除此段落 style=cursor:hand></td></tr></table>"
		srcHtml = srcHtml & "</td></tr></table>"

	  If cktext = "f" Then
	   document.all("Table" & ojno).InsertAdjacentHTML "afterEnd", srcHtml
	  ElseIf cktext = "d" Then
	   document.all("Table" & ojno).InsertAdjacentHTML "beforeBegin", srcHtml
	  End If

	   document.all("Content"& mycount).focus
	   ToolBarSwitch(mycount)
End sub

Sub BoardTB_OnClick()
    set KeyObj = window.event.srcElement
    If KeyObj.TagName = "TEXTAREA" or KeyObj.TagName = "IMG" Then
    If left(KeyObj.id,7) = "ToolBar" Then
     SetTagItem = mid(KeyObj.id,8,1)
     newcount = mid(KeyObj.id,9)

		Select Case SetTagItem
		 Case "a"
		  document.all("ImageFileTr"& newcount).style.display = ""
		  document.all("ToolBara"& newcount).style.display = "none"
		  document.all("ToolBarc"& newcount).style.display = ""
		 Case "c"
			ImageTdSrc = "<input type=file name=imagefile"& newcount &">　　<input type=radio value=left name=ImageWay"& newcount &">"
			ImageTDSrc = ImageTDSrc & " 左　<input type=radio name=ImageWay"& newcount &" value=right checked> 右"
		  document.all("ImageFileTr"& newcount).style.display = "none"
		  document.all("ToolBarc"& newcount).style.display = "none"
		  document.all("ToolBara"& newcount).style.display = ""
		  document.all("ImageTd"& newcount).InnerHTML = ImageTdSrc
		  'alert document.all("imagefile"& newcount).value
		 Case "d"
		  AddKey "d",newcount
		 Case "f"
		  AddKey "f",newcount
		 Case "g"
		  If nowcount = 1 Then
		   Alert "最少要有一段"& Content &"！"
		  Else
		   nowcount = nowcount - 1
		   document.all("TABLE"& newcount).outerHTML = ""
		   trcount = 0
			 For fno=1 to document.BoardForm.length-1
			  IF left(document.BoardForm(fno).name,7) = "Content" Then
			    fono = mid(document.BoardForm(fno).name,8)
			    trcount = trcount + 1
				document.BoardForm("No"& fono).value = trcount
			  End If
			 Next
		   End IF
		End Select
    ElseIF left(KeyObj.Name,7) = "Content" Then
      ToolBarSwitch mid(KeyObj.Name,8)
	End IF
	End If
End Sub

	Sub ToolBarSwitch(barno)
	   trcount = 0
	   for xtableno = 1 to document.BoardForm.length-1
 	    IF left(document.BoardForm(xtableno).name,7) = "Content" Then
	     fonoxtableno = mid(document.BoardForm(xtableno).name,8)
	     document.all("toolbar"& fonoxtableno).style.display = "none"
	     trcount = trcount + 1
	     document.BoardForm("No"& fonoxtableno).value = trcount
	    End If
	   next
	  document.all("toolbar"& barno).style.display = ""
	End Sub

</script>
