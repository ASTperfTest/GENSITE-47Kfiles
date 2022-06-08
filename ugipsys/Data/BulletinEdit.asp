<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<% UnitID = Request.QueryString("UnitID") %>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<title>資料編修</title>
</head><body>
<center>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=Title%>&nbsp;<font size=2>【編修】</font></td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
          <% if (HTProgRight and 8)=8 And EMailDecide = "Y" then %><a href="BulletinEMail.asp?Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>">發佈此公告</a><% End IF %>
          <% if (HTProgRight and 16)=16 then %>　<a href="JavaScript: delUnit();">刪除</a><% End IF %>
          <% if (HTProgRight and 2)=2 then %>　<a href="BulletinList.asp?Language=<%=Language%>&DataType=<%=DataType%>">查詢</a><% End IF %>
   　   </td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15">編修圖文，請點選段落</td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    

<br>
<!--#Include file = "Table.inc" -->
<br>
<table border="0" width="90%" cellspacing="1" cellpadding="0" class="whitetablebg">
 <tr><td>
<% 	SQLCom = "SELECT * FROM DataContent Where UnitID = "& UnitID &" Order By Position"
	SET RSCom = conn.execute(SQLCom)
	 If Not rscom.EOF Then
	  DIVCount = 0
	  Do while not rscom.EOF
	   DIVCount = DIVCount + 1
	   ImageHTELSrc = ""
	   ClientCk = "N"
	   comm = rscom("Content")
	   ncomm = message(comm)
	   If Not (rscom("ImageFile") = "" or IsNull(rscom("ImageFile"))) Then
	    ClientCk = "Y"
	    ImageHTELSrc = "<DIV ID=Con"& rscom("Position") &" OnClick=""VBS: EditContent "& rscom("ContentID") &","& rscom("Position") &",'"& ClientCk &"'"" title=""第 "& rscom("Position") &" 段"& Content &""" style=cursor:hand>"
	    ImageHTELSrc = ImageHTELSrc & "<img src=../Public/Data/" & rscom("ImageFile") &" border=0 align="& rscom("ImageWay") &" id=conimg"& rscom("Position") &" alt=""第 "& rscom("Position") &" 段圖片"">"
	   Else
	    ImageHTELSrc = "<DIV ID=Con"& rscom("Position") &" OnClick=""VBS: EditContent "& rscom("ContentID") &","& rscom("Position") &",'"& ClientCk &"'"" title=""第 "& rscom("Position") &" 段"& Content &""" style=cursor:hand>"
	   End IF
	   response.write ImageHTELSrc & ncomm & "<textarea rows=6 name=Content"& rscom("Position") &" cols=40 style=display:none>"& rscom("Content") &"</textarea><P></DIV>" & vbcrlf
	  rscom.movenext
      loop
     End If %>
 </td></tr>
</table>

    </td>
  </tr>  
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      Best viewed with Internet Explorer 5.5+ , Screen Size set to 800x600.</td>                                          
  </tr> 
</table> 

<table width="100%" id="toolsTable" style="position:absolute;display:none" border=1 cellspacing=0 bgcolor=#D6D3CE bordercolor=#D6D3CE bordercolorlight=#808080 bordercolordark=#FFFFFF>
 <tr><td>
    <div align="right">
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr>
         <td>&nbsp;<font size="2">編輯工具</font>
 		 <img border="0" src="../NewImages/image12.gif" align="absmiddle">
 		 <img id="AddImg" border=0 src=../NewImages/image.gif alt=此段插入圖片 style=cursor:hand align="absmiddle"><img id="ImgLeft" border="0" src="../NewImages/image10.gif" alt=此段圖片靠左 style=cursor:hand align="absmiddle"><img id="ImgRight" border="0" src="../NewImages/image11.gif" alt=此段圖片靠右 style=cursor:hand align="absmiddle"><img id="ChgImg" border="0" src="../NewImages/image3.gif" alt=更換圖片 style=cursor:hand align="absmiddle"><img id="DelImg" border=0 src=../NewImages/image4.gif style=display:'none';cursor:hand alt=刪除此段附圖 align="absmiddle">
 		 <img border="0" src="../NewImages/image12.gif" align="absmiddle">
 		 <img id="AddUpStage" border=0 src=../NewImages/image5.gif style=cursor:hand alt=在此段上方插入一個段落 align="absmiddle"><img id="AddEndStage" border=0 src=../NewImages/image6.gif alt=在此段下方插入一個段落 style=cursor:hand align="absmiddle"><img id="DelStage" border=0 src=../NewImages/image2.gif alt=刪除此段落 style=cursor:hand align="absmiddle"><img id="EditStage" border="0" src="../NewImages/image7.gif" alt=編輯此段<%=Content%> style=cursor:hand align="absmiddle"></td>
          <td align="right" valign="top"><img id="CloseTools" border="0" src="../NewImages/toolsclose.gif" align="top" style=cursor:hand alt=關閉工具列取消選取></td>
        </tr>
      </table>
    </div>
 		 </td></tr>
</table>
</center>
</body></html>
<%
function message(tempstr)
  outstring = ""
  while len(tempstr) > 0
    pos=instr(tempstr, chr(13)&chr(10))
    if pos = 0 then
      outstring = outstring & tempstr & "<br>"
      tempstr = ""
    else
      outstring = outstring & left(tempstr, pos-1) & "<br>"
      tempstr=mid(tempstr, pos+2)
    end if
  wend
  message = outstring
end function

If CatDecide = "Y" Then %><!--#Include file = "AddSelect.inc" --><% End IF %>
<script language="vbscript">
	StageCount = <%=DIVCount%>
	NowType = ""
	ContentID = ""
	ImgType = ""
	nowps = ""
	xname = "OpenWindow"

	IW= 0
	IH= 0
	PX= 0
	PY= 0
	IMGW= 0
	IMGH= 22
	LSAFETY= 20
	TSAFETY= 6

	set wmark = document.all.ToolsTable

	sub wmarkPosition()
	  IH= document.body.clientHeight
	  PY= document.body.scrollTop
	  wmark.style.top = (IH+PY-(IMGH+TSAFETY))
	  wmark.style.left = 0
	end sub

	sub wmarking()
	  oldIW= IW
	  oldIH= IH
	  oldPX= PX
	  oldPY= PY
	  wmarkPosition()
	end sub

	sub window_onload()
	  setInterval "wmarking()",20
	end sub

	Sub WindowPlace()
	 DivTop = document.body.scrolltop + window.event.Y + 30
	 document.all.SetView.style.top = Divtop
     document.all.SetView.style.visibility = ""
 	 document.all.SetView.style.left = "300px"
 	 document.all.SetView.style.width = "255px"
	End Sub

	Sub ToolsWindowPlace()
	 DivTop = document.body.scrolltop + window.event.Y - 300
	 document.all.SetView.style.top = Divtop
     document.all.SetView.style.visibility = ""
	End Sub

	Sub CloseTools_OnClick()
        document.all.SetView.style.visibility = "hidden"
	 	document.all.toolsTable.style.Display = "none"
		 For DIVno = 1 to StageCount
		  document.all("Con"& DIVno).style.backgroundcolor="#FFFFFF"
		 Next
		xType = ""
		DivplayCk
	End Sub

	Sub DivplayCk()
		ListTableForm.reset
		document.all.CatEditTable.style.Display = "none"
		document.all.DataEditTable.style.Display = "none"
		document.all.SubjectEditTable.style.Display = "none"
		document.all.AddImgEditTable.style.Display = "none"
		document.all.ChgImgEditTable.style.Display = "none"
		document.all.AddUpStageEditTable.style.Display = "none"
		document.all.AddEndStageEditTable.style.Display = "none"
		document.all.EditStageEditTable.style.Display = "none"
		document.all.Extend_1EditTable.style.Display = "none"
	End Sub

	Sub EditContent(ConID,Obno,ImgCk)
		DivplayCk
        document.all.SetView.style.visibility = "hidden"
<% If DateDecide = "Y" Then %>
	    document.all.calendar.style.visibility="hidden"
<% End if %>
		If StageCount < 2 Then
		 document.all.DelStage.style.Display = "none"
		Else
		 For DIVno = 1 to StageCount
		  document.all("Con"& DIVno).style.backgroundcolor="#FFFFFF"
		 Next
		End If
		nowps = Obno
		ContentID = ConID
		document.all("Con"& Obno).style.backgroundcolor="#FFFF99"
	 	document.all.toolsTable.style.Display = ""

		If ImgCk = "Y" Then
		 ImgType = "Y"
		 document.all.AddImg.style.Display = "none"
		 document.all.ChgImg.style.Display = ""
		 document.all.DelImg.style.Display = ""
		 If document.all("conimg"& Obno).align = "right" Then
		  document.all.ImgLeft.style.Display = ""
		  document.all.ImgRight.style.Display = "none"
		 ElseIf document.all("conimg"& Obno).align = "left" Then
		  document.all.ImgLeft.style.Display = "none"
		  document.all.ImgRight.style.Display = ""
		 End If
		Else
		 ImgType = "N"
		 document.all.AddImg.style.Display = ""
		 document.all.ChgImg.style.Display = "none"
		 document.all.ImgLeft.style.Display = "none"
		 document.all.ImgRight.style.Display = "none"
		 document.all.DelImg.style.Display = "none"
		End IF
	End Sub

	Sub ImgLeft_OnClick()
	  xType = "ImgLeft"
	  document.location.href="BulletinEditConfig.asp?ContentID="& ContentID &"&Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>&EditType="& xType
	End Sub

	Sub ImgRight_OnClick()
	  xType = "ImgRight"
	  document.location.href="BulletinEditConfig.asp?ContentID="& ContentID &"&Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>&EditType="& xType
	End Sub

	Sub DelImg_OnClick()
     chky=msgbox("刪除此段附圖!"& vbcrlf & vbcrlf &"　確定嗎？"& vbcrlf , 48+1, "請注意！！")
	    if chky=vbok then
	     xType = "DelImg"
	     document.location.href="BulletinEditConfig.asp?ContentID="& ContentID &"&Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>&EditType="& xType
	    Else
	     Exit Sub
	    End If
	End Sub

	Sub AddUpStage_OnClick()
		DivplayCk
 	    ToolsWindowPlace
 	    document.all.SetView.style.left = "50px"
 	    document.all.SetView.style.width = "500px"
		document.all.AddUpStageEditTable.style.Display = ""
		document.all.EditTitle.InnerText = "於第 "& nowps &" 段上方新增段落"
		document.all.AddUsNowps.value = nowps
		NowType = "AddUpStage"
	End Sub

	Sub AddEndStage_OnClick()
		DivplayCk
 	    ToolsWindowPlace
 	    document.all.SetView.style.left = "50px"
 	    document.all.SetView.style.width = "500px"
		document.all.AddEndStageEditTable.style.Display = ""
		document.all.EditTitle.InnerText = "於第 "& nowps &" 段下方新增段落"
		document.all.AddEsNowps.value = nowps
		NowType = "AddEndStage"
	End Sub

	Sub EditStage_OnClick()
		DivplayCk
 	    ToolsWindowPlace
 	    document.all.SetView.style.left = "50px"
 	    document.all.SetView.style.width = "500px"
		document.all.EditStageEditTable.style.Display = ""
		document.all.EditTitle.InnerText = "編輯第 "& nowps &" 段<%=Content%>"
		document.all.EsContent.value = document.all("Content"& nowps).value
		NowType = "EditStage"
	End Sub

	Sub ChgImg_OnClick()
		DivplayCk
 	    ToolsWindowPlace
 	    document.all.SetView.style.left = "50px"
 	    document.all.SetView.style.width = "500px"
		document.all.EditTitle.InnerText = "變更附圖"
		document.all.ChgImgEditTable.style.Display = ""
		NowImgsrc = document.all("conimg"& nowps).src
		document.all.ChgImgsrc.src = NowImgsrc
		NowType = "ChgImg"
	End Sub

	Sub AddImg_OnClick()
		DivplayCk
 	    ToolsWindowPlace
		document.all.EditTitle.InnerText = "新增附圖"
		document.all.AddImgEditTable.style.Display = ""
 		document.all.SetView.style.left = "300px"
	 	document.all.SetView.style.width = "255px"
		NowType = "AddImg"
	End Sub

	Sub CatEdit_OnClick()
		DivplayCk
        WindowPlace
		document.all.EditTitle.InnerText = "編修類別"
		document.all.CatEditTable.style.Display = ""
		NowType = "EditCat"
	End Sub

	Sub DateEdit_OnClick()
		DivplayCk
        WindowPlace
		document.all.EditTitle.InnerText = "編修公佈時間"
		document.all.DataEditTable.style.Display = ""
		NowType = "EditData"
	End Sub

	Sub SubjectEdit_OnClick()
		DivplayCk
        WindowPlace
		document.all.EditTitle.InnerText = "編修<%=Subject%>"
		document.all.SubjectEditTable.style.Display = ""
		NowType = "EditSubject"
	End Sub

	Sub Extend_1Edit_OnClick()
		DivplayCk
        WindowPlace
		document.all.EditTitle.InnerText = "編修<%=Extend_1%>"
		document.all.Extend_1EditTable.style.Display = ""
		NowType = "EditExtend_1"
	End Sub

	Sub DelStage_OnClick()
     IF ImgType = "N" Then
      chky=msgbox("刪除此段落!"& vbcrlf & vbcrlf &"　確定嗎？"& vbcrlf , 48+1, "請注意！！")
     ElseIF ImgType = "Y" Then
      chky=msgbox("刪除此段落與附圖!"& vbcrlf & vbcrlf &"　確定嗎？"& vbcrlf , 48+1, "請注意！！")
     End IF
	    if chky=vbok then
	     xType = "DelStage"
	     document.location.href="BulletinEditConfig.asp?ContentID="& ContentID &"&Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>&EditType="& xType
	    Else
	     Exit Sub
	    End If
	End Sub

	Sub EditSubmit(xType)
		If xType = "EditCat" Then
		  msg2 = "請選擇「類別」！"
		  If document.all.xin_CatID.value = "" Then
		     MsgBox msg2, 16, "Sorry!"
		     document.all.xin_CatID.focus
		     Exit Sub
		  End if
		ElseIf xType = "EditSubject" Then
		  msg1 = "「<%=subject%>」欄位不得為空白！"
		  If document.all.subject.value = Empty Then
		     MsgBox msg1, 16, "Sorry!"
    		 document.all.subject.focus
     		 Exit Sub
		  End if
		ElseIf xType = "EditData" Then
		  msg3 = "「公佈起始時間」欄位不得為空白！"
 		  msg4 = "「公佈結束時間」欄位不得為空白！"
		  msg5 = "「公佈結束時間」早於「公佈起始時間」，請重新輸入 !"
		  If document.all.BeginDate.value = Empty Then
		     MsgBox msg3, 16, "Sorry!"
		     document.all.BeginDate.focus
		     Exit Sub
		  ElseIf document.all.EndDate.value = Empty Then
		     MsgBox msg4, 16, "Sorry!"
		     document.all.EndDate.focus
		     Exit Sub
		  ElseIf CDate(document.all.BeginDate.value) > CDate(document.all.EndDate.value) Then
		     MsgBox msg5, 16, "Sorry!"
 			 document.all.EndDate.focus
		     Exit Sub
		  End if
		ElseIf xType = "AddImg" Then
			If document.all.ImageFile.value <> Empty Then
			   filename = lcase(mid(document.all.ImageFile.value, instr(document.all.ImageFile.value,".")))
			   If filename <> ".gif" And filename <> ".jpg" Then
			      MsgBox "附圖格式本系統不接受，本系統附圖僅支援 GIF、JPG 格式！", 16, "Sorry!"
			      document.all.ImageFile.focus
				  Exit Sub
			   End If
			Else
			   MsgBox "請選擇附圖！", 16, "Sorry!"
			   document.all.ImageFile.focus
			   Exit Sub
			End if
		ElseIf xType = "ChgImg" Then
			If document.all.ChgImageFile.value <> Empty Then
			   filename = lcase(mid(document.all.ChgImageFile.value, instr(document.all.ChgImageFile.value,".")))
			   If filename <> ".gif" And filename <> ".jpg" Then
			      MsgBox "附圖格式本系統不接受，本系統附圖僅支援 GIF、JPG 格式！", 16, "Sorry!"
			      document.all.ChgImageFile.focus
				  Exit Sub
			   End If
			Else
			   MsgBox "請選擇附圖！", 16, "Sorry!"
			   document.all.ChgImageFile.focus
			   Exit Sub
			End if
		ElseIf xType = "AddUpStage" Then
		  msg10 = "「<%=Content%>」欄位不得為空白！"
		  If document.all.AddUsContent.value = Empty Then
		     MsgBox msg10, 16, "Sorry!"
    		 document.all.AddUsContent.focus
     		 Exit Sub
		  End if
		  If document.all.AddUsImageFile.value <> Empty Then
			   filename = lcase(mid(document.all.AddUsImageFile.value, instr(document.all.AddUsImageFile.value,".")))
			   If filename <> ".gif" And filename <> ".jpg" Then
			      MsgBox "附圖格式本系統不接受，本系統附圖僅支援 GIF、JPG 格式！", 16, "Sorry!"
			      document.all.AddUsImageFile.focus
				  Exit Sub
			   End If
		  End If
		ElseIf xType = "AddEndStage" Then
		  msg11 = "「<%=Content%>」欄位不得為空白！"
		  If document.all.AddEsContent.value = Empty Then
		     MsgBox msg11, 16, "Sorry!"
    		 document.all.AddEsContent.focus
     		 Exit Sub
		  End if
		  If document.all.AddEsImageFile.value <> Empty Then
			   filename = lcase(mid(document.all.AddEsImageFile.value, instr(document.all.AddEsImageFile.value,".")))
			   If filename <> ".gif" And filename <> ".jpg" Then
			      MsgBox "附圖格式本系統不接受，本系統附圖僅支援 GIF、JPG 格式！", 16, "Sorry!"
			      document.all.AddEsImageFile.focus
				  Exit Sub
			   End If
		  End If
		ElseIf xType = "EditStage" Then
		ElseIf xType = "" Then
		  Exit Sub
		End If

		ListTableForm.action = "BulletinEditConfig.asp?ContentID="& ContentID &"&Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>&EditType="& xType
		ListTableForm.Submit
	End Sub

	Sub delUnit()	'--- 刪除
		xChoo = Msgbox("確定刪除?", vbYesNo+32, "刪除確認")
		IF xChoo=6 then document.location.href="BulletinEditConfig.asp?Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>&EditType=Del"
	End sub

dim CanTarget
sub btdate(n)
 If document.all.calendar.style.visibility="" Then
   document.all.calendar.style.visibility="hidden"
 Else
   document.all.calendar.style.visibility=""
   DivTop = document.body.scrolltop + window.event.Y + 30
   document.all.calendar.style.top = Divtop
 End If
 CanTarget=n
end sub

sub calendar_onscriptletevent(n,o)
  document.all.calendar.style.visibility="hidden"
  select case CanTarget
     case 1
          document.all.Begindate.value=n
          document.all.Enddate.value=cdate(n) + 15
     case 2
          document.all.Enddate.value=n
  end select
end sub

</script>