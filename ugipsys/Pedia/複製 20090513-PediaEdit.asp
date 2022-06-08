<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="PEDIA01"
HTProgPrefix="Pedia" 

Dim iCtUnitId : iCtUnitId = session("PediaUnitId")
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!-- #INCLUDE FILE="../inc/dbFunc.inc" -->
<%
	FUNCTION mypkStr (s, cname, endchar)
		if s = "" then
			if cname = "xPostDate" then
				mypkStr = "GETDATE()"
			else
				mypkStr = "''" & endchar
			end if
		else
			pos = InStr(s, "'")
			While pos > 0
				s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
				pos = InStr(pos + 2, s, "'")
			Wend
			if cname = "xPostDate" then
				mypkStr = "'" & s & "'" & endchar
			else
				mypkStr = "N'" & s & "'" & endchar
			end if
		end if
	END FUNCTION
	
%>
<% Sub showErrBox() %>
		<script language=VBScript>
			alert "<%=errMsg%>"
			'window.history.back
		</script>
<% End sub %>
<% Sub showDoneBox(lMsg) %>
  <html>
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<meta name="GENERATOR" content="Hometown Code Generator 1.0">
			<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
			<title>編修表單</title>
    </head>
    <body>
			<script language=vbs>
			  alert("<%=lMsg%>")	    	
				<% if request.querystring("btype") = "0" then %>
					document.location.href="<%=HTprogPrefix%>List.asp?keep=Y&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>"
				<% else %>
					document.location.href="<%=HTprogPrefix%>Keyword.asp?icuitem=" & <%=request.querystring("icuitem")%> & "&phase=edit"
				<% end if %>
			</script>
    </body>
  </html>
<% End sub %>
<%
	Sub doUpdateDB()
	
		Dim oldpublic : oldpublic = ""
		Dim newpublic : newpublic = ""
		Dim memberId : memberId = ""
		sql = "SELECT fCTUPublic, memberId FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.icuitem = Pedia.gicuitem "
		sql = sql & "WHERE icuitem = " & request.querystring("icuitem")
		set frs = conn.execute(sql)
		if not frs.eof then
			oldpublic = frs("fCTUPublic")
			memberId = frs("memberId")
		end if
		frs.close
		set frs = nothing	
		newpublic = xUpForm("fCTUPublic")	
		'-----
		sql1 = ""				
		sql2 = ""
		for each form in xup.Form		
			if form.Name <> "xImgFile" and form.Name <> "CalendarTarget" and form.Name <> "submitTask" and form.Name <> "xStatus" _
				 and form.Name <> "memberId" and form.Name <> "realname" and form.Name <> "nickname" Then
				if form.Name = "engTitle" or form.Name = "formalName" or form.Name = "localName" then
					sql2 = sql2 & form.Name & " = '" & form & "',"
				else 
					if form.Name = "xImportant" then
						sql1 = sql1 & form.Name & " = " & form & ","
					elseif form = "null" then
						sql1 = sql1 & form.Name & " = ''," 
					else
						sql1 = sql1 & form.Name & " = " & mypkStr(form, form.Name, "") & ","
					end if
				end if
			end if
			if form.IsFile then
				ofname = Form.FileName
				fnExt = ""
				if instrRev(ofname, ".") > 0 then	fnext = mid(ofname, instrRev(ofname, "."))
				tstr = now()
				nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext 	    
				'nfname = replace(nfname, "/site/coa", "")					
				xup.Form(Form.Name).SaveAs apath & nfname, True								
				sql1 = sql1 & form.Name & " = '" & HTUploadPath & nfname & "',"
			end if	
		next
		sql1 = "UPDATE CuDtGeneric SET " & left(sql1, len(sql1) - 1) & " WHERE icuitem = " & request.querystring("icuitem")
		sql2 = "UPDATE Pedia SET " & left(sql2, len(sql2) - 1) & " WHERE gicuitem = " & request.querystring("icuitem")
		conn.execute(sql1 & ";" & sql2)
		
		'---活動期間才加分---
		Dim actFlag : actFlag = false
		sql = "SELECT * FROM Activity WHERE (ActivityId = 'pedia') AND (GETDATE() BETWEEN ActivityStartTime AND ActivityEndTime)"
		set rs = conn.execute(sql)
		if not rs.eof then
			actFlag = true		
		end if
		rs.close
		set rs = nothing
		
		'---審核不通過變通過---
		if oldpublic = "N" and newpublic = "Y" then
										
			if actFlag then							
				sql = "UPDATE ActivityPediaMember SET commendCount = commendCount + 1 WHERE memberId = '" & memberId & "'"
				conn.execute(sql)
			end if
			
			sql = "SELECT * FROM Pedia WHERE parentIcuitem = " & request.querystring("icuitem") & " AND xStatus = 'N'"
			set rs = conn.execute(sql)
			while not rs.eof
				sql = "UPDATE Pedia SET xStatus = 'Y' WHERE gicuitem = " & rs("gicuitem") 
				conn.execute(sql)
				if actFlag then
					sql = "UPDATE ActivityPediaMember SET commendAdditionalCount = commendAdditionalCount + 1 WHERE memberId = '" & rs("memberId") & "'"
					conn.execute(sql)
				end if
				rs.movenext
			wend
			rs.close
			set rs = nothing
			
		end if
		'---審核通過變成不通過---
		if oldpublic = "Y" and newpublic = "N" then
		
			if actFlag then
				sql = "UPDATE ActivityPediaMember SET commendCount = commendCount - 1 WHERE memberId = '" & memberId & "' AND commendCount > 0"
				conn.execute(sql)
			end if
			
			sql = "SELECT * FROM Pedia WHERE parentIcuitem = " & request.querystring("icuitem") & " AND xStatus = 'Y'"
			set rs = conn.execute(sql)
			while not rs.eof
				sql = "UPDATE Pedia SET xStatus = 'N' WHERE gicuitem = " & rs("gicuitem") 
				conn.execute(sql)
				if actFlag then
					sql = "UPDATE ActivityPediaMember SET commendAdditionalCount = commendAdditionalCount - 1 WHERE memberId = '" & rs("memberId") & "' AND commendAdditionalCount > 0"
					conn.execute(sql)
				end if
				rs.movenext
			wend
			rs.close
			set rs = nothing
			
		end if
		
	End sub 
%>	
<%
	apath = server.mappath(HTUploadPath) & "\"

	if request.querystring("phase")="edit" or session("BatchDphase") = "edit" then
		Set xup = Server.CreateObject("TABS.Upload")
	else
		Set xup = Server.CreateObject("TABS.Upload")
		xup.codepage=65001
		xup.Start apath
	end if

	function xUpForm(xvar)
		xUpForm = xup.form(xvar)
	end function
	    
	if xUpForm("submitTask") = "UPDATE" then
		
		errMsg = ""		
		if errMsg <> "" then
			EditInBothCase()
		else
			doUpdateDB()
			showDoneBox("資料更新成功！")
		end if
			
	elseif xUpForm("submitTask") = "DELETE" then
			
		icuitem = request.queryString("iCuItem")
		'---將排行榜的點數扣除---
		Dim oldstatus : oldstatus = ""		
		Dim memberId : memberId = ""
		sql = "SELECT fCTUPublic, xStatus, memberId FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.icuitem = Pedia.gicuitem WHERE gicuitem = " & request.querystring("icuitem")
		set frs = conn.execute(sql)
		if not frs.eof then
			oldstatus = frs("fCTUPublic")
			memberId = frs("memberId")
		end if
		frs.close
		set frs = nothing	
		
		'---活動期間才加分---
		Dim actFlag : actFlag = false
		sql = "SELECT * FROM Activity WHERE (ActivityId = 'pedia') AND (GETDATE() BETWEEN ActivityStartTime AND ActivityEndTime)"
		set rs = conn.execute(sql)
		if not rs.eof then
			actFlag = true		
		end if
		rs.close
		set rs = nothing
		
		'----刪除主表---將狀態由Y改成D---
		sql = "UPDATE Pedia SET xStatus = 'D' WHERE gicuitem = " & pkStr(icuitem, "")		
		conn.execute(sql)
		
		if oldstatus = "Y" then
				
			if actFlag then	
				sql = "UPDATE ActivityPediaMember SET commendCount = commendCount - 1 WHERE memberId = '" & memberId & "' AND commendCount > 0"
				conn.execute(sql)
			end if
			
			sql = "SELECT gicuitem, memberId FROM Pedia WHERE xStatus = 'Y' AND parentIcuitem = " & icuitem
			set rs = conn.execute(sql)
			while not rs.eof
				'---刪除補充解釋文章---
				sql1 = "UPDATE pedia SET xStatus = 'D' WHERE gicuitem = " & rs("gicuitem")
				conn.execute(sql1)		
				if actFlag then
					sql2 = "UPDATE ActivityPediaMember SET commendAdditionalCount = commendAdditionalCount - 1 WHERE memberId = '" & rs("memberId") & "' AND commendAdditionalCount > 0"
					conn.execute(sql2)
				end if
				rs.movenext
			wend
		end if					
		'------------------------
		showDoneBox("資料刪除成功！")	
	else		
		if errMsg <> "" then	showErrBox()		
		showForm()		
	end if
%>
<% 
Sub showForm() 

	sql = "SELECT *, CONVERT(varchar, xPostDate, 111) AS xDate FROM CuDtGeneric INNER JOIN Pedia ON CuDtGeneric.icuitem = Pedia.gicuitem "
	sql = sql & "LEFT JOIN Member ON Pedia.memberId = Member.account "
	sql = sql & "WHERE CuDtGeneric.icuitem = " & request.queryString("iCuItem")

	Set rs = conn.execute(sql)
	if rs.eof then
		showDoneBox("無此資料!")
	else
	
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="FuncName">
	<h1>資料管理／知識小百科</h1><font size=2>【目錄樹節點: 知識小百科】</font>
	<div id="Nav">	  
	  <a href="/Pedia/PediaAdd.asp?phase=add">新增</a>&nbsp;
	  <a href="/Pedia/PediaAdditionalList.asp?icuitem=<%=request.querystring("icuitem")%>">補充管理</a>&nbsp;
	  <a href="/Pedia/PediaList.asp">回條列</a>	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">	       
	單元資料維護&nbsp;
  <font size=2>【主題單元:知識小百科】</font>
</div>

<form id="Form1" method="POST" name="reg" action="" ENCTYPE="MULTIPART/FORM-DATA">
	<INPUT TYPE=hidden name=submitTask value="">	
	<INPUT TYPE=hidden name="iEditor" value="<%=session("userId")%>">	
	<table cellspacing="0">
	<TR>
		<TD class="Label" align="right"><span class="Must">*</span>詞目</TD>
		<TD class="eTableContent"><input name="sTitle" size="50" value="<%=rs("sTitle")%>">
		<%
			if rs("path") <> "" then
				xurl = session("myWWWSiteURL") & rs("path") 
		%>
			<a href="<%=xurl%>" target="_blank">來源連結</a>
		<% end if %>
		</TD></TR><TR>
		<TD class="Label" align="right">釋義</TD>
		<TD class="eTableContent"><textarea name="xBody" rows="8" cols="60"" title="內容、本文、網頁、說明""><%=rs("xBody")%></textarea></TD>
	</TR>
	<TR>
		<TD class="Label" align="right">英文詞目</TD>
		<TD class="eTableContent"><input name="engTitle" size="50" title="" value="<%=rs("engTitle")%>"></TD>
	</TR>
	<TR>
		<TD class="Label" align="right">學名</TD>
		<TD class="eTableContent"><input name="formalName" size="50" title="" value="<%=rs("formalName")%>"></TD>
	</TR>
	<TR>
		<TD class="Label" align="right">俗名</TD>
		<TD class="eTableContent"><input name="localName" size="50" title="" value="<%=rs("localName")%>"></TD>
	</TR>
	<TR>
		<TD class="Label" align="right"><span class="Must">*</span>是否開放補充</TD>
		<TD class="eTableContent">
			<Select name="vGroup">
				<option value="" <% if rs("vGroup") = "" Then %> selected <% end if %> >請選擇</option>
				<option value="Y" <% if rs("vGroup") = "Y" Then %> selected <% end if %> >是</option>
				<option value="N" <% if rs("vGroup") = "N" Then %> selected <% end if %> >否</option>
			</select>
		</TD>
	</TR>
	<TR>
		<TD class="Label" align="right">分類</TD>
		<TD class="eTableContent">
			<Select name="topCat">
			<option value="" <% if rs("topCat") = "" Then %> selected <% end if %> >請選擇</option>
			<%
				sql = "SELECT * FROM CodeMain WHERE codeMetaID = 'pediacata' ORDER BY mSortValue "
				Set cmrs = conn.execute(sql)
				while not cmrs.eof 
					if cmrs("mCode") = rs("topCat") then
						response.write "<option value=""" & cmrs("mCode") & """ selected>" & cmrs("mValue") & "</option>" & vbcrlf					
					else
						response.write "<option value=""" & cmrs("mCode") & """>" & cmrs("mValue") & "</option>" & vbcrlf
					end if
					cmrs.movenext
				wend
				set cmrs = nothing
			%>			
			</select>
		</TD>
	</TR>
	<TR>
		<TD class="Label" align="right">相關詞</TD>
		<TD class="eTableContent">
			<input name="xKeyword" title="" value="<%=rs("xKeyword")%>" size="50" readonly="true" class="rdonly">      
		</TD>
	</TR>
	<TR>
		<TD class="Label" align="right"><span class="Must">*</span>發布日期</TD>
		<TD class="eTableContent"><input name="xPostDate" size="8" readonly onclick="VBS: popCalendar 'xPostDate',''" value="<%=rs("xDate")%>"></TD>
	</TR>
	<TR>
		<TD class="Label" align="right"><span class="Must">*</span>審核狀態</TD>
		<TD class="eTableContent">
			<Select name="fCTUPublic">
				<option value="" <% if rs("fCTUPublic") = "" Then %> selected <% end if %> >請選擇</option>
				<option value="Y" <% if rs("fCTUPublic") = "Y" Then %> selected <% end if %> >通過</option>
				<option value="N" <% if rs("fCTUPublic") = "N" Then %> selected <% end if %> >不通過</option>
			</select>
		</TD>
	</TR>
	<TR>
		<TD class="Label" align="right">重要性</TD>
		<TD class="eTableContent"><input name="xImportant" size="2" title="(不重要) 0~99 (重要)" value="<%=rs("xImportant")%>"></TD>
	</TR>
	<TR>
		<TD class="Label" align="right">狀態</TD>
		<TD class="eTableContent"><input name="xStatus" size="2" title="" value="<%=rs("xStatus")%>" readonly="true" class="rdonly"></TD>
	</TR>
	<TR>
		<TD class="Label" align="right"><span class="Must">*</span>單位</TD>
		<TD class="eTableContent">
			<Select name="iDept">
			<%
				sqlCom = " SELECT D.deptId,D.deptName,D.parent,len(D.deptId)-1,D.seq," & _
								 " (Select Count(*) from Dept WHERE parent=D.deptId AND nodeKind='D') " & _
								 " FROM Dept AS D Where D.nodeKind='D' " & _
								 " AND D.deptId LIKE '" & session("deptId") & "%'" & _
								 " ORDER BY len(D.deptId), D.parent, D.seq"				
				set RSS = conn.execute(sqlCom)
				if not RSS.EOF then
					ARYDept = RSS.getrows(300)
					glastmsglevel = 0
					genlist 0, 0, 1, 0
			    expandfrom ARYDept(cid, 0), 0, 0
			end if
			%>
			</select>
		</TD>
	</TR>	
	<TR>
		<TD class="Label" align="right">發表會員</TD>
	  <TD class="eTableContent">
			<input name="memberId" class="rdonly" value="<%=rs("memberId")%>" size="10" readonly="true">
	    <input name="realname" class="rdonly" value="<%=rs("realname")%>" size="10" readonly="true">
	    <input name="nickname" class="rdonly" value="<%=rs("nickname")%>" size="10" readonly="true">
		</TD>
	</TR>
	<TR>
		<TD class="Label" align="right">圖檔</TD>
		<TD class="eTableContent">
		<% if Not IsNull(rs("xImgFile")) and rs("xImgFile") <> "" Then %>
			<img src="<%=rs("xImgFile")%>" /><br /><br />
		<% end if %>			
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td width="37">
					<input type="file" name="xImgFile">
				</td>
			</tr>
			</table>
		
		</TD>
	</TR>
	</TABLE>
	<object data="../inc/calendar.asp" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
	<INPUT TYPE=hidden name=CalendarTarget>
	<input name="button" type=button class="cbutton" onClick="formModSubmit(0)" value ="編修存檔">
	<input name="button" type=button class="cbutton" onClick="formModSubmit(1)" value ="儲存並新增相關詞">	
	<input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
  <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
  <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">      
</form>     
</body>
</html>

<script language=vbs>
	cvbCRLF = vbCRLF
	cTabchar = chr(9)

	Dim CanTarget
	Dim followCanTarget

	sub popCalendar(dateName,followName)        
		CanTarget=dateName
		followCanTarget=followName
		xdate = document.all(CanTarget).value
		if not isDate(xdate) then	xdate = date()
		document.all.calendar1.setDate xdate
	
		If document.all.calendar1.style.visibility="" Then           
			document.all.calendar1.style.visibility="hidden"        
		Else        
			ex=window.event.clientX
			ey=document.body.scrolltop+window.event.clientY+10
			if ex>520 then ex=520
			document.all.calendar1.style.pixelleft=ex-80
			document.all.calendar1.style.pixeltop=ey
			document.all.calendar1.style.visibility=""
		End If              
	end sub     

	sub calendar1_onscriptletevent(n,o) 
		document.all("CalendarTarget").value=n         
		document.all.calendar1.style.visibility="hidden"    
		if n <> "Cancle" then
			document.all(CanTarget).value=document.all.CalendarTarget.value
			'document.all("x"&CanTarget).value=document.all.CalendarTarget.value
			if followCanTarget<>"" then
				document.all(followCanTarget).value=document.all.CalendarTarget.value
				'document.all("x"&followCanTarget).value=document.all.CalendarTarget.value
			end if
		end if
	end sub   
 
  sub formModSubmit(btype)
		nMsg = "請務必填寫「{0}」，不得為空白！"
		lMsg = "「{0}」欄位長度最多為{1}！"
		dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
		iMsg = "「{0}」欄位應為數值！"
		pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
		
		if reg.sTitle.value = Empty Then
			MsgBox replace(nMsg,"{0}","標題"), 64, "Sorry!"
			reg.sTitle.focus
			exit sub
		end if
		if reg.vGroup.value = Empty Then
			MsgBox replace(nMsg,"{0}","是否開放補充"), 64, "Sorry!"
			reg.vGroup.focus
			exit sub
		end if
		if reg.xPostDate.value = Empty Then
			MsgBox replace(nMsg,"{0}","發佈日期"), 64, "Sorry!"
			reg.xPostDate.focus
			exit sub
		end if
		if reg.fCTUPublic.value = Empty Then
			MsgBox replace(nMsg,"{0}","是否公開"), 64, "Sorry!"
			reg.fCTUPublic.focus
			exit sub
		end if
		
		reg.submitTask.value = "UPDATE"
		reg.action = "PediaEdit.asp?icuitem=<%=request.querystring("icuitem")%>&btype=" & btype & "&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>"
		reg.Submit
	end sub
	
	sub formDelSubmit()
   	deleteStr = ""
   	deleteStr = deleteStr & "　確定刪除資料嗎？"
		chky=msgbox("注意！"& vbcrlf & vbcrlf &deleteStr& vbcrlf , 48+1, "請注意！！")
    if chky=vbok then			
			reg.submitTask.value = "DELETE"
			reg.action = "PediaEdit.asp?icuitem=<%=request.querystring("icuitem")%>&btype=0&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=request.querystring("pagesize")%>"
	    reg.Submit
    end If
  end sub
	
	sub resetForm 
		window.location.href = "/Pedia/PediaEdit.asp?icuitem=<%=request.querystring("icuitem")%>&phase=edit"
  end sub
</script>
<% 		
	end if 
	set rs = nothing
%>
<% End sub %>


 
