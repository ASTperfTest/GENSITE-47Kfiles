<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="PEDIA01"
HTProgPrefix="Pedia" 
Dim iBaseDSDId : iBaseDSDId = session("PediaBaseDSD")
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
<% Sub showDoneBox(lMsg, newId) %>
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
					document.location.href="<%=HTprogPrefix%>List.asp?keep=Y"
				<% else %>
					document.location.href="<%=HTprogPrefix%>Keyword.asp?icuitem=" & <%=newId%> & "&phase=edit"
				<% end if %>
			</script>
    </body>
  </html>
<% End sub %>
<%
	Function doUpdateDB()
	
		sql1 = ""	: sql11 = ""
		sql2 = "" :	sql22 = ""
		
		sql1 = "ibaseDSD,iCTUnit,siteId,"
		sql11 = iBaseDSDId & "," & iCtUnitId & ",N'4',"

		for each form in xup.Form		
			if form.Name <> "xImgFile" and form.Name <> "CalendarTarget" and form.Name <> "submitTask" and form.Name <> "xStatus" Then
				if form.Name = "engTitle" or form.Name = "formalName" or form.Name = "localName" then
					sql2 = sql2 & form.Name & ","
					sql22 = sql22 & "'" & form & "',"
				else 
					if form.Name = "xImportant" then
						important = form
						if important = "" then important = "0"
						sql1 = sql1 & form.Name & ","
						sql11 = sql11 & important & ","
					elseif form = "null" then
						sql1 = sql1 & form.Name & ","
						sql11 = sql11 & "'',"
					else
						sql1 = sql1 & form.Name & ","
						sql11 = sql11 & mypkStr(form, form.Name, "") & ","
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
				sql1 = sql1 & form.Name & ","
				sql11 = sql11 & "'" & HTUploadPath & nfname & "',"
			end if	
		next		
		sql = "INSERT INTO CuDtGeneric(" & left(sql1, len(sql1) - 1) & ") VALUES(" & left(sql11, len(sql11) - 1) & ")"		
		sql = "set nocount on;" & sql & "; select @@IDENTITY as NewID"				
		Set rs = conn.Execute(sql)
		xNewIdentity = rs(0)		
		Set rs = nothing
		sql2 = "gicuitem,xStatus,memberId,commendTime," & sql2
		sql22 = xNewIdentity & ",N'Y','" & session("userId") & "',GETDATE()," & sql22
		sql = "INSERT INTO Pedia(" & left(sql2, len(sql2) - 1) & ") VALUES(" & left(sql22, len(sql22) - 1) & ")"				
		conn.execute(sql)
		
		doUpdateDB = xNewIdentity
	End Function 
%>	
<%
	Dim hyftdGIPStr
	Dim pKey
	Dim RSreg
	Dim formFunction
	Dim allModel2
	Dim xshowTypeStr,xshowType
	Dim xRef5Count,xref5YN
	Dim orgInputType
	Dim xMMOFolderID,MMOPath,xFTPIPMMO,xFTPPortMMO,xFTPIDMMO,xFTPPWDMMO,MMOFTPfilePath,MMOCount
	MMOCount=0

	taskLable="編輯" & HTProgCap
	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"
		
	apath = server.mappath(HTUploadPath) & "\"

	if request.querystring("phase")="add" or session("BatchDphase") = "add" then
		Set xup = Server.CreateObject("TABS.Upload")
	else
		Set xup = Server.CreateObject("TABS.Upload")
		xup.codepage=65001
		xup.Start apath
	end if

	function xUpForm(xvar)
		xUpForm = xup.form(xvar)
	end function
	    
	if xUpForm("submitTask") = "ADD" then
		
		errMsg = ""		
		if errMsg <> "" then
			EditInBothCase()
		else
			xNewId = doUpdateDB()
			showDoneBox "資料新增成功！", xNewId
		end if			
	else		
		if errMsg <> "" then	showErrBox()
		formFunction = "edit"	
		showForm()		
	end if
	session("BatchDicuitem") = ""
	session("BatchDphase") = ""
	session("BatchDsubmitTask") = ""	

%>
<% 
Sub showForm() 
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
	<h1></h1><font size=2>【目錄樹節點:知識小百科】</font>
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
		<TD class="eTableContent"><input name="sTitle" size="50" value=""></TD></TR><TR>
		<TD class="Label" align="right">釋義</TD>
		<TD class="eTableContent"><textarea name="xBody" rows="8" cols="60"" title="內容、本文、網頁、說明""></textarea></TD>
	</TR>
	<TR>
		<TD class="Label" align="right">英文詞目</TD>
		<TD class="eTableContent"><input name="engTitle" size="50" title="" value=""</TD>
	</TR>
	<TR>
		<TD class="Label" align="right">學名</TD>
		<TD class="eTableContent"><input name="formalName" size="50" title="" value=""></TD>
	</TR>
	<TR>
		<TD class="Label" align="right">俗名</TD>
		<TD class="eTableContent"><input name="localName" size="50" title="" value=""></TD>
	</TR>
	<TR>
		<TD class="Label" align="right"><span class="Must">*</span>是否開放補充</TD>
		<TD class="eTableContent">
			<Select name="vGroup">
				<option value="">請選擇</option>
				<option value="Y" selected>是</option>
				<option value="N">否</option>
			</select>
		</TD>
	</TR>
	<TR>
		<TD class="Label" align="right">分類</TD>
		<TD class="eTableContent">
			<Select name="topCat">
			<option value="" selected>請選擇</option>
			<%
				sql = "SELECT * FROM CodeMain WHERE codeMetaID = 'pediacata' ORDER BY mSortValue "
				Set cmrs = conn.execute(sql)
				while not cmrs.eof 					
					response.write "<option value=""" & cmrs("mCode") & """>" & cmrs("mValue") & "</option>" & vbcrlf										
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
			<input name="xKeyword" title="" value="" size="50" readonly="true" class="rdonly">      
		</TD>
	</TR>
	<TR>
		<TD class="Label" align="right"><span class="Must">*</span>發布日期</TD>
		<TD class="eTableContent"><input name="xPostDate" size="8" readonly onclick="VBS: popCalendar 'xPostDate',''" value="<%=date%>"></TD>
	</TR>
	<TR>
		<TD class="Label" align="right"><span class="Must">*</span>是否公開</TD>
		<TD class="eTableContent">
			<Select name="fCTUPublic">
				<option value="">請選擇</option>
				<option value="Y" selected>公開</option>
				<option value="N">不公開</option>
			</select>
		</TD>
	</TR>
	<TR>
		<TD class="Label" align="right">重要性</TD>
		<TD class="eTableContent"><input name="xImportant" size="2" title="(不重要) 0~99 (重要)" value=""></TD>
	</TR>
	<TR>
		<TD class="Label" align="right">狀態</TD>
		<TD class="eTableContent"><input name="xStatus" size="2" title="" value="" readonly="true" class="rdonly"></TD>
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
		<TD class="Label" align="right">圖檔</TD>
		<TD class="eTableContent">
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
	<input name="button" type=button class="cbutton" onClick="formModSubmit(0)" value ="儲存">	
	<input name="button" type=button class="cbutton" onClick="formModSubmit(1)" value ="儲存並新增相關詞">	
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
				
		reg.submitTask.value = "ADD"
		reg.action = "PediaAdd.asp?btype=" & btype
		reg.Submit
	end sub
	
	sub resetForm 
		window.location.href = "/Pedia/PediaAdd.asp?phase=add"
  end sub
</script>
<% End sub %>


 
