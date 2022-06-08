<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="系統參數維護"
HTProgFunc="編修首頁設定"
HTUploadPath="/"
HTProgCode="GW1M90"
HTProgPrefix="myStyle" 
%>
<!--#Include VIRTUAL = "/inc/server.inc" -->
<!--#Include VIRTUAL = "/inc/checkGIPconfig.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>
<!--#INCLUDE VIRTUAL="/inc/dbutil.inc" -->
<!--#Include VIRTUAL = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction
taskLable="編輯" & HTProgCap
	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

function xUpForm(xvar)
	xUpForm = request.form(xvar)
end function

	xmp = request("xdmpID")
	
		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & xmp & ".xml"
'		response.write LoadXML & "<HR>"
		xv = htPageDom.load(LoadXML)
'		response.write xv & "<HR>"
  		if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  		end if


if xUpForm("submitTask") = "UPDATE" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("資料更新成功！")
	end if

else
	EditInBothCase()
end if


sub EditInBothCase

	showHTMLHead()
	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
'	response.write sqlcom
	showForm()
	initForm()
	showHTMLTail()
end sub


function qqRS(fldName)
	if request("submitTask")="" then
		xValue = nullText(htPageDom.selectSingleNode("SystemParameter/" & fldName))
	else
		xValue = ""
		if request("htx_"&fldName) <> "" then
			xValue = request("htx_"&fldName)
		end if
	end if
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
	end if
end function %>

<%Sub initForm() %>
    <script language=vbs>
      sub resetForm 
	       reg.reset()
	       clientInitForm
      end sub

     sub clientInitForm
     	reg.MenuTree.value = "<%=nullText(htPageDom.selectSingleNode("MpDataSet/MenuTree"))%>"
     	reg.MenuTree2.value = "<%=nullText(htPageDom.selectSingleNode("MpDataSet/MenuTree2"))%>"
     	initRadio "MpStyle" , "<%=nullText(htPageDom.selectSingleNode("MpDataSet/MpStyle"))%>"
<%
	if nullText(htPageDom.selectSingleNode("MpDataSet/ChangeHead"))<>"" then
		for each param in htPageDom.selectNodes("MpDataSet/ChangeHead") 
		'四季+節慶版型
			'set param=htPageDom.selectSingleNode("MpDataSet/ChangeHead")
			prefix2 = "htx_" & nullText(param.selectSingleNode("DataValue")) & "_"
			response.write "reg." & prefix2 & "DataValue.value= """  _
				& nullText(param.selectSingleNode("DataValue")) & """" & vbCRLF
		next
	end if
	'DataSet
	for each param in htPageDom.selectNodes("MpDataSet/DataSet") 
		prefix = "htx_" & nullText(param.selectSingleNode("DataLable"))&""&nullText(param.selectSingleNode("DataLableSeq")) & "_"
		response.write "reg." & prefix & "DataRemark.value= """  _
			& nullText(param.selectSingleNode("DataRemark")) & """" & vbCRLF
		response.write "reg." & prefix & "DataNode.value= """  _
			& nullText(param.selectSingleNode("DataNode")) & """" & vbCRLF
		response.write "reg." & prefix & "SqlTop.value= """  _
			& nullText(param.selectSingleNode("SqlTop")) & """" & vbCRLF
		response.write "reg." & prefix & "ContentStyle.value= """  _
			& nullText(param.selectSingleNode("ContentStyle")) & """" & vbCRLF
		response.write "reg." & prefix & "ContentLength.value= """  _
			& nullText(param.selectSingleNode("ContentLength")) & """" & vbCRLF
		response.write "reg." & prefix & "ContentData.value= """  _
			& nullText(param.selectSingleNode("ContentData")) & """" & vbCRLF
		response.write "reg." & prefix & "Israndom.value= """  _
			& nullText(param.selectSingleNode("Israndom")) & """" & vbCRLF
		response.write "reg." & prefix & "DataSrc.value= """  _
			& nullText(param.selectSingleNode("DataSrc")) & """" & vbCRLF
		response.write "reg." & prefix & "PicWidth.value= """  _
			& nullText(param.selectSingleNode("PicWidth")) & """" & vbCRLF
		response.write "reg." & prefix & "PicHeight.value= """  _
			& nullText(param.selectSingleNode("PicHeight")) & """" & vbCRLF
		response.write "reg." & prefix & "refDataBlock.value= """  _
			& nullText(param.selectSingleNode("refDataBlock")) & """" & vbCRLF
	next
%>
    end sub

    sub window_onLoad
         clientInitForm
    end sub
    
    sub initRadio(xname,value)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    end sub

 </script>	
<%End sub '---- initForm() ----%>


<%Sub showForm()  	'===================== Client Side Validation Put HERE =========== %>
    <script language="vbscript">
      sub formModSubmit()
    
'----- 用戶端Submit前的檢查碼放在這裡 ，如下例 not valid 時 exit sub 離開------------
'msg1 = "請務必填寫「客戶編號」，不得為空白！"
'msg2 = "請務必填寫「客戶名稱」，不得為空白！"
'
'  If reg.tfx_ClientName.value = Empty Then
'     MsgBox msg2, 64, "Sorry!"
'     settab 0
'     reg.tfx_ClientName.focus
'     Exit Sub
'  End if
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
  
<%
'	for each param in refModel.selectNodes("fieldList/field[formList='Y']") 
'		processValid param
'	next
%>
  
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub


</script>

<form id="Form1" name="reg" method="POST" action="">
  <INPUT TYPE=hidden name=submitTask value="" ID="Hidden1">
    <table cellspacing="0" ID="Table1">
      <tr>
        <td>
          <table cellspacing="0" ID="Table2" align="left">
            <tr>    
              <td align="right" class="Label"><span class="Must">*</span>目錄樹一：</td>     
              <td class="whitetablebg">
		<SELECT name=MenuTree size=1 ID="Select1">
		  <option value="">--無--</option>
<%
	sql = "SELECT * FROM CatTreeRoot"
	set RSx = conn.execute(sql)
	while not RSx.eof
		response.write "<OPTION VALUE=""" & RSx("CtRootID") & """>" & RSx("CtRootName") & "</OPTION>" & vbCRLF
		RSx.moveNext
	wend
%>
		</SELECT>
              </td>     
            </tr>
            <tr>    
              <td align="right" class="Label">目錄樹二：</td>     
              <td class="whitetablebg">
		<SELECT name=MenuTree2 size=1 ID="Select2">
		  <option value="">--無--</option>
<%
	sql = "SELECT * FROM CatTreeRoot"
	set RSx = conn.execute(sql)
	while not RSx.eof
		response.write "<OPTION VALUE=""" & RSx("CtRootID") & """>" & RSx("CtRootName") & "</OPTION>" & vbCRLF
		RSx.moveNext
	wend
%>
		</SELECT>
              </td>     
            </tr>
            <tr>    
              <td align="right" class="Label"><span class="Must">*</span>版面樣式：</td>     
              <td class="whitetablebg">
                <table>
<%
		SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId=N'xStyle' Order by msortValue"
			SET RSS=conn.execute(SQL)
			i = 0
			mod_num = 3
			While not RSS.EOF
			i = i + 1
			if i mod mod_num = 1 then	response.write "              <tr>" & vbcrlf
%>
                  <td>
		    <INPUT TYPE="radio" NAME="MpStyle" value="<%=RSS(0)%>" ID="Radio1"><%=RSS(1)%></br>
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		    <IMG SRC="/site/<%=session("mySiteID")%>/mpStyle/<%=RSS(0)%>.gif" width="150" align="top">
                  </td>
			<BR/>
			<%	
			if i mod mod_num = 0 then	response.write "              </tr>" & vbcrlf
			RSS.movenext
			wend
			%>
<% if i=1 then %>
		  <td>
		    <INPUT TYPE="radio" NAME="MpStyle" value="" ID="Radio2">
		  </td>
<% end if %>
		</table>
              </td>
            </tr>
  <%if nullText(htPageDom.selectSingleNode("MpDataSet/ChangeHead"))<>"" then%>
	  <table cellspacing="0" border="1" ID="Table3">
	    <tr align="center">
	      <th>款式</th><th>代碼數值</th><th>說明事項</th>
	    </tr>
		<%
		for each param in htPageDom.selectNodes("MpDataSet/ChangeHead")
			response.write "<tr align=""center"">"
			response.write "<td class=""Label"" align=""right"">&lt;" & nullText(param.selectSingleNode("DataLable")) & "&gt;</td>"
			prefix2 = "htx_" & nullText(param.selectSingleNode("DataValue")) & "_"
			response.write "<td><INPUT name=""" & prefix2 & "DataValue"" size=10></td>" 
			response.write "<td>" & nullText(param.selectSingleNode("Description")) & "</td>"
			response.write "</tr>"
		next
		%>
	  </table>
  <%end if%>
            <tr>
              <td class="whitetablebg" colspan="2"><!-- class="eTableContent" -->
                <table cellspacing="0" border="1" ID="Table4">
                  <tr align="center">
                    <th>Tag</th>
                    <th>區塊名稱</th>
                    <th>目錄樹節點</th>
                    <th>則數</th>
                    <th>字數</th>
                    <th>樣式</th>
                    <th>內頁</th>
                    <th>隨機</th>
                    <th>來源</th>
                    <th>相關單元</th>
                    <th>圖寬</th>
                    <th>圖高</th>
                  </tr>
<%	
	for each param in htPageDom.selectNodes("MpDataSet/DataSet")
		response.write "</TR>"
		response.write "<TD align=""left"">"
		response.write "&lt;" & nullText(param.selectSingleNode("DataLable")) & "&gt;</TD>"
		prefix = "htx_" & nullText(param.selectSingleNode("DataLable"))&""&nullText(param.selectSingleNode("DataLableSeq")) & "_"
		response.write "<TD class=""eTableContent"">"
		response.write "<INPUT name=""" & prefix & "DataRemark"" size=10>" 
		response.write  "</TD>"
		response.write "<TD class=""eTableContent"">"
		response.write "<INPUT name=""" & prefix & "DataNode"" size=2>" 
		response.write "</TD>"
		response.write "<TD class=""eTableContent"">"
		response.write "<INPUT name=""" & prefix & "SqlTop"" size=1>" 
		response.write "</TD>"
		response.write "<TD class=""eTableContent"">"
		response.write "<INPUT name=""" & prefix & "ContentLength"" size=1>" 
		response.write "</TD>"
		response.write "<TD class=""eTableContent"">"
		response.write "<SELECT name=""" & prefix & "ContentStyle"" >"

			SQLCom = "select * from CodeMain Where codeMetaId = 'ContentStyle'"
			set RS = conn.execute(SqlCom)
			Do while not RS.EOF
				response.write "<OPTION value='" & RS("mcode") & "'>" & RS("mvalue") & "</option>"
				RS.MoveNext
			Loop
		response.write "</SELECT>" 
		response.write "</TD>"
		response.write "<TD class=""eTableContent"">"
		response.write "<SELECT name=""" & prefix & "ContentData"" >"

			SQLCom = "select * from CodeMain Where codeMetaId = 'boolYN'"
			set RS = conn.execute(SqlCom)
			Do while not RS.EOF
				response.write "<OPTION value='" & RS("mcode") & "'>" & RS("mvalue") & "</option>"
				RS.MoveNext
			Loop
		response.write "</SELECT>" 
		response.write "</TD>"
		response.write "<TD class=""eTableContent"">"
		response.write "<SELECT name=""" & prefix & "Israndom"" >"

			SQLCom = "select * from CodeMain Where codeMetaId = 'boolYN'"
			set RS = conn.execute(SqlCom)
			Do while not RS.EOF
				response.write "<OPTION value='" & RS("mcode") & "'>" & RS("mvalue") & "</option>"
				RS.MoveNext
			Loop
		response.write "</SELECT>" 
		response.write "</TD>"
		response.write "<TD class=""eTableContent"">"
		response.write "<INPUT name=""" & prefix & "DataSrc"" size=2>" 
		response.write "</TD>"
		response.write "<TD class=""eTableContent"">"
		response.write "<INPUT name=""" & prefix & "refDataBlock"" size=1>" 
		response.write "</TD>"
		response.write "<TD class=""eTableContent"">"
		response.write "<INPUT name=""" & prefix & "PicWidth"" size=1>" 
		response.write "</TD>"
		response.write "<TD class=""eTableContent"">"
		response.write "<INPUT name=""" & prefix & "PicHeight"" size=1>" 
		response.write "</TD>"
		response.write "</TR>"
	next
%>
                </table>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </TABLE>
<object data="../inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget ID="Hidden2">
   <% if (HTProgRight and 8)=8 then %><input type="button" value="編修存檔" name="Enter" class="cbutton" OnClick="formModSubmit()" ID="Button1"><% End IF %> 
   <input type="button" class=cbutton value="清除重填" onClick="resetForm()" ID="Button2" NAME="Button2">
</form>     

<div id="Explain">
	<h1>說明</h1>
	<ul>
		<li><span class="Must">*</span>為必要欄位</li>
	</ul>
</div>

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
    <body>
<div id="FuncName">
	<h1><%=Title%></h1>
	<div id="Nav">
	<% if (HTProgRight and 8)=8 then %>
		<a href="mpStyle.asp?mp=1">首頁設定</a>
	<% End IF %>
	    <a href="VBScript: history.back">回上一頁</a>
	</div>
	<div id="ClearFloat"></div>
</div>
<% 
xdmpID = request("xdmpID") 
sqlx = "SELECT xdmpName FROM XdmpList WHERE xdmpID= "& xdmpID
set RSx = conn.Execute(sqlx)
%>
<div id="FormName"><%=RSx("xdmpName")%>/<%=HTProgFunc%></div>

<%End sub '--- showHTMLHead() ------%>


<%Sub ShowHTMLTail() %>     
</body>
</html>
<%End sub '--- showHTMLTail() ------%>


<%Sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
'  			window.history.back
	</script>
<%
    End sub '---- showErrBox() ----

Sub checkDBValid()	'===================== Server Side Validation Put HERE =================

'---- 後端資料驗證檢查程式碼放在這裡	，如下例，有錯時設 errMsg="xxx" 並 exit sub ------
'	if Request("tfx_TaxID") <> "" Then
'	  SQL = "Select * From Client Where TaxID = N'"& Request("tfx_TaxID") &"'"
'	  set RSreg = conn.Execute(SQL)
'	  if not RSreg.EOF then
'		if trim(RSreg("ClientId")) <> request("pfx_ClientID") Then
'			errMsg = "「統一編號」重複!!請重新輸入客戶統一編號!"
'			exit sub
'		end if
'	  end if
'	end if

End sub '---- checkDBValid() ----

Sub doUpdateDB()
	htPageDom.selectSingleNode("MpDataSet/MenuTree").text = request("MenuTree")
	htPageDom.selectSingleNode("MpDataSet/MenuTree1").text = request("MenuTree")
	htPageDom.selectSingleNode("MpDataSet/MenuTree2").text = xUpForm("MenuTree2")
	htPageDom.selectSingleNode("MpDataSet/MpStyle").text = request("MpStyle")
	'四季版型
  	if nullText(htPageDom.selectSingleNode("MpDataSet/ChangeHead"))<>"" then
		for each param in htPageDom.selectNodes("MpDataSet/ChangeHead")
			prefix2 = "htx_" & nullText(param.selectSingleNode("DataValue")) & "_"
			param.selectSingleNode("DataValue").text = request(prefix2 & "DataValue")
		next
	end if
	'xDataSet
	for each param in htPageDom.selectNodes("MpDataSet/DataSet")
		prefix = "htx_" & nullText(param.selectSingleNode("DataLable"))&""&nullText(param.selectSingleNode("DataLableSeq")) & "_"
		param.selectSingleNode("DataRemark").text = request(prefix & "DataRemark")
		param.selectSingleNode("DataNode").text = request(prefix & "DataNode")
		param.selectSingleNode("SqlTop").text = request(prefix & "SqlTop")
		param.selectSingleNode("ContentLength").text = request(prefix & "ContentLength")
		param.selectSingleNode("ContentStyle").text = request(prefix & "ContentStyle")
		param.selectSingleNode("ContentData").text = request(prefix & "ContentData")
		param.selectSingleNode("Israndom").text = request(prefix & "Israndom")
		param.selectSingleNode("DataSrc").text = request(prefix & "DataSrc")
		param.selectSingleNode("PicWidth").text = request(prefix & "PicWidth")
		param.selectSingleNode("PicHeight").text = request(prefix & "PicHeight")
		param.selectSingleNode("refDataBlock").text = request(prefix & "refDataBlock")
		if request(prefix & "Israndom") = "Y" then
			param.selectSingleNode("SqlOrderBy").text = "ra"
		else
			param.selectSingleNode("SqlOrderBy").text = "xImportant, xPostDate DESC"
		end if
	next
	
	htPageDom.save(LoadXML)
	
	if xmp = "1" and request("MpStyle")="webey1" then
		doUpdateDB2()
		doUpdateDB3()
	end if

End sub '---- doUpdateDB()

Sub doUpdateDB2()
	htPageDom.selectSingleNode("MpDataSet/MenuTree").text = request("MenuTree")
	htPageDom.selectSingleNode("MpDataSet/MenuTree1").text = request("MenuTree")
	htPageDom.selectSingleNode("MpDataSet/MenuTree2").text = xUpForm("MenuTree2")
	htPageDom.selectSingleNode("MpDataSet/MpStyle").text = "webey2"
	'四季版型
  	if nullText(htPageDom.selectSingleNode("MpDataSet/ChangeHead"))<>"" then
		for each param in htPageDom.selectNodes("MpDataSet/ChangeHead")
			prefix2 = "htx_" & nullText(param.selectSingleNode("DataValue")) & "_"
			param.selectSingleNode("DataValue").text = request(prefix2 & "DataValue")
		next
	end if
	'xDataSet
	for each param in htPageDom.selectNodes("MpDataSet/DataSet")
		prefix = "htx_" & nullText(param.selectSingleNode("DataLable"))&""&nullText(param.selectSingleNode("DataLableSeq")) & "_"
		param.selectSingleNode("DataRemark").text = request(prefix & "DataRemark")
		param.selectSingleNode("DataNode").text = request(prefix & "DataNode")
		param.selectSingleNode("SqlTop").text = request(prefix & "SqlTop")
		param.selectSingleNode("ContentLength").text = request(prefix & "ContentLength")
		param.selectSingleNode("ContentStyle").text = request(prefix & "ContentStyle")
		param.selectSingleNode("ContentData").text = request(prefix & "ContentData")
		param.selectSingleNode("Israndom").text = request(prefix & "Israndom")
		param.selectSingleNode("DataSrc").text = request(prefix & "DataSrc")
		param.selectSingleNode("PicWidth").text = request(prefix & "PicWidth")
		param.selectSingleNode("PicHeight").text = request(prefix & "PicHeight")
		param.selectSingleNode("refDataBlock").text = request(prefix & "refDataBlock")
		if request(prefix & "Israndom") = "Y" then
			param.selectSingleNode("SqlOrderBy").text = "ra"
		else
			param.selectSingleNode("SqlOrderBy").text = "xImportant, xPostDate DESC"
		end if
	next
	
	htPageDom.save(server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp2.xml")
	

End sub '---- doUpdateDB()

Sub doUpdateDB3()
	htPageDom.selectSingleNode("MpDataSet/MenuTree").text = request("MenuTree")
	htPageDom.selectSingleNode("MpDataSet/MenuTree1").text = request("MenuTree")
	htPageDom.selectSingleNode("MpDataSet/MenuTree2").text = xUpForm("MenuTree2")
	htPageDom.selectSingleNode("MpDataSet/MpStyle").text = "webey3"
	'四季版型
  	if nullText(htPageDom.selectSingleNode("MpDataSet/ChangeHead"))<>"" then
		for each param in htPageDom.selectNodes("MpDataSet/ChangeHead")
			prefix2 = "htx_" & nullText(param.selectSingleNode("DataValue")) & "_"
			param.selectSingleNode("DataValue").text = request(prefix2 & "DataValue")
		next
	end if
	'xDataSet
	for each param in htPageDom.selectNodes("MpDataSet/DataSet")
		prefix = "htx_" & nullText(param.selectSingleNode("DataLable"))&""&nullText(param.selectSingleNode("DataLableSeq")) & "_"
		param.selectSingleNode("DataRemark").text = request(prefix & "DataRemark")
		param.selectSingleNode("DataNode").text = request(prefix & "DataNode")
		param.selectSingleNode("SqlTop").text = request(prefix & "SqlTop")
		param.selectSingleNode("ContentLength").text = request(prefix & "ContentLength")
		param.selectSingleNode("ContentStyle").text = request(prefix & "ContentStyle")
		param.selectSingleNode("ContentData").text = request(prefix & "ContentData")
		param.selectSingleNode("Israndom").text = request(prefix & "Israndom")
		param.selectSingleNode("DataSrc").text = request(prefix & "DataSrc")
		param.selectSingleNode("PicWidth").text = request(prefix & "PicWidth")
		param.selectSingleNode("PicHeight").text = request(prefix & "PicHeight")
		param.selectSingleNode("refDataBlock").text = request(prefix & "refDataBlock")
		if request(prefix & "Israndom") = "Y" then
			param.selectSingleNode("SqlOrderBy").text = "ra"
		else
			param.selectSingleNode("SqlOrderBy").text = "xImportant, xPostDate DESC"
		end if
	next
	
	htPageDom.save(server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp3.xml")
	

End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) %>
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
	    document.location.href="xdmpList.asp"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
<script language=vbs>
Dim CanTarget

sub popCalendar(dateName)        
 	CanTarget=dateName
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
	end if
end sub   
</script>
y 