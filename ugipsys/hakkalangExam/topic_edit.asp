<%@ CodePage = 65001 %>
<!-- #include file = "exam.header.asp" -->
<%
Response.ContentType = "text/html; charset=utf-8"

iCuItem = Request.QueryString("iCuItem")

'取得腔調題目
Set oTopic = GetTopicByCuitem(connHakka, iCuItem)

'取得 CuDtGeneric
Set oCuDtGeneric = GetCuDtGeneric(connHakka, oTopic.GetField("CuItemId"))

sExamType = oCuDtGeneric.GetField("TopCat")
'sExamType = "9"

If sExamType = "3" Or sExamType = "5" Or sExamType = "7" Or sExamType = "9" Then
	On Error Resume Next
	
	oTopic.SetField "Options", GetAllOption(connHakka, oTopic.GetField("Id"))
	
	If Err.number <> 0 Then
		For I = 0 To 3
			Set oExamOption = New ExamOption
			oExamOption.SetField "TopicId", oTopic.GetField("Id")
			oExamOption.SetField "Title", ""
			oExamOption.SetField "Answer", ""
			oExamOption.SetField "Sort", I
			oTopic.AddOption oExamOption
		Next
	End If
	
	On Error Goto 0
End If

bolUpdateSuccessful = FALSE

If Request.ServerVariables("HTTP_METHOD") = "POST" And InStr(Request.ServerVariables("CONTENT_TYPE"), "multipart/form-data") > 0 Then
	'On Error Resume Next
	
	connHakka.BeginTrans
	
	Set oUpload = Server.CreateObject("TABS.Upload")
	oUpload.codepage = 65001
	oUpload.Start Server.MapPath("./")
	
	sCorrect = oUpload.Form("correct")
	sOptions = oUpload.Form("option")
	
	oTopic.SetField "Correct", sCorrect
	
	UpdateTopic connHakka, oTopic
	
	Select Case sExamType
		Case "3", "5", "9"
			'刪除所有選項
			DelAllOption connHakka, oTopic
			
			'重新新增選項
			aOptions = Split(sOptions, ",")
			For I = 0 To UBound(aOptions)
				If aOptions(I) <> "" Then
					Set oOption = New ExamOption
					oOption.SetField "TopicId", oTopic.GetField("Id")
					oOption.SetField "Title", aOptions(I)
					If CInt(iCorrect) = I Then
						oTopic.SetField "Answer", "Y"
					End If
					oOption.SetField "Sort", I
					CreateOption connHakka, oOption
				End If
			Next
		Case "7"
			'刪除所有選項
			DelAllOption connHakka, oTopic
			
			'重新新增選項
			On Error Resume Next
			aOptions = Split(sOptions, ",")
			For I = 0 To UBound(aOptions) Step 2
				If aOptions(I) <> "" And aOptions(I + 1) <> "" Then
					Set oOption = New ExamOption
					oOption.SetField "TopicId", oTopic.GetField("Id")
					oOption.SetField "Title", aOptions(I)
					oOption.SetField "Answer", aOptions(I + 1)
					oOption.SetField "Sort", I
					CreateOption connHakka, oOption
				End If
			Next
	End Select
	
	If Err.number <> 0 Then
		connHakka.RollbackTrans
		Response.Write Err.Description
	Else
		connHakka.CommitTrans
		bolUpdateSuccessful = TRUE
	End If
	
	'On Error Goto 0
	
End If
%>
<!-- #include file = "exam.footer.asp" -->
<%
If bolUpdateSuccessful Then
	Response.Redirect "../GipEdit/DsdXMLList.asp"
End If
%>
<html>
<head>
	<title></title>
	<link href="/css/form.css" rel="stylesheet" type="text/css" />
	<link href="/css/layout.css" rel="stylesheet" type="text/css" />
</head>
<script language="javascript">
<!--
function modSubmitForm() {
	var f = document.getElementById("Form1");
}

function delSubmitForm() {
	var f = document.getElementById("Form1");
	if(confirm("確定刪除此筆資料？")) {
		f.action = "topic_del.asp?et=<%=oTopic.GetField("Id")%>";
		f.submit();
	}
	
	return false;
}

function resetSubmitForm() {
	var f = document.getElementById("Form1");
	f.reset();
	clientInitForm();
}

function clientInitForm() {
	var f = document.getElementById("Form1");
}

function clientInitForm2() {
	var f = document.getElementById("Form1");
	
	switch('<%=sExamType%>') {
		case '1':
			setRadio(f.correct, '<%=oTopic.GetField("Correct")%>');
			break;
		case '3':
		case '5':
		case '9':
			<%
			Set Options = oTopic.GetField("Options")
			For I = 0 To Options.Count - 1
				Response.Write "addOption('" & Options(I).GetField("Title") & "');"
			Next
			%>
			setRadio(f.correct, '<%=oTopic.GetField("Correct")%>');
			break;
		case '7':
			<%
			Set Options = oTopic.GetField("Options")
			For I = 0 To Options.Count - 1
				Response.Write "addMatch('" & Options(I).GetField("Title") & "', '" & Options(I).GetField("Answer") & "');"
			Next
			%>
			break;
	}
}

function setRadio(radio, value) {
	if(radio.length != null) {
		for(i=0;i<radio.length;i++) {
			if(radio[i].value == value) radio[i].checked = true;
		}
	}
	else {
		if(radio[i].value == value) radio[i].checked = true;
	}
}

function addOption(value) {
	var examOptions = document.getElementById("examOptions");
	
	var oLi = document.createElement("li");
	
	var oText = document.createElement("<input type=\"text\" name=\"option\" />");
	oText.value = value;
	
	var oRadio = document.createElement("<input type=\"radio\" name=\"correct\" />");
	oRadio.value = examOptions.getElementsByTagName("li").length;
	
	oLi.appendChild(oText);
	oLi.appendChild(oRadio);
	examOptions.appendChild(oLi);
}

function addMatch(value1, value2) {
	var examOptions = document.getElementById("examOptions");
	
	var oLi = document.createElement("li");
	
	var oText1 = document.createElement("<input type=\"text\" name=\"option\" />");
	oText1.value = value1;
	
	var oText2 = document.createElement("<input type=\"text\" name=\"option\" />");
	oText2.value = value2;
	
	oLi.appendChild(oText1);
	oLi.appendChild(oText2);
	examOptions.appendChild(oLi);
}

function body_onload() {
	clientInitForm();
	clientInitForm2();
}
//-->
</script>
<body onload="body_onload();">
	<div id="FuncName">
		<h1>資料上稿／題庫選項編輯</h1>
		<div id="ClearFloat"></div>
	</div>
	
	<form action="" id="Form1" name="Form1" method="post" enctype="multipart/form-data">
		<input type="hidden" name="act" value="" />
		<table border="0" cellspacing="0">
			<% If sExamType = "1" Then %>
			<tr>
				<td class="Label" align="right">對錯</td>
				<td>
					<input type="radio" name="correct" value="Y" />對
					<input type="radio" name="correct" value="N" />錯
				</td>
			</tr>
			<% ElseIf sExamType = "3" Or sExamType = "5" Or sExamType = "9" Then %>
			<tr>
				<td class="Label" align="right">選項</td>
				<td>
					<ol id="examOptions" type="1"></ol>
					<input type="button" class="cbutton" name="btnNewOption" value="新增選項" onclick="addOption('');" />
				</td>
			</tr>
			<% ElseIf sExamType = "7" Then %>
			<tr>
				<td class="Label" align="right">配合</td>
				<td>
					<ol id="examOptions" type="1"></ol>
					<input type="button" class="cbutton" name="btnNewOption" value="新增選項" onclick="addMatch('', '');" />
				</td>
			</tr>
			<% End If %>
		</table>
		<input type="submit" value="編修存檔" class="cbutton" onclick="return modSubmitForm();" />
		<input type="button" value="重　填" class="cbutton" onclick="resetSubmitForm();" />
		<input type="button" value="回前頁" class="cbutton" onclick="history.back();" />
	</form>
</body>
</html>