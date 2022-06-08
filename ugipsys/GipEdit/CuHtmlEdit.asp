<%@ CodePage = 65001 %>
<%  CodePage=65001
Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath="/"
HTProgCode="GC1AP1"
HTProgPrefix="DsdXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

function htmlMessage(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	htmlMessage=""
  	exit function
  elseif instr(1,xs,"<P",1)>0 or instr(1,xs,"<BR",1)>0 or instr(1,xs,"<td",1)>0 then
  	htmlMessage=xs
  	exit function
  end if
  	xs = replace(xs,vbCRLF&vbCRLF,"<P>")
  	xs = replace(xs,vbCRLF,"<BR>")
  	htmlMessage = replace(xs,chr(10),"<BR>")
end function


	siteStr = "http://" & request.serverVariables("SERVER_NAME") 
	myURL = "http://" & request.serverVariables("SERVER_NAME") & request.serverVariables("URL") & "?" & request.serverVariables("QUERY_STRING")
'	response.write 	myURL & "<HR>"
	for each x in request.serverVariables
'		response.write x & "=>" & request(x) & "<BR>"
	next
	
	if request.form("submitTask") = "UPDATE" then
		icuitem = request.form("htx_icuitem")
		xbody = request("htx_xbody")
'		response.write xbody & "<HR>"
		xbody = replace(xbody, myURL, "")
		xbody = replace(xbody, siteStr& "/GipEdit/", "")
		xbody = replace(xbody, siteStr, "")
		xbody = replace(xbody, "<A href=""http://", "<A target=""_nwMof"" href=""http://")
		xbody = replace(xbody, "<A href=""/public/", "<A target=""_nwMof"" href=""/public/")
'		xbody = replace(xbody, siteStr, "")
		
'		response.write xbody & "<HR>"
'		response.end
		if isNumeric(icuitem) then
			sql = "UPDATE CuDtGeneric SET xbody = " & pkStr(xbody,"") _
				& " WHERE icuitem=" & icuitem
'			response.write sql & "<HR>"
'			response.end
			conn.execute sql
			response.redirect "DsdXMLEdit.asp?phase=edit&icuitem=" & icuitem
		
		
		else
			response.end
		end if
	end if


  	set htPageDom = session("codeXMLSpec")
  	set refModel = htPageDom.selectSingleNode("//dsTable")
  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")

	sqlCom = "SELECT ghtx.* FROM CuDtGeneric AS ghtx "_
		& " WHERE ghtx.icuitem=" & pkStr(request.queryString("icuitem"),"")
	Set RSreg = Conn.execute(sqlcom)
	pKey = "icuitem=" & RSreg("icuitem")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>HTML 編輯器</title>
<link href="/css/editor.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<HTML><HEAD><TITLE>HTML 編輯器</TITLE>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

<SCRIPT>
function BtnOver(btn){
btn.style.borderTopColor="#efece8";
btn.style.borderBottomColor="#888070";
btn.style.borderLeftColor="#efece8";
btn.style.borderRightColor="#888070";
}
function BtnClick(btn){
btn.style.borderTopColor="#888070";
btn.style.borderBottomColor="#efece8";
btn.style.borderLeftColor="#888070";
btn.style.borderRightColor="#efece8";
}
function BtnOut(btn){
btn.style.borderColor="#d9cec4";
}
function doCut(){
doc.execCommand('Cut');
Editor.focus();
}
function doCopy(){
doc.execCommand('Copy');
doc.selection.focus();
//Editor.focus();
}
function doPaste(){
doc.execCommand('Paste');
Editor.focus();
}
function doUndo(){
doc.execCommand('Undo');
Editor.focus();
}
function doRedo(){
doc.execCommand('Redo');
Editor.focus();
}
function doDelete(){
doc.execCommand('Delete');
Editor.focus();
}
function doFontName(fn){
doc.execCommand('FontName', false, fn);
Editor.focus();
}
function doFontSize(fs){
doc.execCommand('FontSize', false, fs);
Editor.focus();
}
function doBold(){
doc.execCommand('Bold');
Editor.focus();
}
function doItalic(){
doc.execCommand('Italic');
Editor.focus();
}
function doUnderline(){
doc.execCommand('Underline');
Editor.focus();
}
function doStrikeThrough(){
doc.execCommand('StrikeThrough');
Editor.focus();
}
function doSubscript(){
doc.execCommand('Subscript');
Editor.focus();
}
function doSuperscript(){
doc.execCommand('Superscript');
Editor.focus();
}
function doJustifyLeft(){
doc.execCommand('JustifyLeft');
Editor.focus();
}
function doJustifyRight(){
doc.execCommand('JustifyRight');
Editor.focus();
}
function doJustifyCenter(){
doc.execCommand('JustifyCenter');
Editor.focus();
}
function doIndent(){
doc.execCommand('Indent');
Editor.focus();
}
function doOutdent(){
doc.execCommand('Outdent');
Editor.focus();
}
function doForeColor(){
var fcolor=showModalDialog("editor_color.htm",false,"dialogWidth:106px;dialogHeight:126px;status:0;");
doc.execCommand('ForeColor',false,fcolor);
Editor.focus();
}
function doBackColor(){
var bcolor=showModalDialog("editor_color.htm",false,"dialogWidth:106px;dialogHeight:126px;status:0;");
doc.execCommand('BackColor',false,bcolor);
Editor.focus();
}
function doInsertTable(){
var dotable=showModalDialog("editor_table.htm",false,"dialogWidth:200px;dialogHeight:156px;status:0;");
if (dotable!=undefined){
doc.body.innerHTML=doc.body.innerHTML+dotable;
}else{
return false;
}
Editor.focus();
}
function doInsertOrderedList(){
doc.execCommand('InsertOrderedList');
Editor.focus();
}
function doInsertUnorderedList(){
doc.execCommand('InsertUnorderedList');
Editor.focus();
}
function xdoCreateLink(){
	var txRange = doc.selection.createRange();
		alert(txRange.htmlText);
	var fcolor=showModalDialog("edit_link.htm",false,"scrollbars=yes,width=460,height=190");
	if (fcolor!=undefined){
		var newHTML = '<A href="' + fcolor + '" title="haha">' + txRange.htmlText + '</A>';
		txRange.pasteHTML(newHTML);
		txRange.select();
		Editor.focus();
		
	}
}
function doCreateLink(){
	var txRange = doc.selection.createRange();
//	alert(txRange.htmlText);
	var el = txRange.parentElement();
	var myObject = new Object();
//		alert(el.href);
    	myObject.href = el.href;
    	myObject.title = el.title;
	var fcolor=showModalDialog("edit_link.htm",myObject,"dialogWidth:550px;dialogHeight:250px;status:0;");
	if (fcolor!=undefined){
		var oXML = new ActiveXObject("Microsoft.XMLDOM");
		oXML.async = false;
		oXML.loadXML(fcolor);
//		alert(fcolor);
		var tx, re;
		tx = oXML.selectSingleNode("//URL").text;
		re = /&amp;/g
		tx = tx.replace(re,"&");
		doc.execCommand('CreateLink',false,tx);
		Editor.focus();
		txRange = doc.selection.createRange();
//		alert(txRange.htmlText);
		el = txRange.parentElement();
		el.title = oXML.selectSingleNode("//TITLE").text;
		if ("Y" == oXML.selectSingleNode("//newWindow").text) 
			{ 
			el.target = "_gipNW";
			<%if checkGIPconfig("gipengmoe") then %>
			el.title = el.title + "(open new window)";
			<%else%>
			el.title = el.title + "(另開視窗)";
			<%end if%>
			}
	}
}


function doAttachLink(){
//Editor.focus();

var fcolor=showModalDialog("editor_attach.asp?icuitem=<%=request("icuitem")%>",false,"dialogWidth:500px;dialogHeight:200px;status:0;");
	if (fcolor!=undefined){
	//doc.execCommand('CreateLink',false,'public/Attachment/'+fcolor);
	
	var fileArray = fcolor.split(".");
	//alert(fileArray.length);
	var sFileName = "";
	var sExtFileName = "";
	if(fileArray.length > 0)
		sFileName = fileArray[0];
	if(fileArray.length > 1)
		sExtfileName = fileArray[1];
	
	var txRange = doc.selection.createRange();
	
	var tmpArray = txRange.text.split("(" + sExtfileName + "檔案)");
	
	var sTitle="";
	sTitle = txRange.text;
	if(tmpArray.length > 0)
		sTitle = tmpArray[0];		
	
	sTitle =  sTitle.replace("<DIV>","").replace("</DIV>","").replace("\r","").replace("\n","");
	//alert(txRange.htmlText + "    " + sTitle);
	<%if checkGIPconfig("gipengmoe") then %>
	txRange.pasteHTML("<A href=\"" + "public/Attachment/" + fcolor +  "\"  title=\""+sTitle+"("+sExtfileName + "file download;open new window)\" target='_new' >" + sTitle + "(" + sExtfileName + " file)" + "</A>");
	<%else%>
	txRange.pasteHTML("<A href=\"" + "public/Attachment/" + fcolor +  "\"  title=\""+sTitle+"("+sExtfileName + "檔案下載;另開新視窗)\" target='_new' >" + sTitle + "(" + sExtfileName + "檔案)" + "</A>");
	<%end if%>
	//txRange.pasteHTML("<A href=\"" + "public/Attachment/" + fcolor +  "\"  title=\""+sTitle+"("+sExtfileName + "檔案下載;另開新視窗)\" target='_new' >" + sTitle + "</A>");
	txRange.select();
	Editor.focus();
	}
}


function doPageLink(){
//Editor.focus();
var fcolor=showModalDialog("editor_pgLink.asp?icuitem=<%=request("icuitem")%>",false,"dialogWidth:500px;dialogHeight:200px;status:0;");
	if (fcolor!=undefined){
		//if checkGIPconfig("jspFront"){
		//	doc.execCommand('CreateLink',false,'?a=content&CuItem='+fcolor);
		//}else{
			doc.execCommand('CreateLink',false,'content.asp?CuItem='+fcolor);
		//}		
	Editor.focus();
	}
}
function doInsertImage(MMOHTMLTag){
	Editor.focus();
	var txRange = doc.selection.createRange();
//	alert(txRange.htmlText);
	txRange.pasteHTML(txRange.htmlText + MMOHTMLTag);
	txRange.select();
	Editor.focus();
//	doc.execCommand('InsertImage',false,MMOPath);
//	var txRange = doc.selection.createRange();
//	alert(txRange.htmlText);
//Editor.focus();
}
function doInsertImageLeft(MMOPath, xTitle){
	Editor.focus();
	var txRange = doc.selection.createRange();
//	alert(txRange.htmlText);
	txRange.pasteHTML(txRange.htmlText + '<a href="' + MMOPath + '"><IMG SRC="' + MMOPath + '" alt="' + xTitle + '" style="float:left" /></a>');
	txRange.select();
	Editor.focus();
// doc.execCommand('InsertImage',false,MMOPath);
//	var txRange = doc.selection.createRange();
//	alert(txRange.htmlText);
Editor.focus();
}
function doInsertImageRight(MMOPath, xTitle){
	Editor.focus();
	var txRange = doc.selection.createRange();
//	alert(txRange.htmlText);
	txRange.pasteHTML(txRange.htmlText + '<a href="' + MMOPath + '"><IMG SRC="' + MMOPath + '" alt="' + xTitle + '" style="float:right" /></a>');
	txRange.select();
	Editor.focus();
// doc.execCommand('InsertImage',false,MMOPath);
//	var txRange = doc.selection.createRange();
//	alert(txRange.htmlText);
Editor.focus();
}

function doImgLeft(){
	//doSpan('float:left', function (oStyle) { return oStyle.styleFloat.toLowerCase() == 'left' });
	doSpan(function (oStyle) { oStyle.styleFloat = 'left' });
	Editor.focus();
}
function doImgRight(){
	//doSpan('float:right', function (oStyle) { return oStyle.styleFloat.toLowerCase() == 'right' });
	doSpan(function (oStyle) { oStyle.styleFloat = 'right' });
	Editor.focus();
}
// New functions
//function doSpan(cssText, returnCaseFunc) {
function doSpan(setStyleFunc) {
	Editor.focus();
	var selType = doc.selection.type.toLowerCase();
	if (selType == 'none')
		return
	var range = doc.selection.createRange();
	var span;
	do {
		if (selType == 'control') {
			var rangeFirst = range(0);
			//if span with appropriate style is selected, return
			if (range.length == 1 &&
				rangeFirst.nodeName.toLowerCase() == 'span' /*&&
				returnCaseFunc(range(0).style)*/) {
				alert('1');	//case not verified
				span = rangeFirst;
				break;
			}
			//if parent element of selected elements is span with appropriate style, return
			var rangeParent = rangeFirst.parentNode;
			if (rangeParent.nodeName &&
				rangeParent.nodeName.toLowerCase() == 'span' /*&&
				returnCaseFunc(rangeParent.style)*/) {
				alert('2');	//case not verified
				span = rangeParent;
				break;
			}
			//range(0).insertAdjacentHTML('beforeBegin', '<span style="' + cssText + '"></span>');
			rangeFirst.insertAdjacentHTML('beforeBegin', '<span></span>');
			span = rangeFirst.previousSibling;
			for (var i = 0; i < range.length; i++) {
				span.appendChild(range(i));
			}
			//following doesn't work since insertAdjacentHTML automatically adds ending tags
			//range(0).insertAdjacentHTML('beforeBegin', '<span style="float:left">');
			//range(range.length - 1).insertAdjacentHTML('afterEnd', '</span>');
		} else {	//if selType == 'text'
			//if element containing selected text is span with appropriate style, return
			var rangeParent = range.parentElement();
			if (rangeParent.nodeName.toLowerCase() == 'span' /*&&
				returnCaseFunc(rangeParent.style)*/) {
				alert('3');	//case verified
				span = rangeParent;
				break;
			}
			//if parent element of element containing selected text is span with appropriate style, return
			var rangeParentParent = rangeParent.parentNode;
			if (rangeParentParent.nodeName.toLowerCase() == 'span' /*&&
				returnCaseFunc(rangeParentParent.style)*/) {
				alert('4');	//case verified
				span = rangeParentParent;
				break;
			}
			//if span with appropriate style is selected, return
			if (rangeParent.children.length == 1 &&
				rangeParent.children(0).nodeName.toLowerCase() == 'span' /*&&
				returnCaseFunc(rangeParent.children(0).style)*/) {
				alert('5');	//case not verified
				span = rangeParent.children(0);
				break;
			}
			//var newHTML = '<span style="' + cssText + '">' + range.htmlText + '</span>';
			var newHTML = '<span>' + range.htmlText + '</span>';
			range.pasteHTML(newHTML);
			//apparently pasteHTML collapses range to end of pasted HTML,
			//	so go back 1 character (not including markup text) to enter span
			range.move('character', -1);
			var span = range.parentElement();
		}
	} while (false);
	setStyleFunc(span.style);
	range.select();
}

function xdoRemoveFormat(){
	doc.execCommand('RemoveFormat');
	var txRange = doc.selection.createRange();
	alert(txRange.htmlText);
	Editor.focus();
}

function doShowDetails() {
//	doc.execCommand('RemoveFormat');
	Editor.focus();
	var txRange = doc.selection.createRange();
		var el = txRange.parentElement();
		alert(el.tagName);
//		el.valign="top";

}


// end New funtions
function doIndent(){
doc.execCommand('Indent');
Editor.focus();
}
function doInsertBR(){
	Editor.focus();
	var txRange = doc.selection.createRange();
//	alert(txRange.htmlText);
	txRange.pasteHTML(txRange.htmlText + "<BR/>");
	txRange.select();
	Editor.focus();
}
function doInsertInputButton(){
Editor.focus();
doc.execCommand('InsertInputButton');
}
function doInsertHorizontalRule(){
Editor.focus();
doc.execCommand('InsertHorizontalRule');
}
function doInsertInputCheckbox(){
Editor.focus();
doc.execCommand('InsertInputCheckbox');
}
function doInsertInputRadio(){
Editor.focus();
doc.execCommand('InsertInputRadio');
}
function doInsertInputText(){
Editor.focus();
doc.execCommand('InsertInputText');
}
function doInsertInputPassword(){
Editor.focus();
doc.execCommand('InsertInputPassword');
}
function doInsertInputSubmit(){
Editor.focus();
doc.execCommand('InsertInputSubmit');
ShowMessage();
}
function doInsertInputReset(){
Editor.focus();
doc.execCommand('InsertInputReset');
ShowMessage();
}
function doInsertMarquee(){
Editor.focus();
doc.execCommand('InsertMarquee');
ShowMessage();
}
function doInsertSelectDropdown(){
Editor.focus();
doc.execCommand('InsertSelectDropdown');
}
function doInsertTextArea(){
Editor.focus();
doc.execCommand('InsertTextArea');
}
function doPrint(){
doc.execCommand('Print');
Editor.focus();
}
function doSaveAs(){
doc.execCommand('SaveAs',0,"未命名");
Editor.focus();
}
function doOpen(){
doc.execCommand('Open');
Editor.focus();
}


function EditResource(){
Preview.value=doc.body.innerHTML;
return false;
}
function ClearAll(){
doc.body.innerHTML='';
Preview.value='';
return false;
}
function SeePreview(){
doc.body.innerHTML=Preview.value;
return false;
}
function AutoPreview(){
//if(vx.checked){
//SeePreview();
//}
}
function EditMode(){
doc.designMode = "On";
window.setTimeout('SeePreview()',100);
Preview.focus();
}
function PreviewMode(){
doc.designMode = "Off";
window.setTimeout('SeePreview()',100);
Preview.focus();
}
function ShowMessage(){
alert("請按兩下物件編輯內容");
}

// add by ken 2006-3-9
function doInsertTable() {
	Editor.focus();
	var range = doc.selection.createRange();
	editTableDialog(range);
	Editor.focus();
}
function editTableDialog(arg) {
	return window.showModalDialog("CuHtmlEditTable.html", arg, "dialogWidth:600px;dialogHeight:750px;status:0;resizable:1");
}
function doEditTable() {
	Editor.focus();
	var range = doc.selection.createRange();
	var table;
	if (range.length) {
		//find the first element with an ancestor table and use that table
		for (var i = 0; i < range.length; i++) {
			table = range(i);
			var tag = table.tagName.toLowerCase();
			while (tag != 'table') {
				if (tag == 'body')
					return;
				table = table.parentNode;
				tag = table.tagName.toLowerCase();
			}
		}
	} else {
		table = range.parentElement();
		var tag = table.tagName.toLowerCase();
		while (tag != 'table') {
			if (tag == 'body')
				return;
			table = table.parentNode;
			tag = table.tagName.toLowerCase();
		}
	}
	editTableDialog(table);
	Editor.focus();
}
</SCRIPT>

<META content="MSHTML 5.00.3315.2870" name=GENERATOR></HEAD>
<BODY onkeyup=AutoPreview(); onLoad="SeePreview();Preview.focus();">
<div id="FuncName">
	<h1><%=HTProgCap%></h1>
	<div id="Nav"><a href="DsdXMLEdit.asp?phase=edit&icuitem=<%=RSreg("icuitem")%>">回上頁</a></div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">
	<%=HTProgCap%>&nbsp;
    <font size=2>【編輯--主題單元:<%=session("CtUnitName")%> / 單元資料:<%=nullText(htPageDom.selectSingleNode("//tableDesc"))%>】
</div>

<table cellspacing="0" id="LayoutTable">
  <tr>
    <td class="Func">
		<table cellspacing="0" id="FuncIcon">
	      <tr>
	        <td class="Holder"><img src="editor/images/main_holder.gif"></td>
	        <td class="Icon"><a href="javascript:;" title="存檔"><img src="editor/images/icon_edit_save.gif"></a></td>
	        <td class="Icon"><a href="javascript:;" title="列印"><img src="editor/images/icon_edit_print.gif"></a></td>
	      </tr>
		</table>
		<table cellspacing="0" id="FuncIcon">
	        <tr>
	          <td class="Icon"><a href="javascript:;" title="剪下"><img src="editor/images/icon_edit_cut.gif"
	          onclick="doCut();"></a></td>
	          <td class="Icon"><a href="javascript:;" title="複製"><img src="editor/images/icon_edit_copy.gif"
	          onclick="doCopy();"></a></td>
	          <td class="Icon"><a href="javascript:;" title="貼上"><img src="editor/images/icon_edit_past.gif"
	          onclick="doPaste();"></a></td>
	        </tr>
		</table>
		<table cellspacing="0" id="FuncIcon">
	        <tr>
	          <td class="Icon"><a href="javascript:;" title="刪除"><img src="editor/images/icon_edit_delete.gif"
	          onclick="doDelete();"></a></td>
	          <td class="Icon"><a href="javascript:;" title="復原"><img src="editor/images/icon_edit_undo.gif"
	          onclick="doUndo();"></a></td>
	          <td class="Icon"><a href="javascript:;" title="取消復原"><img src="editor/images/icon_edit_redo.gif"
	          onclick="doRedo();"></a></td>
			</tr>
		</table>
		<table cellspacing="0" id="FuncIcon">
	      <tr>
	        <td class="Holder"><img src="editor/images/main_holder.gif"></td>
	        <td class="Lead"><img src="editor/images/icon_html.gif" alt="插入HTML標籤"></td>
	        <td class="Icon"><a href="javascript:;" title="分隔線 hr"><img src="editor/images/icon_html_hr.gif"
	        onclick="doInsertHorizontalRule();"></a></td>
	        <td class="Icon"><a href="javascript:;" title="斷行 br"><img src="editor/images/icon_html_br.gif"
	        onclick="doInsertBR();"></a></td>
	        <td class="Icon"><a href="javascript:;" title="移除格式"><img src="HTMLeditor/editor_e.gif"
	        onclick="VBS: doRemoveFormat()"></a></td>
	        <td class="Icon"><a href="javascript:;" title="顯示標籤"><img src="editor/images/details.gif"
	        onclick="doShowDetails();"></a></td>
		  </tr>
	    </table>
	</td>
  </tr>
  <tr>
    <td class="Func">
    	<table cellspacing="0" id="FuncIcon">
	      <tr>
	        <td class="Holder"><img src="editor/images/main_holder.gif"></td>
	        <td class="Lead"><img src="editor/images/icon_text.gif" alt="文字格式設定"></td>
	        <td class="N">
					<select name="select" class="InputSelect">
					  <option>內文</option>
					  <option>段落文字1</option>
					  <option>段落文字2</option>
					  <option>段落文字3</option>
					  <option>標題1</option>
					  <option>標題2</option>
					  <option>標題3</option>
	          </select>
	        </td>
	        <td class="Icon"><a href="javascript:;" title="粗體字"><img src="editor/images/icon_text_strong.gif"
	        onclick="doBold();"></a></td>
	        <td class="Icon"><a href="javascript:;" title="斜體字"><img src="editor/images/icon_text_italic.gif"
	        onclick="doItalic();"></a></td>
	        <td class="Icon"><a href="javascript:;" title="底線"><img src="editor/images/icon_text_underline.gif"
	        onclick="doUnderline();"></a></td>
	        <td class="Icon"><a href="javascript:;" title="刪除線"><img src="editor/images/icon_text_through.gif"
	        onclick="doStrikeThrough();"></a></td>
	        </tr>
		</table>
		<table cellspacing="0" id="FuncIcon">
	        <tr>
	        <td class="Icon"><a href="javascript:;" title="編號清單"><img src="editor/images/icon_text_ol.gif" width="17" height="16"
	        onclick="doInsertOrderedList();"></a></td>
	        <td class="Icon"><a href="javascript:;" title="清單"><img src="editor/images/icon_text_ul.gif" width="17" height="16"
	        onclick="doInsertUnorderedList();"></a></td>
	        </tr>
		</table>
		<table cellspacing="0" id="FuncIcon">
	        <tr>
	        <td class="Icon"><a href="javascript:;" title="靠左對齊"><img src="editor/images/editor_f16.gif"
	        onclick="doJustifyLeft();"></a></td>
	        <td class="Icon"><a href="javascript:;" title="靠中對齊"><img src="editor/images/editor_f17.gif"
	        onclick="doJustifyCenter();"></a></td>
	        <td class="Icon"><a href="javascript:;" title="靠右對齊"><img src="editor/images/editor_f18.gif"
	        onclick="doJustifyRight();"></a></td>
	        </tr>
		</table>
		<table cellspacing="0" id="FuncIcon">
	        <tr>
	        <td class="Icon"><a href="javascript:;" title="增加縮排"><img src="editor/images/editor_f19.gif"
	        onclick="doIndent();"></a></td>
	        <td class="Icon"><a href="javascript:;" title="減少縮排"><img src="editor/images/editor_f20.gif"
	        onclick="doOutdent();"></a></td>
	        </tr>
		</table>
      </td>
  </tr>
  <tr>
    <td class="Func">
		<!--table cellspacing="0" id="FuncIcon">
	      <tr>
	        <td class="Holder"><img src="editor/images/main_holder.gif"></td>
	        <td class="Lead"><img src="editor/images/icon_table.gif" alt="插入表格"></td>
	        <td class="Icon"><a href="javascript:;" title="資料標題於上列"><img src="editor/images/icon_table_HTop.gif" 
	        onClick="MM_openBrWindow('edit_table.asp','','scrollbars=yes,width=560,height=400')"></a></td>
	        <td class="Icon"><a href="javascript:;" title="資料標題於左欄"><img src="editor/images/icon_table_HLeft.gif"></a></td>
	        <td class="Icon"><a href="javascript:;" title="資料標題於上及左"><img src="editor/images/icon_table_HBoth.gif"></a></td>
	        <td class="Icon"><a href="javascript:;" title="無資料標題"><img src="editor/images/icon_table_HNone.gif"></a></td>
	      </tr>
		</table-->
		<!-- Add by ken 2006-3-9 -->
		<table cellspacing="0" id="FuncIcon">
	      <tr>
	        <td class="Holder"><img src="editor/images/main_holder.gif"></td>
	        <!--td class="Lead"><img src="editor/images/icon_table.gif" alt="插入表格"></td>
	        <td class="Icon"><a href="javascript:;" title="資料標題於上列"><img src="editor/images/icon_table_HTop.gif"
	        onclick="MM_openBrWindow('edit_table.asp','','scrollbars=yes,width=560,height=400')"></a></td>
	        <td class="Icon"><a href="javascript:;" title="資料標題於左欄"><img src="editor/images/icon_table_HLeft.gif"></a></td>
	        <td class="Icon"><a href="javascript:;" title="資料標題於上及左"><img src="editor/images/icon_table_HBoth.gif"></a></td>
	        <td class="Icon"><a href="javascript:;" title="無資料標題"><img src="editor/images/icon_table_HNone.gif"></a></td-->
	        <td class="Icon"><a href="javascript:;" title="插入表格"><img src="editor/images/icon_table.gif"
	        	onclick="doInsertTable();"></a></td>
	        <td class="Icon"><a href="javascript:;" title="修改表格"><img src="editor/images/icon_table.gif"
	        	onclick="doEditTable();"></a></td>
	      </tr>
		</table>
		<!-- Add by ken 2006-3-9 -->
		<table cellspacing="0" id="FuncIcon">
	      <tr>
	        <td class="Holder"><img src="editor/images/main_holder.gif"></td>
	        <td class="Lead"><img src="editor/images/icon_media.gif" alt="插入多媒體物件"></td>
	        <td class="Icon"><a href="javascript:;" title="插入多媒體物件"><img src="editor/images/icon_media_image.gif" onClick="MM_openBrWindow('CuMMOQuery.asp','','scrollbars=yes,width=700,height=500')"></a></td>
	        <td class="Icon"><a href="javascript:;" title="插入圖檔靠左"><img src="editor/images/icon_media_image_left.gif" onClick="MM_openBrWindow('CuMMOQuery.asp?imgpos=left','','scrollbars=yes,width=700,height=500')"></a></td>
	        <td class="Icon"><a href="javascript:;" title="插入圖檔靠右"><img src="editor/images/icon_media_image_right.gif" onClick="MM_openBrWindow('CuMMOQuery.asp?imgpos=right','','scrollbars=yes,width=700,height=500')"></a></td>
	        <!--td class="Icon"><a href="javascript:;" title="圖檔靠左"><img src="editor/images/icon_media_image_left.gif"
	        onclick="doImgLeft();"></a></td>
	       	<td class="Icon"><a href="javascript:;" title="圖檔靠右"><img src="editor/images/icon_media_image_right.gif"
	        onclick="doImgRight();"></a></td-->

	        <!--td class="Icon"><a href="javascript:;" title="插入聲音檔"><img src="editor/images/icon_media_audio.gif"></a></td>
	        <td class="Icon"><a href="javascript:;" title="插入Flash檔"><img src="editor/images/icon_media_flash.gif"></a></td>
	        <td class="Icon"><a href="javascript:;" title="插入影片檔"><img src="editor/images/icon_media_video.gif"></a></td-->
	      </tr>
	    </table>
	    <table cellspacing="0" id="FuncIcon">
	        <tr>
	          <td class="Holder"><img src="editor/images/main_holder.gif"></td>
	          <td class="Lead"><img src="editor/images/icon_link.gif" alt="連結設定"></td>
	          <td class="Icon"><a href="javascript:;" title="連結網址"><img src="editor/images/icon_link_http.gif" 
	          onClick="doCreateLink();"></a></td>
	          <td class="Icon"><a href="javascript:;" title="連結附件"><img src="editor/images/icon_link_att.gif" width="17" height="15"
	          onclick="doAttachLink();"></a></td>
	          <td class="Icon"><a href="javascript:;" title="連結本單元網頁"><img src="editor/images/icon_link_page.gif" width="18" height="14"
	          onclick="doPageLink();"></a></td>
	          <!--td class="Icon"><a href="javascript:;" title="連結其他單元"><img src="editor/images/icon_link_unit.gif" width="19" height="14"></a></td>
	          <td class="Icon"><a href="javascript:;" title="連結其他單元網頁"><img src="editor/images/icon_link_unitpage.gif" width="18" height="14"></a></td-->
		    </tr>
	    </table>
    </td>
  </tr>
  <tr>
    <td class="PageView">頁面呈現<br>
      <IFRAME id=Editor marginWidth=1 src="about:blank"></IFRAME>
<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=htx_icuitem value="<%=RSreg("icuitem")%>">
<INPUT TYPE=hidden name=htx_sTitle value="<%=RSreg("sTitle")%>">
<INPUT TYPE=hidden name=submitTask value="">
<INPUT TYPE=hidden name=htx_xbody  value="">
<input type=button value ="編修存檔" class="InputButton" onClick="formModSubmit()">
   </FORM>
		</td>
  </tr>
  <tr>
    <td class="CodeView">原始碼　<BUTTON class="InputButton" 
            onclick=EditResource();Preview.focus();>檢視原始碼</BUTTON>&nbsp;&nbsp;
      <BUTTON class="InputButton"  
            onclick=SeePreview();Preview.focus();>結果預覽</BUTTON>
      
      <br>    
	<TEXTAREA id=Preview xreadonly="readonly" wrap="virtual"><%=htmlMessage(RSreg("xbody"))%></TEXTAREA>
		</td>
  </tr>
</table>

<!--TABLE border=1 borderColorDark=#efece8 borderColorLight=#888070 cellPadding=0 
cellSpacing=0 width="100%">
  <TBODY>
  <TR>
    <TD bgColor=#d9cec4>
      <TABLE border=0 cellPadding=0 cellSpacing=1>
        <TBODY>
        <TR>
          <TD><IMG src="HTMLeditor/editor_h.gif"></TD>
          <TD>
            <DIV onclick=doSaveAs(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=存檔><IMG 
            src="HTMLeditor/editor_f01.gif"></DIV></TD>
          <TD>
            <DIV onclick=doPrint(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=列印><IMG 
            src="HTMLeditor/editor_f02.gif"></DIV></TD>
          <TD><IMG src="HTMLeditor/editor_s.gif"></TD>
          <TD>
            <DIV onclick=doCut(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=剪下><IMG 
            src="HTMLeditor/editor_f03.gif"></DIV></TD>
          <TD>
            <DIV onclick=doCopy(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=複製><IMG 
            src="HTMLeditor/editor_f04.gif"></DIV></TD>
          <TD>
            <DIV onclick=doPaste(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=貼上><IMG 
            src="HTMLeditor/editor_f05.gif"></DIV></TD>
          <TD>
            <DIV onclick=doDelete(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=刪除><IMG 
            src="HTMLeditor/editor_f06.gif"></DIV></TD>
          <TD>
            <DIV onclick=doUndo(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=復原><IMG 
            src="HTMLeditor/editor_f07.gif"></DIV></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD bgColor=#d9cec4>
      <TABLE border=0 cellPadding=0 cellSpacing=1>
        <TBODY>
        <TR>
          <TD><IMG src="HTMLeditor/editor_h.gif"></TD>
          <TD><SELECT 
            onchange=doFontName(this[this.selectedIndex].value);this.selectedIndex=0;> 
              <OPTION selected value="">字型<OPTION value=細明體>細明體<OPTION 
              value=新細明體>新細明體<OPTION value=標楷體>標楷體<OPTION 
              value=arial>arial<OPTION 
          value=wingdings>wingdings</OPTION></SELECT></TD>
          <TD><SELECT 
            onchange=doFontSize(this[this.selectedIndex].value);this.selectedIndex=0;> 
              <OPTION selected value="">大小<OPTION value=1>1<OPTION 
              value=2>2<OPTION value=3>3(預設)<OPTION value=4>4<OPTION 
              value=5>5<OPTION value=6>6<OPTION value=7>7</OPTION></SELECT></TD>
          <TD><IMG src="HTMLeditor/editor_s.gif"></TD>
          <TD>
            <DIV onclick=doBold(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=粗體字><IMG 
            src="HTMLeditor/editor_f08.gif"></DIV></TD>
          <TD>
            <DIV onclick=doItalic(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=斜體字><IMG 
            src="HTMLeditor/editor_f09.gif"></DIV></TD>
          <TD>
            <DIV onclick=doUnderline(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=劃底線><IMG 
            src="HTMLeditor/editor_f10.gif"></DIV></TD>
          <TD>
            <DIV onclick=doStrikeThrough(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=刪除線><IMG 
            src="HTMLeditor/editor_f11.gif"></DIV></TD>
          <TD><IMG src="HTMLeditor/editor_s.gif"></TD>
          <TD>
            <DIV onclick=doSuperscript(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=上標字><IMG 
            src="HTMLeditor/editor_f12.gif"></DIV></TD>
          <TD>
            <DIV onclick=doSubscript(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=下標字><IMG 
            src="HTMLeditor/editor_f13.gif"></DIV></TD>
          <TD><IMG src="HTMLeditor/editor_s.gif"></TD>
          <TD>
            <DIV onclick=doForeColor(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=文字顏色><IMG 
            src="HTMLeditor/editor_f14.gif"></DIV></TD>
          <TD>
            <DIV onclick=doBackColor(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=背景顏色><IMG 
            src="HTMLeditor/editor_f15.gif"></DIV></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD bgColor=#d9cec4>
      <TABLE cellPadding=0 cellSpacing=1>
        <TBODY>
        <TR>
          <TD><IMG src="HTMLeditor/editor_h.gif"></TD>
          <TD>
            <DIV onclick=doJustifyLeft(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=靠左對齊><IMG 
            src="HTMLeditor/editor_f16.gif"></DIV></TD>
          <TD>
            <DIV onclick=doJustifyCenter(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=靠中對齊><IMG 
            src="HTMLeditor/editor_f17.gif"></DIV></TD>
          <TD>
            <DIV onclick=doJustifyRight(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=靠右對齊><IMG 
            src="HTMLeditor/editor_f18.gif"></DIV></TD>
          <TD><IMG src="HTMLeditor/editor_s.gif"></TD>
          <TD>
            <DIV onclick=doIndent(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=增加縮排><IMG 
            src="HTMLeditor/editor_f19.gif"></DIV></TD>
          <TD>
            <DIV onclick=doOutdent(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=減少縮排><IMG 
            src="HTMLeditor/editor_f20.gif"></DIV></TD>
          <TD><IMG src="HTMLeditor/editor_s.gif"></TD>
          <TD>
            <DIV onclick=doInsertOrderedList(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=數字標題><IMG 
            src="HTMLeditor/editor_f21.gif"></DIV></TD>
          <TD>
            <DIV onclick=doInsertUnorderedList(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=無數字標題><IMG 
            src="HTMLeditor/editor_f22.gif"></DIV></TD>
          <TD><IMG src="HTMLeditor/editor_s.gif"></TD>
          <TD>
            <DIV onclick=doInsertHorizontalRule(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=插入分隔線><IMG 
            src="HTMLeditor/editor_f23.gif"></DIV></TD>
          <TD>
            <DIV onclick=doInsertTable(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=插入表格><IMG 
            src="HTMLeditor/editor_f24.gif"></DIV></TD>
          <TD>
            <DIV onclick=doCreateLink(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=插入連結><IMG 
            src="HTMLeditor/editor_f25.gif"></DIV></TD>
          <TD>
            <DIV onclick=doAttachLink(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=連結附件><IMG 
            src="HTMLeditor/editor_f35.gif"></DIV></TD>
          <TD>
            <DIV onclick=doPageLink(); onmousedown=BtnClick(this); 
            onmouseout=BtnOut(this); onmouseover=BtnOver(this); 
            onmouseup=BtnOver(this); title=連結網頁><IMG 
            src="HTMLeditor/editor_f36.gif"></DIV></TD>

          <TD>&nbsp;</TD>
          <TD>&nbsp;</TD>
          <TD>&nbsp;</TD>
          <TD><IMG src="HTMLeditor/editor_h.gif"></TD>
          <TD><INPUT CHECKED name=vm onclick=EditMode(); title=編輯模式 
          type=radio></TD>
          <TD>
            <DIV title=編輯模式><IMG src="HTMLeditor/editor_e.gif"></DIV></TD>
          <TD>&nbsp;</TD>
          <TD><INPUT name=vm onclick=PreviewMode(); title=預覽模式 type=radio></TD>
          <TD>
            <DIV title=預覽模式><IMG src="HTMLeditor/editor_p.gif"></DIV></TD>
            
            </TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD bgColor=#d9cec4>
      <TABLE cellPadding=0 cellSpacing=1>
        <TBODY>
        <TR>
          <TD><IMG src="HTMLeditor/editor_h.gif"></TD>
          <TD><BUTTON class=function 
            onclick=EditResource();Preview.focus();>編輯原始碼</BUTTON></TD>
          <TD><BUTTON class=function 
            onclick=ClearAll();Preview.focus();>全部清除</BUTTON></TD>
          <TD><BUTTON class=function 
            onclick=SeePreview();Preview.focus();>結果預覽</BUTTON></TD>
          <TD><IMG src="HTMLeditor/editor_s.gif"></TD>
          <TD><INPUT CHECKED name=vx onclick=AutoPreview();Preview.focus(); 
            title=自動預覽 type=checkbox></TD>
          <TD>
            <DIV title=自動預覽><IMG 
        src="HTMLeditor/editor_a.gif"></DIV></TD></TR></TBODY></TABLE></TD></TR>
   <TR><TH></TD></TR> 
    </TBODY></TABLE--></CENTER>
<SCRIPT>
var doc;
doc=document.frames.Editor.document;
// alert(doc.body.innerHTML);
doc.designMode = "On";
window.setTimeout('Editor.focus()',300);
window.setTimeout('SeePreview()',200);
//SeePreview();
//Editor.focus();
</SCRIPT>
<script language=vbs>
sub xb1_onClick
	msgbox "hah"
	msgBox doc.body.innerHTML
end sub
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
 
 	reg.htx_xbody.value = doc.body.innerHTML
' 	msgBox  reg.htx_xbody.value
 
'	doc.body.innerHTML = purifyHTML(doc.body.innerHTML)
	reg.submitTask.value = "UPDATE"
  	reg.Submit
end sub

sub purifyMe(xme)
	for each xi in xme.children
		purifyMe xi
	next
	if xme.innerHTML = "" then
		xme.outerHTML = ""
	end if
'	msgbox xme.style
	
end sub

sub doRemoveFormat()
	doc.body.innerHTML = purifyHTML(doc.body.innerHTML)
	for each x in doc.body.all.tags("IMG")
		if x.alt="" then	x.alt=reg.htx_sTitle.value + " "
	next
	xstr = ""
'	for each xi in doc.body.children
'		purifyMe xi
'		xstr = xstr & xi.outerHTML & vbCRLF
'	next
'	msgBox xstr
end sub

sub xdoRemoveFormat()
	doc.execCommand("RemoveFormat")
	set txRange = doc.selection.createRange()
'	alert(txRange.htmlText)
	txRange.pasteHTML(purifyHTML(txRange.htmlText))
	txRange.select()
	Editor.focus()
end sub

function purifyHTML(xHTMLstr)
	xstr = xHTMLstr
'	xstr = "The quick brown fox jumped over the lazy dog."
'	msgBox xstr
'	xstr = ReplaceTest("\s*mso-[^:]+:.[^;""]+[;]", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
'	msgBox xstr
'	xstr = ReplaceTest("\s*mso-[^:]+:.[^""]+[""]", """", xstr)	'-- mso-bidi-font-weight: normal"
'	msgBox xstr
	xstr = ReplaceTest("\sstyle=""[^""]*""", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("<SPAN[^>]*>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("</SPAN>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("\swidth=[0-9]*\s", " ", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("\swidth=[0-9]*>", ">", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("\sbgcolor=[^>\s]*\s", " ", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("\sbgcolor=[^>\s]*>", ">", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("\sheight=[^>\s]*\s", " ", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("\sheight=[^>\s]*>", ">", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("<FONT[^>]*>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("</FONT>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("<\?xml:namespace[^>]*>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("<o:p></o:p>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
'	xstr = ReplaceTest("<table[^>]*>", "<table summary=""列表資料 "">", xstr)	'-- mso-bidi-font-size: 10.0pt;
'	xstr = ReplaceTest("<table([^>]*)>", "<table $1 summary=""列表資料 "">", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = aaaTable(xstr)
'	msgBox xstr
	purifyHTML = xstr
	
end function

Function RegExpTest(patrn, strng)
   Dim regEx, Match, Matches   ' Create variable.
   Set regEx = New RegExp   ' Create a regular expression.
   regEx.Pattern = patrn   ' Set pattern.
   regEx.IgnoreCase = True   ' Set case insensitivity.
   regEx.Global = True   ' Set global applicability.
   Set Matches = regEx.Execute(strng)   ' Execute search.
   For Each Match in Matches   ' Iterate Matches collection.
      RetStr = RetStr & "Match found at position "
      RetStr = RetStr & Match.FirstIndex & ". Match Value is '"
      RetStr = RetStr & Match.Value & "'." & vbCRLF
   Next
   RegExpTest = RetStr
End Function

Function ReplaceTest(patrn, replStr, str1)
  Dim regEx               ' Create variables.
'  str1 = "The quick brown fox jumped over the lazy dog."
  Set regEx = New RegExp            ' Create regular expression.
  regEx.Pattern = patrn            ' Set pattern.
  regEx.IgnoreCase = True            ' Make case insensitive.
  regEx.Global = True   ' Set global applicability.
  ReplaceTest = regEx.Replace(str1, replStr)   ' Make replacement.
End Function

Function aaaTable(xdoc)
   Dim regEx, Match, Matches   ' Create variable.
   Set regEx = New RegExp   ' Create a regular expression.
   regEx.Pattern = "<table([^>]*)>"   ' Set pattern.
   regEx.IgnoreCase = True   ' Set case insensitivity.
   regEx.Global = True   ' Set global applicability.
   target = xdoc
   Set Matches = regEx.Execute(target)   ' Execute search.
   For Each Match in Matches   ' Iterate Matches collection.
   		if inStr(lcase(Match.value)," summary=") = 0 then
   			xmImgStr = Match.value
   			xnImgStr = "<table summary=""列表資料""" & mid(xmImgStr,7)
'   			msgbox xnImgStr
   			target = replace(target, xmImgStr, xnImgStr)
   		end if
   Next
   aaaTable = target
End Function

Function aaaImage(xdoc, altStr)
   Dim regEx, Match, Matches   ' Create variable.
   Set regEx = New RegExp   ' Create a regular expression.
   regEx.Pattern = "<img[^>]*>"   ' Set pattern.
   regEx.IgnoreCase = True   ' Set case insensitivity.
   regEx.Global = True   ' Set global applicability.
   target = xdoc
   Set Matches = regEx.Execute(target)   ' Execute search.
   For Each Match in Matches   ' Iterate Matches collection.
   		if inStr(lcase(Match.value)," alt=") = 0 then
   			xmImgStr = Match.value
   			xnImgStr = "<IMG alt=""" & altStr & """" & mid(xmImgStr,5)
   			msgbox xnImgStr
   			target = replace(target, xmImgStr, xnImgStr)
   		end if
   Next
   aaaImage = target
End Function
</script>

</BODY></HTML>
