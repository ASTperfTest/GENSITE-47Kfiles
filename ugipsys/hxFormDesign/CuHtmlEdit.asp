<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="動態表單設計"
HTProgFunc="編修"
HTUploadPath="/"
HTProgCode="HF011"
HTProgPrefix="DsdXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
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

function qqRS(xValue)
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
		qqRS = replace(xqqRS,chr(10),chr(34)&"&vbCRLF&"&chr(34))
	end if
end function 

	siteStr = "http://" & request.serverVariables("SERVER_NAME") 
	myURL = "http://" & request.serverVariables("SERVER_NAME") & request.serverVariables("URL") & "?" & request.serverVariables("QUERY_STRING")
'	response.write 	myURL & "<HR>"
	for each x in request.serverVariables
'		response.write x & "=>" & request(x) & "<BR>"
	next
	
	if request.form("submitTask") = "UPDATE" then
		iCuItem = request.form("htx_iCuItem")
		xBody = request("htx_xBody")
	  	set htPageDom = session("hyXFormSpec")
'	  	htPageDom.selectSingleNode("//pHTML").text = "<![CDATA[" & xBody & "]]>"
	  	htPageDom.selectSingleNode("//pHTML").text = xBody
'	  	htPageDom.selectSingleNode("//pHTML").replaceChild htPageDom.createCDATASection(xBody), htPageDom.selectSingleNode("//pHTML").childNodes.Item(0)

		if isNumeric(iCuItem) then
			sql = "UPDATE CuDtxDFormSpec SET xmlSpec = " & pkStr(htPageDom.xml,"") _
				& ", pHTML = " & pkStr(xBody,"") _
				& " WHERE giCUItem=" & iCuItem
'			response.write sql
'			response.end
			conn.execute sql
'			sql = "UPDATE CuDtGeneric SET xBody = " & pkStr(xBody,"") _
'				& " WHERE iCUItem=" & iCuItem
'			conn.execute sql
	htPageDom.save(Server.MapPath(".")&"\hf_" & session("hyFormID") & ".xml")
%>      <script language=vbs>
	       window.close
      </script>
<%		response.end
		
		else
			response.end
		end if
	end if

	hyFormID = request("iCuItem")
	fSql = "SELECT sTitle, xmlSpec FROM CuDTxDFormSpec AS g JOIN CuDtGeneric AS h ON g.giCuItem=h.iCuItem " _
		& " WHERE h.iCuItem=" & pkStr(hyFormID,"")
	set RSform = conn.execute(fSql)
	if RSform.eof then
%>      <script language=vbs>
           msgbox "無此資料!"
	       window.history.back
      </script>
<%		response.end
	end if

  	session("hyFormID") = hyFormID
	xmlSpec = RSform("xmlSpec")
	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	if isnull(RSform("xmlSpec")) OR RSform("xmlSpec") = "" then
		LoadXML = server.MapPath("hf_"&hyFormID&".xml")
		xv = htPageDom.load(LoadXML)
	  if not xv then
		LoadXML = server.MapPath("hyForm0.xml")
		xv = htPageDom.load(LoadXML)
'		response.write xv & "hyForm0.xml<HR>"
	  	if htPageDom.parseError.reason <> "" then 
	    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
	    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
	    		Response.End()
	    end if
	  end if
	else
		
		xv = htPageDom.loadxml(xmlSpec)
'		response.write xv & "RSform<HR>"
	  	if htPageDom.parseError.reason <> "" then 
	    		Response.Write("htStr parseError on line " &  htPageDom.parseError.line)
	    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
	    		Response.End()
	    end if
	end if
  	set session("hyXFormSpec") = htPageDom
  	set refModel = htPageDom.selectSingleNode("//dsTable")
  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")


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
			el.title = el.title + "(另開視窗)";
			}
	}
}
function doAttachLink(){
//Editor.focus();
var fcolor=showModalDialog("editor_attach.asp?iCUItem=<%=request("iCuItem")%>",false,"dialogWidth:500px;dialogHeight:200px;status:0;");
if (fcolor!=undefined){
doc.execCommand('CreateLink',false,'public/Attachment/'+fcolor);
Editor.focus();
}}
function doPageLink(){
//Editor.focus();
var fcolor=showModalDialog("editor_pgLink.asp?iCUItem=<%=request("iCuItem")%>",false,"dialogWidth:500px;dialogHeight:200px;status:0;");
if (fcolor!=undefined){
doc.execCommand('CreateLink',false,'webPage.asp?CuItem='+fcolor);
Editor.focus();
}}
function doInsertImage(MMOPath, xTitle){
	Editor.focus();
	var txRange = doc.selection.createRange();
//	alert(txRange.htmlText);
	txRange.pasteHTML(txRange.htmlText + '<IMG SRC="' + MMOPath + '" alt="' + xTitle + '" />');
	txRange.select();
	Editor.focus();
// doc.execCommand('InsertImage',false,MMOPath);
//	var txRange = doc.selection.createRange();
//	alert(txRange.htmlText);
Editor.focus();
}


// New funtions
function xdoRemoveFormat(){
	doc.execCommand('RemoveFormat');
	var txRange = doc.selection.createRange();
	alert(txRange.htmlText);
	Editor.focus();
}

function xdoShowDetails() {
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
if(vx.checked){
SeePreview();
}
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
</SCRIPT>

<META content="MSHTML 5.00.3315.2870" name=GENERATOR></HEAD>
<BODY onkeyup=AutoPreview(); onLoad="SeePreview();Preview.focus();">
<div id="FuncName">
	<h1><%=HTProgCap%></h1>
	<div id="Nav"><a href="DsdXMLEdit.asp?iCuItem=">回上頁</a></div>
	<div id="ClearFloat"></div>
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
		  </tr>
	    </table>
		<table cellspacing="0" id="FuncIcon">
	      <tr>
	        <td class="Holder"><img src="editor/images/main_holder.gif"></td>
	        <td class="Lead"><img src="editor/images/form.gif" alt="插入表單元件"></td>
	        <td class="Icon"><a href="javascript:;" title="欄位條列"><img src="editor/images/details.gif"
	        onclick="VBS: doShowDetails()"></a></td>
	        <td class="Icon"><a href="javascript:;" title="輸入框"><img src="editor/images/textbox.gif"
	        onclick="fieldInsert('textbox');"></a></td>
	        <td class="Icon"><a href="javascript:;" title="文字盒"><img src="editor/images/textarea.gif"
	        onclick="fieldInsert('textarea');"></a></td>
	        <td class="Icon"><a href="javascript:;" title="勾選框"><img src="editor/images/checkbox.gif"
	        onclick="fieldInsert('checkbox')"></a></td>
	        <td class="Icon"><a href="javascript:;" title="圓鈕"><img src="editor/images/radio.gif"
	        onclick="fieldInsert('radio')"></a></td>
	        <td class="Icon"><a href="javascript:;" title="下拉選項"><img src="editor/images/pulldown.gif"
	        onclick="fieldInsert('pulldown')"></a></td>
	        <td class="Icon"><a href="javascript:;" title="表列點單"><img src="editor/images/listbox.gif"
	        onclick="fieldInsert('listbox')"></a></td>
	        <td class="Icon"><a href="javascript:;" title="aaaForm"><img src="HTMLeditor/editor_e.gif"
	        onclick="VBS: aaaForm()"></a></td>
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
		<table cellspacing="0" id="FuncIcon">
	      <tr>
	        <td class="Holder"><img src="editor/images/main_holder.gif"></td>
	        <td class="Lead"><img src="editor/images/icon_table.gif" alt="插入表格"></td>
	        <td class="Icon"><a href="javascript:;" title="資料標題於上列"><img src="editor/images/icon_table_HTop.gif" 
	        onClick="MM_openBrWindow('edit_table.asp','','scrollbars=yes,width=560,height=400')"></a></td>
	        <td class="Icon"><a href="javascript:;" title="資料標題於左欄"><img src="editor/images/icon_table_HLeft.gif"></a></td>
	        <td class="Icon"><a href="javascript:;" title="資料標題於上及左"><img src="editor/images/icon_table_HBoth.gif"></a></td>
	        <td class="Icon"><a href="javascript:;" title="無資料標題"><img src="editor/images/icon_table_HNone.gif"></a></td>
	      </tr>
		</table>
		<table cellspacing="0" id="FuncIcon">
	      <tr>
	        <td class="Holder"><img src="editor/images/main_holder.gif"></td>
	        <td class="Lead"><img src="editor/images/icon_media.gif" alt="插入多媒體物件"></td>
	        <td class="Icon"><a href="javascript:;" title="插入圖檔"><img src="editor/images/icon_media_image.gif" onClick="MM_openBrWindow('CuMMOQuery.asp','','scrollbars=yes,width=700,height=500')"></a></td>
	        <td class="Icon"><a href="javascript:;" title="插入聲音檔"><img src="editor/images/icon_media_audio.gif"></a></td>
	        <td class="Icon"><a href="javascript:;" title="插入Flash檔"><img src="editor/images/icon_media_flash.gif"></a></td>
	        <td class="Icon"><a href="javascript:;" title="插入影片檔"><img src="editor/images/icon_media_video.gif"></a></td>
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
	          <td class="Icon"><a href="javascript:;" title="連結其他單元"><img src="editor/images/icon_link_unit.gif" width="19" height="14"></a></td>
	          <td class="Icon"><a href="javascript:;" title="連結其他單元網頁"><img src="editor/images/icon_link_unitpage.gif" width="18" height="14"></a></td>
		    </tr>
	    </table>
    </td>
  </tr>
  <tr>
    <td class="PageView">頁面呈現<br>
      <IFRAME id=Editor marginWidth=1 src="about:blank"></IFRAME>
<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<INPUT TYPE=hidden name=htx_iCuItem value="<%=hyFormID%>">
<INPUT TYPE=hidden name=htx_xBody  value="">
<input type=button value ="編修存檔" class="InputButton" onClick="formModSubmit()">
   </FORM>
		</td>
  </tr>
  <tr>
    <td class="CodeView">原始碼
      
      <br>    
	<TEXTAREA id=Preview xreadonly="readonly" wrap="virtual"></TEXTAREA>
		</td>
  </tr>
</table>

<TABLE border=1 borderColorDark=#efece8 borderColorLight=#888070 cellPadding=0 
cellSpacing=0 width="100%">
  <TBODY>
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
          <TD><IMG src="HTMLeditor/editor_h.gif"></TD>
          <TD><INPUT CHECKED name=vm onclick=EditMode(); title=編輯模式 
          type=radio></TD>
          <TD>
            <DIV title=編輯模式><IMG src="HTMLeditor/editor_e.gif"></DIV></TD>
          <TD>&nbsp;</TD>
          <TD><INPUT name=vm onclick=PreviewMode(); title=預覽模式 type=radio></TD>
          <TD>
            <DIV title=預覽模式><IMG src="HTMLeditor/editor_p.gif"></DIV></TD>
          <TD>
            <DIV title=自動預覽><IMG 
        src="HTMLeditor/editor_a.gif"></DIV></TD></TR></TBODY></TABLE></TD></TR>
   <TR><TH></TD></TR> 
    </TBODY></TABLE></CENTER>
<script language=vbs>
Preview.value = "<%=qqrs(nullText(htPageDom.selectSingleNode("//pHTML")))%>"
</script>
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
 
' 	doc.body.innerHTML = purifyHTML(doc.body.innerHTML)
 	reg.htx_xBody.value = doc.body.innerHTML
' 	msgBox  reg.htx_xBody.value
 
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
'		if x.alt="" then	x.alt=reg.htx_sTitle.value + " "
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
	xstr = ReplaceTest("<FONT[^>]*size=[^>]*>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
'	xstr = ReplaceTest("</FONT>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("<\?xml:namespace[^>]*>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("<o:p></o:p>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("<table([^>]*)>", "<table $1 summary=""列表資料 "">", xstr)	'-- mso-bidi-font-size: 10.0pt;
'	msgBox xstr
	purifyHTML = xstr
	
end function

function xpurifyHTML(xHTMLstr)
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
	xstr = ReplaceTest("<table([^>]*)>", "<table $1 summary=""列表資料 "">", xstr)	'-- mso-bidi-font-size: 10.0pt;
'	msgBox xstr
	xpurifyHTML = xstr
	
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

sub xdoShowDetails()
'	msgBox RegExpTest("<img[^>]*>",doc.body.innerHTML)
	doc.body.innerHTML = aaaImage(doc.body.innerHTML, "xxx")
end sub

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

sub insertField (xhtm, grid)
	Editor.document.body.insertAdjacentHTML "BeforeEnd",xhtm
	Editor.focus
	set txRange = doc.selection.createRange()
'	txRange.pasteHTML(xhtm)
	txRange.select()
'	Editor.focus()
end sub

sub doInsertField (xhtm, grid)
	Editor.focus()
	set txRange = doc.selection.createRange()
'	txRange.pasteHTML(txRange.htmlText + '<IMG SRC="' + MMOPath + '" alt="' + xTitle + '" />');
	txRange.pasteHTML(xhtm)
	txRange.select()
	Editor.focus()
end sub

sub doShowDetails
	window.open "pickField.asp","","scrollbars=yes,width=700,height=500"
end sub

sub fieldInsert(fieldtype)
	window.open "hfDFieldAdd.asp?fkind=" & fieldType,"","scrollbars=yes,width=700,height=500"
end sub

function xfCode(fname, xtIndex)
	set xfDom = CreateObject("MICROSOFT.XMLDOM")
	xfDom.async = false
	LoadXML = "ws/processparamField.asp?fname="&fname&"&xtIndex="&xtIndex
'	msgbox loadXML
	xv = xfDom.load(LoadXML)
  	if xfDom.parseError.reason <> "" then 
    		msgBox(LoadXML & " xfDom parseError on line " &  xfDom.parseError.line)
    		msgBox("Reason: " &  xfDom.parseError.reason)
    		exit function
  	end if
	xfCode = nullText(xfDom.selectSingleNode("//pFieldHTML"))
end function

sub aaaForm()
	xTabIndex = 0
	for each xa in doc.body.all
		xTagName = ucase(xa.tagName)
		if xTagName="INPUT" OR xTagName="SELECT" OR xTagName="TEXTAREA" then
			xTabIndex = xTabIndex + 1
			xFieldName = mid(xa.name,5)
			'if xTabIndex < 2	then
			'msgbox xfCode(xFieldName,xTabIndex)
			xa.outerHTML = xfCode(xFieldName,xTabIndex)
			'end if
		end if
	next
end sub

Function xaaaForm(xdoc)
   Dim regEx, Match, Matches   ' Create variable.
   Set regEx = New RegExp   ' Create a regular expression.
   regEx.Pattern = "<((INPUT)|(SELECT))[^>]*>"   ' Set pattern.
   regEx.IgnoreCase = True   ' Set case insensitivity.
   regEx.Global = True   ' Set global applicability.
   target = xdoc
   Set Matches = regEx.Execute(target)   ' Execute search.
   For Each Match in Matches   ' Iterate Matches collection.
	xField = regExpFind(match.value, "name=([^> \""]*)")
	response.write xField & "==>"
   	response.write Match.value & "<BR/>" & vbCRLF
   		if inStr(lcase(Match.value)," xxalt=") > 0 then
   			xmImgStr = Match.value
   			xaltStr = altStr
   			if instr(lcase(Match.value)," src=../../img/1.jpg") > 0 then	xaltStr = "方方土"
   			xnImgStr = "<IMG alt=""" & xaltStr & """" & mid(xmImgStr,5)
   			target = replace(target, xmImgStr, xnImgStr)
   		end if
   Next
   aaaINPUT = target
End Function

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function


</script>

</BODY></HTML>
