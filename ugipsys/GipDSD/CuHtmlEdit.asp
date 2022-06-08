<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath="/"
HTProgCode="GC1AP1"
HTProgPrefix="DsdXML" %>
<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbFunc.inc" -->
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
		iCuItem = request.form("htx_iCuItem")
		xBody = request("htx_xBody")
'		response.write xBody & "<HR>"
		xBody = replace(xBody, myURL, "")
		xBody = replace(xBody, siteStr& "/GipEdit/", "")
		xBody = replace(xBody, siteStr, "")
		xBody = replace(xBody, "<A href=""http://", "<A target=""_nwMof"" href=""http://")
		xBody = replace(xBody, "<A href=""/public/", "<A target=""_nwMof"" href=""/public/")
		xBody = replace(xBody, siteStr, "")
		
'		response.write xBody & "<HR>"
'		response.end
		if isNumeric(iCuItem) then
			sql = "UPDATE CuDTGeneric SET xBody = " & pkStr(xBody,"") _
				& " WHERE iCuItem=" & iCuItem
			conn.execute sql
			response.redirect "DsdXMLEdit.asp?iCuItem=" & iCuItem
'			response.write sql & "<HR>"
'			response.end
		
		
		else
			response.end
		end if
	end if


  	set htPageDom = session("codeXMLSpec")
  	set refModel = htPageDom.selectSingleNode("//dsTable")
  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")

	sqlCom = "SELECT ghtx.* FROM CuDtGeneric AS ghtx "_
		& " WHERE ghtx.iCuItem=" & pkStr(request.queryString("iCuItem"),"")
	Set RSreg = Conn.execute(sqlcom)
	pKey = "iCuItem=" & RSreg("iCuItem")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>HTML 編輯器</TITLE>
<META content="text/html; charset=utf-8" http-equiv=Content-Type>
<STYLE>BODY {
	FONT-FAMILY: 細明體; FONT-SIZE: 12px
}
TD {
	FONT-FAMILY: 細明體; FONT-SIZE: 12px
}
BUTTON {
	FONT-FAMILY: 細明體; FONT-SIZE: 12px
}
INPUT {
	FONT-FAMILY: 細明體; FONT-SIZE: 12px
}
DIV {
	BACKGROUND: #d9cec4; BORDER-BOTTOM: #d9cec4 1px solid; BORDER-LEFT: #d9cec4 1px solid; BORDER-RIGHT: #d9cec4 1px solid; BORDER-TOP: #d9cec4 1px solid; CURSOR: default; HEIGHT: 20px; TEXT-ALIGN: center; WIDTH: 24px; borderColor: #ffffff
}
.function {
	WIDTH: 80px
}
</STYLE>

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
Editor.focus();
}
function doPaste(){
doc.execCommand('Paste');
Editor.focus();
}
function doUndo(){
doc.execCommand('Undo');
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
function doCreateLink(){
doc.execCommand('CreateLink');
Editor.focus();
}
function doAttachLink(){
//Editor.focus();
var fcolor=showModalDialog("editor_attach.asp?iCUItem=<%=request("iCuItem")%>",false,"dialogWidth:500px;dialogHeight:200px;status:0;");
if (fcolor!=undefined){
doc.execCommand('CreateLink',false,'/public/Attachment/'+fcolor);
Editor.focus();
}}
function doPageLink(){
//Editor.focus();
var fcolor=showModalDialog("editor_pgLink.asp?iCUItem=<%=request("iCuItem")%>",false,"dialogWidth:500px;dialogHeight:200px;status:0;");
if (fcolor!=undefined){
doc.execCommand('CreateLink',false,'webPage.asp?CuItem='+fcolor);
Editor.focus();
}}
function doInsertImage(){
Editor.focus();
doc.execCommand('InsertImage','xxx');
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
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td class="FormName" align="left"><%=HTProgCap%>&nbsp;
	    <font size=2>【編輯--主題單元:<%=session("CtUnitName")%> / 單元資料:<%=nullText(htPageDom.selectSingleNode("//tableDesc"))%>】</td>
	    <td class="FormLink" valign="top" align=right>
	       <a href="DsdXMLEdit.asp?iCuItem=<%=RSreg("iCuItem")%>">回前頁</a>
		</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
</TABLE>
<CENTER>
<TABLE border=1 borderColorDark=#efece8 borderColorLight=#888070 cellPadding=0 
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
    <TD><IFRAME id=Editor marginWidth=1 src="about:blank" 
      style="BACKGROUND-COLOR: white; HEIGHT: 300px; WIDTH: 100%"></IFRAME></TD></TR>
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
  <TR>
    <TD><TEXTAREA id=Preview style="BACKGROUND-COLOR: #ffffff; HEIGHT: 80px; OVERFLOW: auto; WIDTH: 100%"><%=htmlMessage(RSreg("xBody"))%></TEXTAREA></TD></TR>
   <TR><TH><form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=htx_iCuItem value="<%=RSreg("iCuItem")%>">
<INPUT TYPE=hidden name=submitTask value="">
<INPUT TYPE=hidden name=htx_xBody  value="">
<input type=button value ="編修存檔" class="cbutton" onClick="formModSubmit()">
   </FORM></TD></TR> 
    </TBODY></TABLE></CENTER>
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
 
	doc.body.innerHTML = purifyHTML(doc.body.innerHTML)
 	reg.htx_xBody.value = doc.body.innerHTML
' 	msgBox  reg.htx_xBody.value
 
  reg.submitTask.value = "UPDATE"
  reg.Submit
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
	xstr = ReplaceTest("\swidth=[0-9]*\s", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("\swidth=[0-9]*>", ">", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("<FONT[^>]*>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("</FONT>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("<\?xml:namespace[^>]*>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
	xstr = ReplaceTest("<o:p></o:p>", "", xstr)	'-- mso-bidi-font-size: 10.0pt;
'	msgBox xstr
	purifyHTML = xstr
	
end function

Function ReplaceTest(patrn, replStr, str1)
  Dim regEx               ' Create variables.
'  str1 = "The quick brown fox jumped over the lazy dog."
  Set regEx = New RegExp            ' Create regular expression.
  regEx.Pattern = patrn            ' Set pattern.
  regEx.IgnoreCase = True            ' Make case insensitive.
  regEx.Global = True   ' Set global applicability.
  ReplaceTest = regEx.Replace(str1, replStr)   ' Make replacement.
End Function

</script>

</BODY></HTML>
