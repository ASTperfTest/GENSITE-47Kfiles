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
'formFunction="add"
%>
<html>
<body style="margin-left:0; margin-top:3;margin-right:0">
<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<table cellspacing=2 border=1>
<TR><TH>標題</TH>
<TD><INPUT name="xuf_fieldLabel" size="30"></TD>
<TR><TH>欄位名稱</TH>
<TD><INPUT name="xuf_fieldName" size="20"></TD>
</TR>
<TR><TH>資料型別</TH>
<TD><SELECT name="xuf_dataType" size="1">
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'htDdataType' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select></TD>
</TR>
<TR><TH>輸入方式</TH>
<TD><SELECT name="xuf_inputType" size="1">
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'htDinputType' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select></TD>
</TR>
<TR><TH>長度</TH>
<TD><INPUT name="xuf_inputLen" size="2"></TD>
</TR>
<TR><TH>必填</TH>
<TD>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'boolYN' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

<input type="radio" name="xuf_canNull" value="<%=RSS(0)%>" <%=pdxc%>>
			<%=RSS(1)%>&nbsp;&nbsp;
			<%	RSS.movenext
			wend%>
</TR>
</TABLE>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
<%if formFunction = "edit" then %>
        <%if (HTProgRight AND 8) <> 0 then %>
               <input type=button value ="編修存檔" class="cbutton" onClick="formModSubmit()">
        <% End If %>
        
       <% if (HTProgRight AND 16) <> 0 then %>
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>           
        <input type=button value ="重　填" class="cbutton" onClick="resetForm()">

<% Else '-- add ---%>          
          <%if (HTProgRight AND 4)<>0 then %>
               <input type=button value ="新增存檔" class="cbutton" onClick="formAddSubmit()">
          <%end if%>     
          
          <%if (HTProgRight AND 4)<>0  then %>
              <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
          <%end if%>    

<% End If %>
 </td></tr>
</table>
<% 
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

  	set htPageDom = session("hyXFormSpec")
   	set dsRoot = htPageDom.selectSingleNode("//DataSchemaDef") 	

%>
</body>
<script language=vbs>
	set htPageDom = CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	LoadXML = "ws/processparamField.asp?fname="

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

sub document_onClick
	set srcItem = window.event.srcElement
  if srcItem.tagName = "SPAN" then
'		msgBox srcItem.id
		grid = srcItem.id
'		msgBox parent.parent.frames("fedit").document.all.tags("SPAN").length
'		set allSpan = parent.parent.frames("fedit").document.all.tags("SPAN")
'		for each xs in allSpan
'			if xs.id = grid then
'				srcItem.outerHTML = ""
'				exit sub
'			end if
'		next
	xStyle=" id="&grid&" style=""background-color:#EEE;"">"
	xv = htPageDom.load(LoadXML&mid(grid,4))
  	if htPageDom.parseError.reason <> "" then 
    		msgBox(LoadXML&mid(grid,4) & " htPageDom parseError on line " &  htPageDom.parseError.line)
    		msgBox("Reasonxx: " &  htPageDom.parseError.reason)
    		exit sub
  	end if
	writeCodeStr = nullText(htPageDom.selectSingleNode("//pFieldHTML"))
	xHTML = "<SPAN "&xstyle&writeCodeStr&"</SPAN><BR/>"
'	call parent.parent.frames("fedit").insertField (xHTML,grid)
		window.opener.doInsertField writeCodeStr,grid
		window.close


  end if
end sub

Sub formAddSubmit()	
    
'----- 用戶端Submit前的檢查碼放在這裡 ，如下例 not valid 時 exit sub 離開------------
'msg1 = "請務必填寫「客戶編號」，不得為空白！"
'msg2 = "請務必填寫「客戶名稱」，不得為空白！"
'
'  If reg.pfx_ClientID.value = Empty Then
'     MsgBox msg1, 64, "Sorry!"
'     settab 0
'     reg.pfx_ClientID.focus
'     Exit Sub
'  End if
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
	IF reg.xuf_fieldLabel.value = Empty Then
		MsgBox replace(nMsg,"{0}","標題"), 64, "Sorry!"
		reg.xuf_fieldLabel.focus
		exit sub
	END IF
	IF reg.xuf_fieldName.value = Empty Then
		MsgBox replace(nMsg,"{0}","資料欄位名稱"), 64, "Sorry!"
		reg.xuf_fieldName.focus
		exit sub
	END IF
	IF reg.xuf_dataType.value = Empty Then
		MsgBox replace(nMsg,"{0}","資料型別"), 64, "Sorry!"
		reg.xuf_dataType.focus
		exit sub
	END IF
	IF (reg.xuf_inputLen.value <> "") AND (NOT isNumeric(reg.xuf_inputLen.value)) Then
		MsgBox replace(iMsg,"{0}","輸入長度"), 64, "Sorry!"
		reg.htx_xdataLen.focus
		exit sub
	END IF	

  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>
</html>