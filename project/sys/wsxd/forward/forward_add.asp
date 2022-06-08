<% 
function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function
	
	xItem = request.QueryString("xItem")
	if xItem <>"" then xItem= "xItem=" & xItem
	CtNode = request.QueryString("CtNode")
	if CtNode <>"" then CtNode = "&CtNode=" & CtNode
	mp = request.QueryString("mp")
	if mp="" then mp ="1"
	xmp=mp	
	mp = "&mp=" & mp
	
	xctUrl = xItem & CtNode & mp
	
	
 %>
 <h4>轉寄文章</h4>
<p>請填寫下列輸入欄位，即可轉寄文章。</p>
<table class="cptable" summary="排版表格">
<form name="fm_area" method="post" onSubmit="return checkform(<%=xmp%>);">
<input name="ctUrl" type="hidden" value="<%=deAmp(xctUrl)%>"/>
<tr>
<th scope="row" width="22%"><label for="email">收件人e-Mail：</label></th>
<td><input name="email" id="email" type="text" size="30"/> </td>
</tr>
<tr>
<th scope="row"><label for="name">我的名字：</label></th>
<td><input name="name" type="text" id="name" size="30"/></td>
</tr>
<tr>
<th scope="row"><label for="email2">我的e-Mail：</label></th>
<td><input name="email2" id="email2" type="text" size="30"/> </td>
</tr>
<tr style='display:none;'>
<th scope="row"><label for="message">給朋友的話：</label></th>
<td><textarea name="message" id="message" cols="35" rows="10"></textarea>
</td>
</tr>
<tr>
<th scope="row">&amp;nbsp;</th>
<td><input name="send" type="submit" value="確定送出" class="button"/>  
  <input name="reset" type="reset" value="重新填寫" class="button"/></td>
</tr>

</form>
</table>
<script LANGUAGE="JavaScript">
<!--
function checkform(xmp){
	var fm = document.fm_area;
	var flag=false;
	var filter=/^([\w-]+(?:\.[\w-]+)*)@((?:[\w-]+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$/i;

	if (fm.email.value == "") {
		alert("請輸入收件人電子信箱!"); fm.email.focus();  return flag;
	}
	if(fm.email2.value =="") {
		alert("請輸入寄件人電子信箱!"); fm.email2.focus();  return flag;
	}
	if (!filter.test(fm.email.value)) {
		alert("請輸入合法的收件人電子郵件"); fm.email.focus(); return flag;
	}
	if (!filter.test(fm.email2.value)) {
		alert("請輸入合法的寄件人電子郵件"); fm.email2.focus(); return flag;
	}		
	
	fm.action="forward_send.asp";	

}


//-->
</script>

