<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap=""
HTProgFunc=""
HTProgCode="KForumlist"
HTProgPrefix=""
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<%
	Dim queryType : queryType = request.querystring("queryType")
	Dim queryText : queryText = request.querystring("queryText")
		
	sql = "SELECT account, realname, member_org FROM Member WHERE (id_type3 = 1) AND (status = 'Y') "
	if queryType = "0" then
		sql = sql & "AND ( realname LIKE '%" & queryText & "%' OR KMcat LIKE '%" & queryText & "%' " & _
								"OR keyword LIKE '%" & queryText & "%' OR member_org LIKE '%" & queryText & "%' ) "
	elseif queryType = "1" then
		sql = sql & "AND realname LIKE '%" & queryText & "%' "
	elseif queryType = "2" then
		sql = sql & "AND KMcat LIKE '%" & queryText & "%' "
	elseif queryType = "3" then
		sql = sql & "AND keyword LIKE '%" & queryText & "%' "
	elseif queryType = "4" then
		sql = sql & "AND member_org LIKE '%" & queryText & "%' "
	end if 
	sql = sql & "ORDER BY account"
	
	set rs = conn.execute(sql)
	if rs.eof then
		showDoneBox "此條件下無資料"
	else
		showForm
	end if
%>
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
				history.back()
			</script>
    </body>
  </html>
<% End sub %>
<% sub showForm %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>專家選擇清單</title>
<link href="/KnowledgeForum/css/selectList.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="/KnowledgeForum/moveloc.js"></script>
        
<script language="javascript">

var pickElement = top.opener.document.getElementById("experts"); 
var olds = new Array();

window.onload = init;

String.prototype.trim= function() {  
    return this.replace(/(^\s*)|(\s*$)/g, "");  
}
//---initial---
function init(){
	var oldFields = top.opener.document.getElementById("expertSelected").value;	
	var oldFieldsValue = top.opener.document.getElementById("experts").value;
	fields = oldFields.split(",");
	fieldsValue = oldFieldsValue.split(",");	
	for ( i = 0; i < fields.length; i++ ) {		
	  var old = new Array(fields[i].trim(), fieldsValue[i].trim());
		olds[i] = old;		
	}
	copyoldIDV(document.getElementById("mylocation"),olds);
	loadLoc(document.getElementById("mylocation"));
}

//---確認送出---
function pick(){
	var myloc = document.getElementById("mylocation");	
	var picker = "";
	var pickerVal = "";
	for ( i = 0; i < myloc.length; i++ ) {
		var myopt = myloc.options[i];
		picker += myopt.text+",";
		pickerVal += myopt.value+",";
	}
	picker = picker.substring(0,picker.lastIndexOf(","));
	pickerVal = pickerVal.substring(0,pickerVal.lastIndexOf(","));	
	top.opener.document.getElementById("expertSelected").value = picker;	
	top.opener.document.getElementById("experts").value = pickerVal;	
	window.top.close();
}
</script>
</head>

<body>
	<div class="warp">
		<h1>專家選擇清單</h1>
		<form>
			<div class="inlineSearch">
				<select name="queryType">
					<option value="0" <% if queryType = "0" then %> selected <% end if %>>全部</option>
					<option value="1" <% if queryType = "1" then %> selected <% end if %>>真實姓名</option>
					<option value="2" <% if queryType = "2" then %> selected <% end if %>>研究領域及專長</option>
					<option value="3" <% if queryType = "3" then %> selected <% end if %>>專長關鍵字</option>
					<option value="4" <% if queryType = "4" then %> selected <% end if %>>所屬機關</option>
				</select>
				<input name="queryText" type="text" size="30" class="text" value="<%=queryText%>" />
				<input name="queryBtn" type="image" src="/KnowledgeForum/images/btn_go.gif" alt="查詢" title="查詢" onclick="QueryClick()" />
			</div>

			<div class="optionList">
				<select class="all" name="location" multiple size="15" onDblClick="JavaScript:addloc(document.all.location,document.all.mylocation)">
				<%					
					while not rs.eof
						response.write "<option value=""" & trim(rs("account")) & """>" & trim(rs("realname")) & "(" & trim(rs("account")) & ")【" & trim(rs("member_org")) & "】</option>"
						rs.movenext
					wend
					rs.close
					set rs = nothing
				%>
				</select>
				<div class="picker">
					<a name="AddLoc" onclick="JavaScript:addloc(document.all.location,document.all.mylocation)"><img src="/KnowledgeForum/images/btn_add.gif" alt="加入" /></a>
					<a name="DelLoc" onclick="JavaScript:delloc(document.all.location,document.all.mylocation)"><img src="/KnowledgeForum/images/btn_remove.gif" alt="移除" /></a>
				</div>
				<select class="selected" name="mylocation" size="15" multiple>
				</select>				
			</div>
			<div class="button">
				<input name="SubmitBtn" type="image" src="/KnowledgeForum/images/btn_submit.gif" alt="確定送出" onclick="pick()" />
			</div>
		</form>
	</div>
</body>
</html>
<script language="javascript">
	function QueryClick()
	{
		var myloc = document.getElementById("mylocation");	
		var picker = "";
		var pickerVal = "";
		for ( i = 0; i < myloc.length; i++ ) {
			var myopt = myloc.options[i];
			picker += myopt.text+",";
			pickerVal += myopt.value+",";
		}
		picker = picker.substring(0,picker.lastIndexOf(","));
		pickerVal = pickerVal.substring(0,pickerVal.lastIndexOf(","));	
		top.opener.document.getElementById("expertSelected").value = picker;	
		top.opener.document.getElementById("experts").value = pickerVal;	
		
		window.location.href = "/KnowledgeForum/ExpertSelect.asp?queryType=" + document.getElementById("queryType").value + 
													 "&queryText=" + document.getElementById("queryText").value;
	}
</script>
<% end sub %>