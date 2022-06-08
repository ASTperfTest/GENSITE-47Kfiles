<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	'HTProgCap = "功能管理"
	'HTProgFunc = "問卷調查"
	HTProgCode = "GW1_vote01"
	'HTProgPrefix = "mSession"
	response.expires = 0
%>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<%	
	
	subjectid = request("subjectid")
	
	if subjectid = "" then
		response.write "<script language='JavaScript'>history.go(-1);</script>"
		response.end
	end if
	
	
	sql = "" & _
		" select m011_subject, m011_bdate, m011_edate, m011_haveprize from m011 where " & _
		" m011_subjectid = " & subjectid
	set rs = conn.execute(sql)
	if rs.EOF then
		response.write "<script language='JavaScript'>history.go(-1);</script>"
		response.end
	end if
	
	
	set ts = conn.execute("select count(*) from m014 where m014_subjectid = " & subjectid)
	if ts(0) = 0 then
		response.write "<script language='JavaScript'>alert('目前尚無任何答題資料！');history.go(-1);</script>"
		response.end
	end if
	
	
	select2 = trim(request("select2"))
	select3 = trim(request("select3"))
	
	other_sql = ""
	if select3 <> "" and select2 <> "" then
		if select2 = "sex" then
			other_sql = " and m014_sex = '" & select3 & "' "
		elseif select2 = "age" then
			other_sql = " and m014_age = " & select3 & " "
		elseif select2 = "addrarea" then
			other_sql = " and m014_addrarea = " & select3 & " "
		elseif select2 = "familymember" then
			other_sql = " and m014_familymember = " & select3 & " "
		elseif select2 = "money" then
			other_sql = " and m014_money = " & select3 & " "
		elseif select2 = "job" then
			other_sql = " and m014_job = '" & select3 & "' "
		elseif select2 = "edu" then
			other_sql = " and m014_edu = '" & select3 & "' "
		end if
	end if
	
	sql = "" & _
		" select m014_name, m014_sex, m014_idnumber, m014_email, m014_age, " & _
		" m014_addrarea, " & _
		" m014_familymember, isNull(m014_money, 0) m014_money, m014_job, " & _
		" isnull(m014_edu, '') m014_edu " & _
		" from m014 " & _
		" where m014_subjectid = " & replace(subjectid, "'", "''") & other_sql & _
		" order by m014_id " 
	set list = conn.execute(sql)
	
	
	sql = "select count(*) from m014 where m014_subjectid = " & subjectid & other_sql
	set ts = conn.Execute(sql)
	totalrecord = ts(0)
	pagecount = 10
	if totalrecord > 0 then              'ts代表筆數
		totalpage = totalrecord \ pagecount
		if (totalrecord mod pagecount) <> 0 then
			totalpage = totalpage + 1
		end if
	else
		totalpage = 1
	end if
	
	if request("page") = empty then
		page = 1
	else
		page = request("page")
	end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF">
<p>問卷調查分析</p>
<table border="1" cellspacing="0" cellpadding="2" width="100%" bordercolordark="#FFFFFF">
  <tr> 
    <td colspan="9">統計項目： 
      <input type="button" value="答案累計統計" onClick="javascript:location.replace('02_result.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
<%
	'if rs("m011_haveprize") = "1" then
%>
      <input type="button" value="調查對象分析" onClick="javascript:location.replace('02_group.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
      <input type="button" value="調查對象明細" onClick="javascript:location.replace('02_namelist.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
<%
	'end if
%>
      <input type="button" value="意見回覆" onClick="javascript:location.replace('02_feedback_note.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
    </td>
  </tr>
  
  <tr>
    <td colspan="9">
      <input type="button" value="資料匯出" onClick="javascript:location.replace('02_namelist_exl.asp?subjectid=<%= subjectid %>');">
      <font color="#993300">[請按資料匯出之後,會出現另存新檔,副檔名請存成xls檔即可]</font>
    </td>
  </tr>  
  <tr><td colspan="9">起訖時間：<%= DateValue(rs("m011_bdate")) %> ~ <%= DateValue(rs("m011_edate")) %></td></tr>
  <tr><td colspan="9">調查主題：<%= trim(rs("m011_subject")) %></tr>  
  
  <form method="post" action="02_namelist.asp">
  <input type="hidden" name="functype" value="<%= functype %>">
  <input type="hidden" name="subjectid" value="<%= subjectid %>">
  <tr> 
    <td colspan="3">列出符合條件： 
      <select name="select2" onchange='javascript:select2_Onchange(this, this.form.select3)'>
        <option value="sex"<% if select2 = "sex" then %> selected<% end if %>>性別</option>
        <option value="age"<% if select2 = "age" then %> selected<% end if %>>年齡</option>
        <option value="addrarea"<% if select2 = "addrarea" then %> selected<% end if %>>縣市</option>
        <option value="familymember"<% if select2 = "familymember" then %> selected<% end if %>>成員</option>
        <option value="money"<% if select2 = "money" then %> selected<% end if %>>收入</option>
        <option value="job"<% if select2 = "job" then %> selected<% end if %>>職業</option>
        <option value="edu"<% if select2 = "edu" then %> selected<% end if %>>學歷</option>
      </select>
      <select name="select3">
<%
	if select2 = "sex" or select2 = "" then
%>
        <option value=""<% if select3 = "" then %> selected<% end if %>>不拘</option>
        <option value="M"<% if select3 = "M" then %> selected<% end if %>>男</option>
        <option value="F"<% if select3 = "F" then %> selected<% end if %>>女</option>
<%
	elseif select2 = "age" then
%>
        <option value=""<% if select3 = "" then %> selected<% end if %>>不拘</option>
        <option value="1"<% if select3 = "1" then %> selected<% end if %>>0到11歲以下</option>
        <option value="2"<% if select3 = "2" then %> selected<% end if %>>12到17歲</option>
        <option value="3"<% if select3 = "3" then %> selected<% end if %>>18到23歲</option>
        <option value="4"<% if select3 = "4" then %> selected<% end if %>>24到29歲</option>
        <option value="5"<% if select3 = "5" then %> selected<% end if %>>30到39歲</option>
        <option value="6"<% if select3 = "6" then %> selected<% end if %>>40到49歲</option>
        <option value="7"<% if select3 = "7" then %> selected<% end if %>>50到59歲</option>
        <option value="8"<% if select3 = "8" then %> selected<% end if %>>60歲以上</option>
<%	
	elseif select2 = "addrarea" then
%>
        <option value=""<% if select3 = "" then %> selected<% end if %>>不拘</option>
        <option value="1"<% if select3 = "1" then %> selected<% end if %>>臺北縣</option>
        <option value="2"<% if select3 = "2" then %> selected<% end if %>>宜蘭縣</option>
        <option value="3"<% if select3 = "3" then %> selected<% end if %>>桃園縣</option>
        <option value="4"<% if select3 = "4" then %> selected<% end if %>>新竹縣</option>
        <option value="5"<% if select3 = "5" then %> selected<% end if %>>苗栗縣</option>
        <option value="6"<% if select3 = "6" then %> selected<% end if %>>臺中縣</option>
        <option value="7"<% if select3 = "7" then %> selected<% end if %>>彰化縣</option>
        <option value="8"<% if select3 = "8" then %> selected<% end if %>>南投縣</option>
        <option value="9"<% if select3 = "9" then %> selected<% end if %>>雲林縣</option>
        <option value="10"<% if select3 = "10" then %> selected<% end if %>>嘉義縣</option>
        <option value="11"<% if select3 = "11" then %> selected<% end if %>>臺南縣</option>
        <option value="12"<% if select3 = "12" then %> selected<% end if %>>高雄縣</option>
        <option value="13"<% if select3 = "13" then %> selected<% end if %>>屏東縣</option>
        <option value="14"<% if select3 = "14" then %> selected<% end if %>>臺東縣</option>
        <option value="15"<% if select3 = "15" then %> selected<% end if %>>花蓮縣</option>
        <option value="16"<% if select3 = "16" then %> selected<% end if %>>澎湖縣</option>
        <option value="17"<% if select3 = "17" then %> selected<% end if %>>基隆市</option>
        <option value="18"<% if select3 = "18" then %> selected<% end if %>>新竹市</option>
        <option value="19"<% if select3 = "19" then %> selected<% end if %>>臺中市</option>
        <option value="20"<% if select3 = "20" then %> selected<% end if %>>嘉義市</option>
        <option value="21"<% if select3 = "21" then %> selected<% end if %>>臺南市</option>
        <option value="22"<% if select3 = "22" then %> selected<% end if %>>臺北市</option>
        <option value="23"<% if select3 = "23" then %> selected<% end if %>>高雄市</option>
        <option value="24"<% if select3 = "24" then %> selected<% end if %>>金門縣</option>
        <option value="25"<% if select3 = "25" then %> selected<% end if %>>連江縣</option>
<%	
	elseif select2 = "familymember" then
%>
        <option value=""<% if select3 = "" then %> selected<% end if %>>不拘</option>
        <option value="1"<% if select3 = "1" then %> selected<% end if %>>1人</option>
        <option value="2"<% if select3 = "2" then %> selected<% end if %>>2人</option>
        <option value="3"<% if select3 = "3" then %> selected<% end if %>>3人</option>
        <option value="4"<% if select3 = "4" then %> selected<% end if %>>4人</option>
        <option value="5"<% if select3 = "5" then %> selected<% end if %>>5至8人</option>
        <option value="6"<% if select3 = "6" then %> selected<% end if %>>8人以上</option>
<%	
	elseif select2 = "money" then
%>
        <option value=""<% if select3 = "" then %> selected<% end if %>>不拘</option>
        <option value="1"<% if select3 = "1" then %> selected<% end if %>>5000元以下</option>
        <option value="2"<% if select3 = "2" then %> selected<% end if %>>5000到1萬元</option>
        <option value="3"<% if select3 = "3" then %> selected<% end if %>>1萬到3萬元</option>
        <option value="4"<% if select3 = "4" then %> selected<% end if %>>3萬到5萬元</option>
        <option value="5"<% if select3 = "5" then %> selected<% end if %>>5萬到10萬元</option>
        <option value="6"<% if select3 = "6" then %> selected<% end if %>>10萬元以上</option>
<%	
	elseif select2 = "job" then
%>
        <option value=""<% if select3 = "" then %> selected<% end if %>>不拘</option>
        <option value="1"<% if select3 = "1" then %> selected<% end if %>>軍</option>
        <option value="2"<% if select3 = "2" then %> selected<% end if %>>公</option>
        <option value="3"<% if select3 = "3" then %> selected<% end if %>>商</option>
        <option value="4"<% if select3 = "4" then %> selected<% end if %>>教職</option>
        <option value="5"<% if select3 = "5" then %> selected<% end if %>>自由業</option>
        <option value="6"<% if select3 = "6" then %> selected<% end if %>>服務業</option>
        <option value="7"<% if select3 = "7" then %> selected<% end if %>>其他</option>
<%	
	elseif select2 = "edu" then
%>
        <option value=""<% if select3 = "" then %> selected<% end if %>>不拘</option>
        <option value="1"<% if select3 = "1" then %> selected<% end if %>>國小</option>
        <option value="2"<% if select3 = "2" then %> selected<% end if %>>國中</option>
        <option value="3"<% if select3 = "3" then %> selected<% end if %>>高中</option>
        <option value="4"<% if select3 = "4" then %> selected<% end if %>>大學</option>
        <option value="5"<% if select3 = "5" then %> selected<% end if %>>碩士</option>
        <option value="6"<% if select3 = "6" then %> selected<% end if %>>博士</option>
        <option value="7"<% if select3 = "7" then %> selected<% end if %>>其他</option>
<%
	else
%>
        <option value=""<% if select3 = "" then %> selected<% end if %>>不拘</option>
        <option value="M"<% if select3 = "M" then %> selected<% end if %>>男</option>
        <option value="F"<% if select3 = "F" then %> selected<% end if %>>女</option>
<%
	end if
%>
      </select>
      <input type="submit" value="GO">
    </td>
    <td colspan="6" align="right">
      共有 <% =totalpage %> 頁，目前在第 <% =page %> 頁，跳到： 
      <select name="select" onChange='javascript:Page_Onchange(this.form)'>
<%
	n = 1
	While n <= totalpage 
		Response.Write "<option value='" & n & "'"
		If Int(n) = Int(Page) Then 
			Response.Write " selected"
		End If
		Response.Write ">" & n & "</option>"
		n = n + 1
	WEnd
%>
      </select>
      頁【每頁<%= pagecount %>筆】
    </td>
  </tr>

  <tr> 
    <td bgcolor="#EFEFEF" width="12%" nowrap>姓名</td>
    <td bgcolor="#EFEFEF" width="5%" nowrap>性別</td>
    <td bgcolor="#EFEFEF" width="23%" nowrap>email</td>
    <td bgcolor="#EFEFEF" width="10%" nowrap>年齡</td>
    <td bgcolor="#EFEFEF" width="8%" nowrap>縣市</td>
    <td bgcolor="#EFEFEF" width="12%" nowrap>家庭成員</td>
    <td bgcolor="#EFEFEF" width="14%" nowrap>收入</td>
    <td bgcolor="#EFEFEF" width="8%" nowrap>職業</td>
    <td bgcolor="#EFEFEF" width="8%" nowrap>學歷</td>
  </tr>
  
<%
	Dim City(25)
	City(1) = "臺北縣"
	City(2) = "宜蘭縣"
	City(3) = "桃園縣"
	City(4) = "新竹縣"
	City(5) = "苗栗縣"
	City(6) = "臺中縣"
	City(7) = "彰化縣"
	City(8) = "南投縣"
	City(9) = "雲林縣"
	City(10) = "嘉義縣"
	City(11) = "臺南縣"
	City(12) = "高雄縣"
	City(13) = "屏東縣"
	City(14) = "臺東縣"
	City(15) = "花蓮縣"
	City(16) = "澎湖縣"
	City(17) = "基隆市"
	City(18) = "新竹市"
	City(19) = "臺中市"
	City(20) = "嘉義市"
	City(21) = "臺南市"
	City(22) = "臺北市"
	City(23) = "高雄市"
	City(24) = "金門縣"
	City(25) = "連江縣"
	
	Dim Age(8)
	Age(1) = "0到11歲以下"
	Age(2) = "12到17歲"
	Age(3) = "18到23歲"
	Age(4) = "24到29歲"
	Age(5) = "30到39歲"
	Age(6) = "40到49歲"
	Age(7) = "50到59歲"
	Age(8) = "60歲以上"
	
	Dim Member(6)
	Member(1) = "1人"
	Member(2) = "2人"
	Member(3) = "3人"
	Member(4) = "4人"
	Member(5) = "5至8人"
	Member(6) = "8人以上"
	
	Dim Money(6)
	Money(1) = "5000元以下"
	Money(2) = "5000到1萬元"
	Money(3) = "1萬到3萬元"
	Money(4) = "3萬到5萬元"
	Money(5) = "5萬到10萬元"
	Money(6) = "10萬元以上"
	
	Dim Job(7)
	Job(1) = "軍"
	Job(2) = "公"
	Job(3) = "商"
	Job(4) = "教職"
	Job(5) = "自由業"
	Job(6) = "服務業"
	Job(7) = "其他"
	
	Dim edu(7)
	edu(1) = "國小"
	edu(2) = "國中"
	edu(3) = "高中"
	edu(4) = "大學"
	edu(5) = "碩士"
	edu(6) = "博士"
	edu(7) = "其他"
          
          
	i = 1
	while not list.EOF
		if i <= (page*pagecount) and i > (page-1)*pagecount then
		
			if i mod 2 = 1 then
				response.write "<tr>"
			else
				response.write "<tr bgcolor='#efefef'>"
			end if
			
			response.write "<td>" & list("m014_name") & "&nbsp;</td>"
			
			if list("m014_sex") = "M" then
				response.write "<td>男</td>"
			else
				response.write "<td>女</td>"
			end if
			
			response.write "<td><a href='mailto:" & trim(list("m014_email")) & "'>" & trim(list("m014_email")) & "</a>&nbsp;</td>"
			
			if int(list("m014_age")) > 0 then
				response.write "<td>" & age(int(list("m014_age"))) & "</td>"
			else
				response.write "<td>　</td>"
			end if
			
			if int(list("m014_addrarea")) > 0 then
				response.write "<td>" & city(int(list("m014_addrarea"))) & "</td>"
			else
				response.write "<td>　</td>"
			end if
			
			if int(list("m014_familymember")) > 0 then
				response.write "<td>" & member(int(list("m014_familymember"))) & "</td>"
			else
				response.write "<td>　</td>"
			end if
			
			if int(list("m014_money")) > 0 then
				response.write "<td>" & money(int(list("m014_money"))) & "</td>"
			else
				response.write "<td>　</td>"
			end if
			
			if trim(list("m014_job")) <> "" then
				response.write "<td>" & job(int(trim(list("m014_job")))) & "</td>"
			else
				response.write "<td>　</td>"
			end if
			
			if trim(list("m014_edu")) <> "" then
				response.write "<td>" & edu(int(trim(list("m014_edu")))) & "</td>"
			else
				response.write "<td>　</td>"
			end if
			
			response.write "</tr>"

		end if
		list.MoveNext
		i = i + 1
	wend
%>
  </form>
</table>
<p><a href="javascript:history.go(-1);">Previous Page</a></p>
</body>
</html>

<script language="JavaScript">
<!--
function Page_Onchange(form) {
	location.assign("02_namelist.asp?page=" + form.select.value + "&select2=" + form.select2.value + "&select3=" + form.select3.value + "&functype=<%= functype %>&subjectid=<%= subjectid %>")
}
 
function select2_Onchange(form, select3) {
	if ( form.options[form.selectedIndex].value == "sex" ) {
		select3.length = "3";
		select3.options[0].value = "";
		select3.options[0].text = "不拘";
		select3.options[1].value = "M";
		select3.options[1].text = "男";
		select3.options[2].value = "F";
		select3.options[2].text = "女";
		select3.selectedIndex = 0;
	} else if ( form.options[form.selectedIndex].value == "age" ) {
		select3.length = "9";
		select3.options[0].value = "";
		select3.options[0].text = "不拘";
		select3.options[1].value = "1";
		select3.options[1].text = "0到11歲以下";
		select3.options[2].value = "2";
		select3.options[2].text = "12到17歲";
		select3.options[3].value = "3";
		select3.options[3].text = "18到23歲";
		select3.options[4].value = "4";
		select3.options[4].text = "24到29歲";
		select3.options[5].value = "5";
		select3.options[5].text = "30到39歲";
		select3.options[6].value = "6";
		select3.options[6].text = "40到49歲";
		select3.options[7].value = "7";
		select3.options[7].text = "50到59歲";
		select3.options[8].value = "8";
		select3.options[8].text = "60歲以上";
		select3.selectedIndex = 0;
	} else if ( form.options[form.selectedIndex].value == "addrarea" ) {
		select3.length = "26";
		select3.options[0].value = "";
		select3.options[0].text = "不拘";
		select3.options[1].value = "1";
		select3.options[1].text = "台北縣";
		select3.options[2].value = "2";
		select3.options[2].text = "宜蘭縣";
		select3.options[3].value = "3";
		select3.options[3].text = "桃園縣";
		select3.options[4].value = "4";
		select3.options[4].text = "新竹縣";
		select3.options[5].value = "5";
		select3.options[5].text = "苗栗縣";                                                
		select3.options[6].value = "6";
		select3.options[6].text = "台中縣";
		select3.options[7].value = "7";
		select3.options[7].text = "彰化縣";
		select3.options[8].value = "8";
		select3.options[8].text = "南投縣";
		select3.options[9].value = "9";
		select3.options[9].text = "雲林縣";
		select3.options[10].value = "10";
		select3.options[10].text = "嘉義縣";
		select3.options[11].value = "11";
		select3.options[11].text = "台南縣";
		select3.options[12].value = "12";
		select3.options[12].text = "高雄縣";
		select3.options[13].value = "13";
		select3.options[13].text = "屏東縣";
		select3.options[14].value = "14";
		select3.options[14].text = "台東縣";
		select3.options[15].value = "15";
		select3.options[15].text = "花蓮縣";
		select3.options[16].value = "16";
		select3.options[16].text = "澎湖縣";
		select3.options[17].value = "17";
		select3.options[17].text = "基隆市";
		select3.options[18].value = "18";
		select3.options[18].text = "新竹市";                                
		select3.options[19].value = "19";
		select3.options[19].text = "台中市";
		select3.options[20].value = "20";
		select3.options[20].text = "嘉義市";
		select3.options[21].value = "21";
		select3.options[21].text = "台南市";                                
		select3.options[22].value = "22";
		select3.options[22].text = "台北市";
		select3.options[23].value = "23";
		select3.options[23].text = "高雄市";                                                                
		select3.options[24].value = "24";
		select3.options[24].text = "金門縣";
		select3.options[25].value = "25";
		select3.options[25].text = "連江縣";                
		select3.selectedIndex = 0;                
	} else if ( form.options[form.selectedIndex].value == "familymember" ) {
		select3.length = "7";
		select3.options[0].value = "";
		select3.options[0].text = "不拘";
		select3.options[1].value = "1";
		select3.options[1].text = "1人";
		select3.options[2].value = "2";
		select3.options[2].text = "2人";
		select3.options[3].value = "3";
		select3.options[3].text = "3人";
		select3.options[4].value = "4";
		select3.options[4].text = "4人";
		select3.options[5].value = "5";
		select3.options[5].text = "5至8人";
		select3.options[6].value = "6";
		select3.options[6].text = "8人以上";
		select3.selectedIndex = 0;                
	} else if ( form.options[form.selectedIndex].value == "money" ) {
		select3.length = "7";
		select3.options[0].value = "";
		select3.options[0].text = "不拘";
		select3.options[1].value = "1";
		select3.options[1].text = "5000元以下";
		select3.options[2].value = "2";
		select3.options[2].text = "5000到1萬元";
		select3.options[3].value = "3";
		select3.options[3].text = "1萬到3萬元";
		select3.options[4].value = "4";
		select3.options[4].text = "3萬到5萬元";
		select3.options[5].value = "5";
		select3.options[5].text = "5萬到10萬元";
		select3.options[6].value = "6";
		select3.options[6].text = "10萬元以上";
		select3.selectedIndex = 0;
	} else if ( form.options[form.selectedIndex].value == "job" ) {
		select3.length = "8";
		select3.options[0].value = "";
		select3.options[0].text = "不拘";
		select3.options[1].value = "1";
		select3.options[1].text = "軍";
		select3.options[2].value = "2";
		select3.options[2].text = "公";
		select3.options[3].value = "3";
		select3.options[3].text = "商";
		select3.options[4].value = "4";
		select3.options[4].text = "教職";
		select3.options[5].value = "5";
		select3.options[5].text = "自由業";
		select3.options[6].value = "6";
		select3.options[6].text = "服務業";
		select3.options[7].value = "7";
		select3.options[7].text = "其他";
		select3.selectedIndex = 0;
	} else if ( form.options[form.selectedIndex].value == "edu" ) {
		select3.length = "8";
		select3.options[0].value = "";
		select3.options[0].text = "不拘";
		select3.options[1].value = "1";
		select3.options[1].text = "國小";
		select3.options[2].value = "2";
		select3.options[2].text = "國中";
		select3.options[3].value = "3";
		select3.options[3].text = "高中";
		select3.options[4].value = "4";
		select3.options[4].text = "大學";
		select3.options[5].value = "5";
		select3.options[5].text = "碩士";
		select3.options[6].value = "6";
		select3.options[6].text = "博士";
		select3.options[7].value = "7";
		select3.options[7].text = "其他";
		select3.selectedIndex = 0;
	}
}
-->
</script>   
 