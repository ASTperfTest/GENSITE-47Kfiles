<%@ CodePage = 65001 %>
<%
	HTProgCode = "GW1_vote01"
%>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	subjectid = request("subjectid")
	sql = "select m011_subject, m011_questno, m011_bdate, m011_edate from m011 where m011_subjectid = " & replace(subjectid, "'", "''")
	set rs = conn.execute(sql)
	if rs.eof then	response.end
	
	sex = array("", "男", "女")
	sex_type = array("", "M", "F")
	City = array("", "臺北縣", "宜蘭縣", "桃園縣", "新竹縣", "苗栗縣", "臺中縣", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "臺南縣", "高雄縣", "屏東縣", "臺東縣", "花蓮縣", "澎湖縣", "基隆市", "新竹市", "臺中市", "嘉義市", "臺南市", "臺北市", "高雄市", "金門縣", "連江縣")
	Age = array("", "0到11歲以下", "12到17歲", "18到23歲", "24到29歲", "30到39歲", "40到49歲", "50到59歲", "60歲以上")
	Member = array("", "1人", "2人", "3人", "4人", "5至8人", "8人以上")
	Money = array("", "5000元以下", "5000到1萬元", "1萬到3萬元", "3萬到5萬元", "5萬到10萬元", "10萬元以上")
	Job = array("", "軍", "公", "商", "教職", "自由業", "服務業", "其他")
	edu = array("", "國小", "國中", "高中", "大學", "碩士", "博士", "其他")
	
	response.ContentType = "application/csv"
	response.AddHeader "Content-Disposition", "filename=export.csv;"
	
	response.write """調查對象分析""" & vbcrlf
	response.write """調查主題：" & trim(rs("m011_subject")) & """" & vbcrlf
	response.write """起訖時間：" & datevalue(rs("m011_bdate")) & " ~ " & datevalue(rs("m011_edate")) & """" & vbcrlf

	response.write """姓名"", ""性別"", ""email"", ""年齡"", ""居住縣市"", ""家庭成員人數"", ""收入"", ""職業""" & vbcrlf

	sql = "select m014_name, m014_sex, m014_email, m014_age, m014_addrarea, m014_familymember, isNull(m014_money, 0) m014_money, m014_job, isnull(m014_edu, 0) m014_edu from m014 where m014_subjectid = " & replace(subjectid, "'", "''") & " order by m014_id"
	set rs = conn.execute(sql)
	while not rs.eof
		response.write """" & trim(rs("m014_name")) & """, """ & replace(replace(trim(rs("m014_sex")), "M", "男"), "F", "女") & """, """ & trim(rs("m014_email")) & """, """ & age(int(rs("m014_age"))) & """, """ & city(int(rs("m014_addrarea"))) & """, """ & member(int(rs("m014_familymember"))) & """, """ & money(int(rs("m014_money"))) & """, """ & job(int(rs("m014_job"))) & """, """ & edu(int("0" & rs("m014_edu"))) & """" & vbcrlf
		rs.movenext
	wend
%>
