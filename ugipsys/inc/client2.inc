<%
Set Conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeout = 20
conn.CommandTimeout = 30
conn.CursorLocation = 3

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

xRootID=session("xRootID")

'Button的ID & Value
Public Const cBtnCreateId="btnCreate"
Public Const cBtnCreateName="新增"
Public Const cBtnReadId="btnRead"
Public Const cBtnReadName="查詢"
Public Const cBtnUpdateId="btnUpdate"
Public Const cBtnUpdateName="儲存"
Public Const cBtnDeleteId="btnDelete"
Public Const cBtnDeleteName="刪除"
'回上頁
Public Const cBtnBackId="btnBack"
Public Const cBtnBackName="回上頁"

Public const cSaveOk= "資料已儲存成功!!"
Public const cDelOk= "資料已刪除成功!!"
Public const cKeydup="輸入之資料已存在!!"
Public const cNoData="很抱歉,目前查無相關資料!!"
'權限變數名稱(新增/查詢/修改/刪除)
dim pubAuthorityCreate
dim pubAuthorityRead
dim pubAuthorityUpdate
dim pubAuthorityDelete
'pubAuthorityCreate/pubAuthorityRead/pubAuthorityUpdate/pubAuthorityDelete
'新增 GetButton(cBtnCreateId,cBtnCreateName,pubAuthorityCreate,"")
'查詢 GetButton(cBtnReadId,cBtnReadName,pubAuthorityRead,"")
'儲存 GetButton(cBtnUpdateId,cBtnUpdateName,pubAuthorityUpdate,"")
'刪除 GetButton(cBtnDeleteId,cBtnDeleteName,pubAuthorityDelete,"")
'回上頁 GetButton(cBtnBackId,cBtnBackName,"1",Get_UrlString("key1"))
'Get Button,
'傳入1.Button ID 2.Button Name 3.權限代碼(mid(HexToBin(xx),?,1)
' x111 依序表示 (新增/查詢/修改/刪除) 的權限
Public Sub GetButton(ByVal ButtonId,ByVal ButtonName,ByVal Authority,ByVal Url_String)
	if Authority="1" then
		if trim(Url_String) & "" <> "" then
			response.write "<INPUT type=button value='" & ButtonName & "' id='" & ButtonId & "' name='" & ButtonId & "' LANGUAGE=javascript onclick=""return " & ButtonId & "_onclick('" & Url_String & "')"">" & vbCrLf
		else
			response.write "<INPUT type=button value='" & ButtonName & "' id='" & ButtonId & "' name='" & ButtonId & "' LANGUAGE=javascript onclick=""return " & ButtonId & "_onclick()"">" & vbCrLf
		end if
	end if
end Sub

'抓出連結字串
Public function Get_UrlString(ByVal str_search)
	dim pri_i,pri_j,pri_str_search_length
	'找key1= or 找page=
	pri_i=instr(1,Request.ServerVariables("QUERY_STRING"),str_search)
	pri_str_search_length=len(Request.QueryString(str_search))
	if pri_i = 0 then
		Get_UrlString=Request.ServerVariables("QUERY_STRING")
	else
		'從找到str_search的位置開始,再加上page=的長度,及值的長度再加個 & 
		pri_j=pri_i+len(str_search)+len("=") +pri_str_search_length+1
		Get_UrlString=mid(Request.ServerVariables("QUERY_STRING"),pri_j)
	end if
end Function

'題目題型(靜態下拉式選單)
'A=便民服務
'B=一般進口海運
'C=一般進口空運
'D=一般出口海運
'E=一般出口空運
'F=外銷沖退稅‧物流中心
'G=保稅倉庫‧科學工業園區
'H=保稅工廠‧加工出口區
'I=旅客通關
Public Function GetPollType(ByVal id_name,ByVal id_value)
  with response
	.Write vbCrLf
	.Write "<select name=" & id_name & " id=" & id_name & " >" & vbCrLf
	.Write "<option " & IsSelect(id_value,"") & " value=''>　</option>" & vbCrLf
	.Write "<option " & IsSelect(id_value,"A") & " value='A'>便民服務</option>" & vbCrLf
	.Write "<option " & IsSelect(id_value,"B") & " value='B'>一般進口海運</option>" & vbCrLf
	.Write "<option " & IsSelect(id_value,"C") & " value='C'>一般進口空運</option>" & vbCrLf
	.Write "<option " & IsSelect(id_value,"D") & " value='D'>一般出口海運</option>" & vbCrLf
	.Write "<option " & IsSelect(id_value,"E") & " value='E'>一般出口空運</option>" & vbCrLf
	.Write "<option " & IsSelect(id_value,"F") & " value='F'>外銷沖退稅‧物流中心</option>" & vbCrLf
	.Write "<option " & IsSelect(id_value,"G") & " value='G'>保稅倉庫‧科學工業園區</option>" & vbCrLf
	.Write "<option " & IsSelect(id_value,"H") & " value='H'>保稅工廠‧加工出口區</option>" & vbCrLf
	.Write "<option " & IsSelect(id_value,"I") & " value='I'>旅客通關</option>" & vbCrLf
	.Write "</select>"	
  end with
end Function

'計劃類型
'將計劃類型轉為名稱
Public Function ConvPollType(ByVal id)
select case ucase(trim(id))
	case "A" 
		ConvPollType="便民服務"
	case "B"
		ConvPollType="一般進口海運"
	case "C"
		ConvPollType="一般進口空運"
	case "D"
		ConvPollType="一般出口海運"
	case "E"
		ConvPollType="一般出口空運"
	case "F"
		ConvPollType="外銷沖退稅‧物流中心"
	case "G"
		ConvPollType="保稅倉庫‧科學工業園區"
	case "H"
		ConvPollType="保稅工廠‧加工出口區"
	case "I"
		ConvPollType="旅客通關"
	case else
		ConvPollType="&nbsp;"
end select
end Function

'檢查靜態的 Listbox 是否有選擇, 傳入
'1.s1=所選擇的Listbox 的"值"
'2.s2=option的值
public function IsSelect(ByVal s1,ByVal s2)
	dim pri_i
	'if ucase(s3)="SELECTED" then
	'	IsSelect="selected"
	'else
		if ucase(s1)=ucase(s2) then
			IsSelect="selected"
		else
			IsSelect=""
		end if
	'end if
end function


%>
