<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="newmember"
HTProgPrefix = "newMember"
%>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include file="newMemberMail.inc" -->
<!--#include virtual = "/inc/checkGIPAPconfig.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../css/list.css" rel="stylesheet" type="text/css">
<link href="../css/layout.css" rel="stylesheet" type="text/css">
<link href="../css/form.css" rel="stylesheet" type="text/css" /></head>
<title>資料管理／資料上稿</title>
</head>

<%

'審核狀態
dim validatetype 
dim roletype
otherCondition = request("otherCondition")
role = request("role")
if Request.QueryString = "" then
	session("validate")=""
else
	session("validate")=request("validate")
end if
validate = session("validate")
'validate=request("validate")
validatetype = validate
roletype = role

'---整批通過或不通過-
if request("submitTask") = "AllPass" then
	
	Dim selectedItems : selectedItems = request("selectedItems")
	Dim items : items = split(selectedItems, ";")
	
	for each item in items
		if item <> "" then
			checksql="SELECT id_type1, id_type2, email, realname FROM Member WHERE account = '" & item & "'"
			Set RSreg = conn.execute(checksql)
			if not RSreg.eof then
				id_type1 = RSreg("id_type1")
				id_type2 = RSreg("id_type2")
				email = RSreg("email")
				realname = RSreg("realname")
				if id_type1 = "1" and id_type2 = "1" then
					sql = "UPDATE Member SET scholarValidate = 'Y' WHERE account = '" & item & "'"
					conn.execute(sql)
				end if				
				if email <> "" then
					ePaperTitle = "農業知識入口網學者會員審核狀況通知"
					Body = "親愛的 " & realname & " 您好:<br/><br/>" & vbCrLf & vbCrLf
					Body = Body & "感謝您申請「農業知識入口網」的學者會員。<br/>" & vbCrLf
					Body = Body & "「農業知識入口網」""通過""您的學者會員申請。<br/>"& vbCrLf
					Body = Body & "有了您的參與將使得這個農業知識平台更理想更茁壯。我們熱切期盼您提供寶貴的意見與建議。<br/>" & vbCrLf
					Body = Body & "謝謝!<br/>" & vbCrLf
					Body = Body & "                                      敬祝平安<br/>" & vbCrLf
					Body = Body & "                                                系統管理員 敬上<br/>" & vbCrLf

					'S_Email = "taft_km@mail.coa.gov.tw"
					S_Email = """"&getGIPApconfigText("EmailFromName")&""" <"&getGIPApconfigText("EmailFrom")&">"
					R_Email = email
			
					Call Send_Email(S_Email,R_Email,ePaperTitle,Body)
				end if
			end if
		end if
	next
	
elseif request("submitTask") = "AllNoPass" then
	
	Dim selectedItem : selectedItem = request("selectedItems")
	Dim Item : Items = split(selectedItem, ";")
	for each item in Items
		if Item <> "" then 
			checksql = "SELECT id_type1, id_type2, email, realname FROM Member WHERE account = '" & item & "'"
			Set RSreg = conn.execute(checksql)
			if not RSreg.eof then
				id_type1 = RSreg("id_type1")
				id_type2 = RSreg("id_type2")
				email = RSreg("email")
				realname = RSreg("realname")
				if id_type1 = "1" and id_type2 = "1" then
					sql = "UPDATE Member SET scholarValidate = 'N' WHERE account = '" & item & "'"
					conn.execute(sql)
				end if
				if email <> "" then
					ePaperTitle = "農業知識入口網學者會員審核狀況通知"
					Body = "親愛的 " & realname & " 您好:<br/><br/>" & vbCrLf & vbCrLf
					Body = Body & "感謝您申請「農業知識入口網」的學者會員。<br/>" & vbCrLf
					Body = Body & "「農業知識入口網」""不通過""您的學者會員申請。<br/>"& vbCrLf
					Body = Body & "有了您的參與將使得這個農業知識平台更理想更茁壯。我們熱切期盼您提供寶貴的意見與建議。<br/>" & vbCrLf
					Body = Body & "謝謝!<br/>" & vbCrLf
					Body = Body & "                                      敬祝平安<br/>" & vbCrLf
					Body = Body & "                                                系統管理員 敬上<br/>" & vbCrLf

					'S_Email = "taft_km@mail.coa.gov.tw"
					S_Email = """"&getGIPApconfigText("EmailFromName")&""" <"&getGIPApconfigText("EmailFrom")&">"
					R_Email = email
			
					Call Send_Email(S_Email,R_Email,ePaperTitle,Body)
				end if
			end if
		end if
	next
end if

cSql = " SELECT COUNT(*) FROM Member" 
cSql = cSql & vbcrlf & " where 1=1 "

fSql = " member.account"
fSql = fSql & vbcrlf & " , realname"
fSql = fSql & vbcrlf & " , email"
fSql = fSql & vbcrlf & " , createtime"
fSql = fSql & vbcrlf & " , id_type1"
fSql = fSql & vbcrlf & " , id_type2"
fSql = fSql & vbcrlf & " , id_type3"
fSql = fSql & vbcrlf & " , scholarValidate"
fSql = fSql & vbcrlf & " , status"
fSql = fSql & vbcrlf & " , mcode"
fSql = fSql & vbcrlf & " , ValidCount"
fSql = fSql & vbcrlf & " , InviteFriends_Head.account as Inviter "
fSql = fSql & vbcrlf & " FROM member"
fSql = fSql & vbcrlf & " left join InviteFriends_Detail on InviteFriends_Detail.inviteAccount = Member.account"
fSql = fSql & vbcrlf & " left join InviteFriends_Head on InviteFriends_Detail.InvitationCode = InviteFriends_Head.InvitationCode"
fSql = fSql & vbcrlf & " where 1=1 "

select case request("otherCondition")
    case "invitee"
        cSql =csql & " and Member.account in (select inviteAccount from InviteFriends_Detail)" 
        fSql =fsql & " and Member.account in (select inviteAccount from InviteFriends_Detail)"         
end select


if request("keep")="Y" then

	if validatetype = "EY" Then
		cSql =csql & " and mcode='Y'" 
		fSql =fSql & " and mcode='Y'" 
   
	ElseIf validatetype = "ES" Then
		cSql =csql & " AND ((mcode <> 'Y' or mcode is null) and ValidCount > 0) " 
		fSql =fSql & " AND ((mcode <> 'Y' or mcode is null) and ValidCount > 0) " 

	ElseIf validatetype = "EA" Then
		cSql =csql & " AND ((mcode <> 'Y' or mcode is null) and ValidCount = 0) " 
		fSql =fSql & " AND ((mcode <> 'Y' or mcode is null) and ValidCount = 0) " 
   
	ElseIf validatetype <>"" and validatetype <> "all" then
       cSql =csql & " and scholarValidate = " &"'"& validatetype & "'" 
	   fSql =fSql & " and scholarValidate = " &"'"& validatetype & "'" 
   
	End If
  
   if roletype <> "" and roletype <> "all" then
        cSql =csql & " and id_type" & roletype &  " ='1'"  
		fSql =fSql & " and id_type" & roletype &  " ='1'"   
   end if
   
end if 
  nowPage =1
  if Request.QueryString("nowPage") <> "" then
	nowPage = Request.QueryString("nowPage")  '現在頁數
  end if
  nowpagesize = 15
  if Request.QueryString("pagesize") <> "" then 
	nowpagesize = Request.QueryString("pagesize")
  end if
  PerPageSize = cint(nowpagesize)
  if PerPageSize <= 0 then  PerPageSize = 15 

 set RSc = conn.execute(cSql)
  totRec = RSc(0)       '總筆數
  totPage = int(totRec / PerPageSize + 0.999)

  if cint(nowPage) < 1 then 
    nowPage = 1
  elseif cint(nowPage) > totPage then 
    nowPage = totPage 
  end if 
  
            	
    fSql = "SELECT TOP " & nowPage * PerPageSize & fSql & " ORDER BY createtime DESC"


	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
'	RSreg.CacheSize = PerPageSize

'----------HyWeb GIP DB CONNECTION PATCH----------
'	RSreg.Open fSql,Conn, 3, 1
Set RSreg = Conn.execute(fSql)
	RSreg.CacheSize = PerPageSize
'----------HyWeb GIP DB CONNECTION PATCH----------

'response.write fSql
	if Not RSreg.eof then
		if totRec > 0 then 
      RSreg.PageSize = PerPageSize       '每頁筆數
      RSreg.AbsolutePage = nowPage      
		end if    
	end if 


 	
	   
%>
<body>
<div id="FuncName">
	<h1>會員整合管理</h1>
	<div id="Nav">
			<A href="newMember_Query.asp" title="指定查詢條件">會員查詢</A>
			<A href="newMember_Add.asp" title="新增資料">新增會員</A>	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">
	會員資料維護&nbsp;
	<font size="2">【主題單元:新增專家資料 / 單元資料:純網頁】</font></div>
<!--  條列頁簡易查詢功能  -->


<Form id="Form2" name=reg method="POST" >
  <INPUT TYPE=hidden name=submitTask value="">
  <INPUT TYPE=hidden name=parentIcuitem value="">
  <INPUT TYPE=hidden name=selectedItems value="">
	<!-- 分頁 -->
<div class="browseby">
條件篩選依： 審核狀態
  <select name="validate" id="validate" >
  		
    <option value="" <%if validate="" then%>selected<%END IF%>>請選擇</option>
    <option value="all" <%if validate="all" then%>selected<%END IF%>>全部</option>
    <option value="W" <%if validate="W" then%>selected<%END IF%>>待審核</option>
    <option value="Y" <%if validate="Y" then%>selected<%END IF%>>審核通過</option>
    <option value="N" <%if validate="N" then%>selected<%END IF%>>審核不通過</option>
    <option value="Z" <%if validate="Z" then%>selected<%END IF%>>無(未申請學者)</option>
    <option value="EY" <%if validate="EY" then%>selected<%END IF%>>Email 認證(已通過)</option>
    <option value="ES" <%if validate="ES" then%>selected<%END IF%>>Email 認證(未通過-已申請)</option>
    <option value="EA" <%if validate="EA" then%>selected<%END IF%>>Email 認證(未通過-未申請)</option>
  </select>
｜ 會員身分
<select name="role" id="role">
  <option value="" <% if role="" then %>selected<%end if%>>請選擇</option>
  <option value="all" <%if role="all" then%>selected<%END IF%>>全部</option>
  <option value="1" <%if role="1" then%>selected<%END IF%>>一般會員</option>
  <option value="2" <%if role="2" then%>selected<%END IF%>>學者會員</option>
  <option value="3" <%if role="3" then%>selected<%END IF%>>專家</option>
</select>
｜ 其他
<select name="otherCondition" id="id_otherCondition" onchange="onchange_otherCondition()">
  <option value="" <% if otherCondition="" then %>selected<%end if%>>請選擇</option>
  <option value="invitee" <%if otherCondition="invitee" then%>selected<%END IF%>>經由邀請而註冊成功</option>  
</select>
</div>
<div id="Page">
  共<em><%=totRec%></em>筆資料，目前在第<em><%=nowPage%> /<%=totPage%></em>	頁	
  
       
	
  
  每頁顯示
  
   
<select id=PerPage size="1" style="color:#FF0000" class="select">            
     <option value="15"<%if PerPageSize=15 then%> selected<%end if%>>15</option>                       
     <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
     <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
</select> 
筆


<% if cint(nowPage) <> 1 then %>
			<img src="/images/arrow_previous.gif" alt="上一頁">       		
			<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>&validate=<%=validate%>&role=<%=role%>">上一頁</a> ｜
    <% end if %>跳至第
    <select id=GoPage size="1" style="color:#FF0000" class="select">
		<% For iPage=1 to totPage %> 
			<option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
    <% Next %>   
    </select>      
		頁	
	<% if cint(nowPage)<>totPage then %> 
     ｜<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>&validate=<%=validate%>&role=<%=role%>"">下一頁
      <img src="/images/arrow_next.gif" alt="下一頁"></a> 
    <% end if %>  
		
		
		


  </div>
	<!-- 列表 -->
	<table cellspacing="0" id="ListTable">
		<tr>

		<th class="First" scope="col">&nbsp;</th>
					<th scope="col">會員帳號</th>
					<th scope="col">真實姓名</th>
					<th scope="col">電子郵件</th>
					<th scope="col">申請日期</th>
					<th scope="col">身分</th>
					<th scope="col">學者審核</th>
					<th scope="col">專家</th>
					<th scope="col">狀態</th>
					<!--<th class="First" scope="col">預覽</th>   -->
	  </tr>                  
<tr>
     
    
   
	<%
	while not RSreg.eof
	%>
	<td align="center"><span class="eTableContent">
    <input type="checkbox" name="checkbox"  id="<%=trim(RSreg("account"))%>" value="checkbox">
    </span></td>  
	<td align="center"><a href="newMemberEdit.asp?account=<%=trim(RSreg("account"))%>&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>&from=list"><%=CheckAndReplaceID(RSreg("account"))%></a></td>
	<td align="center"><%=CheckAndReplaceID(RSreg("realname"))%>&nbsp;</td>
	
            
    <% If RSreg("mcode") = "Y" Then %>
    <td align="center"><%=ReplaceEmailId(RSreg("email"))%>&nbsp;</td>
    <% Else %>
    <td align="center"><font color='grey'><%=ReplaceEmailId(RSreg("email"))%>&nbsp;(<%=RSreg("ValidCount")%>)</font></td>
    <% End If %>            
	
	<td align="center"><%=FormatDateTime(RSreg("createtime"), 2) + " " + FormatDateTime(RSreg("createtime"), 4)%>&nbsp;
	    <%
	        if not isnull(RSreg("inviter")) then
	            response.Write "(" & RSreg("inviter") & ")"
	        end if
	    %>
	</td>
	<td align="center">
	<% if RSreg("id_type1")="1" then %> 
	一般會員
	<% end if%>
	<%if RSreg("id_type2")="1" then %>	
	, 學者會員
	<% end if%>
	&nbsp;
	</td>
	<td>
	<%if  RSreg("scholarValidate")="W" then%>
	<FONT COLOR=#FF0000>待審核</FONT>
	<% end if%>
	<%if RSreg("scholarValidate")="N" then%>
	不通過
	<% end if%>
	<%if RSreg("scholarValidate")="Y" then%>
	通過
	<% end if%>
	<%if RSreg("scholarValidate")="Z" then%>
	無
	<% end if%>
	&nbsp;
	</td>
	<td >
	<% if RSreg("id_type3")="1" then %> 
	V
	<% end if%>
	&nbsp;
	</td>
  <td>
	<% if RSreg("status")="Y" then%>
	正常
	<% end if%>
	<% if RSreg("status")="N" then%>
	停權
	<% end if%>
	&nbsp;
	</td>
</tr>
 <%	
			RSreg.MoveNext
		wend
	%>                    
           
</TABLE>
           <div align="center">
             <input type="button" name="button" value="整批通過" onClick="AllPassValidate()"  class="cbutton" />
             <input type="button" name="button" value="整批不通過" onClick="AllNoPassValidate()" class="cbutton" />   
           </div>
	
       <!-- 程式結束 ---------------------------------->  
</form>  




</body>
</html> 


<script type="text/javascript">
    function onchange_otherCondition() {
        urlStr = "<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>";
        urlStr += "&role=" + reg.role.value + "&validate=" + reg.validate.value + "&otherCondition=" + reg.otherCondition.value;        
        document.location.href = urlStr;
    }    
</script>

<script language="javascript">
	function AllPassValidate() 
	{
	    var obj=document.getElementsByName("checkbox");
        var len = obj.length;
        var checked = false;
        var CheckList="";
        for (i = 0; i < len; i++)
        {		
            if (obj[i].checked == true)
            {
			  CheckList+=obj[i].id + ";" ;
            }
        } 
	    document.getElementById("selectedItems").value = CheckList;	
		document.getElementById("submitTask").value = "AllPass";
		document.getElementById("reg").submit();
	}
	function AllNoPassValidate() 
	{
	    var obj=document.getElementsByName("checkbox");
        var len = obj.length;
        var checked = false;
        var CheckList="";
        for (i = 0; i < len; i++)
        {		
            if (obj[i].checked == true)
            {
			  CheckList+=obj[i].id + ";" ;
            }
        } 
	    document.getElementById("selectedItems").value = CheckList;	
		document.getElementById("submitTask").value = "AllNoPass";
		document.getElementById("reg").submit();
	}
	
</script>       
<script Language=VBScript>
	
		sub formModSubmit()
			reg.submitTask.value = "DELETE"
			reg.action = "<%=HTProgPrefix%>List.asp?icuitem=<%=request.querystring("icuitem")%>"
			reg.Submit
		end sub
		
		dim gpKey
    sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=" & newPage & "&pagesize=<%=PerPageSize%>"_
		   & "&validate=" & reg.validate.value & "&role=" & reg.role.value + "&otherCondition=" + reg.otherCondition.value
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>" & "&pagesize=" & newPerPage _
             & "&validate=" & reg.validate.value & "&role=" & reg.role.value  + "&otherCondition=" + reg.otherCondition.value
     end sub 

     sub shortLongList(sorl)
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>" _
			& "&shortLongList=" & sorl
     end sub   

    sub role_OnChange
	
    //if reg.role.value = "1" then
		//reg.validate.value = "Z"
	//end if

	 document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>" _
			& "&role=" & reg.role.value & "&validate=" & reg.validate.value  + "&otherCondition=" + reg.otherCondition.value
     end sub	 
	 
	 sub validate_OnChange
		document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>" _
			& "&validate=" & reg.validate.value & "&role=" & reg.role.value + "&otherCondition=" + reg.otherCondition.value
	 end sub
	</script>                                       
