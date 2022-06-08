<%@  codepage="65001" %>
<% 
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTUploadPath=session("Public")+"data/"
HTProgPrefix = "newMemberQuery_Act"
Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="newMember"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="../css/list.css" rel="stylesheet" type="text/css">
    <link href="../css/layout.css" rel="stylesheet" type="text/css">
    <link href="../css/form.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>
<title>資料管理／資料上稿</title></head> <%

'---計算筆數-

cSql = "SELECT COUNT(*) FROM Member where 1=1 " 
fsql="* FROM Member where 1=1 "  

		'---更新圖片狀態---
if request("submitTask") = "AllPass" then
	
	Dim selectedItems : selectedItems = request("selectedItems")
	Dim items : items = split(selectedItems, ";")
	for each item in items
		if item <> "" then 
		checksql="SELECT id_type1,id_type2 FROM  Member WHERE  account = "&"'"& item &"'"
		Set RSreg=conn.execute(checksql)
		if RSreg("id_type1")="1" and RSreg("id_type2")="1" then
			sql = "UPDATE Member SET scholarValidate ='Y' WHERE account="&"'"& item &"'"
			conn.execute(sql)
		end if
		end if
	next
elseif request("submitTask") = "AllNoPass" then
	
	Dim selectedItem : selectedItem = request("selectedItems")
	Dim Item : Items = split(selectedItem, ";")
	for each item in Items
		if Item <> "" then 
		checksql="SELECT id_type1,id_type2 FROM  Member WHERE  account = "&"'"& item &"'"
		Set RSreg=conn.execute(checksql)
		if RSreg("id_type1")="1" and RSreg("id_type2")="1" then
			sql = "UPDATE Member SET scholarValidate ='N' WHERE account="&"'"& item &"'"
			conn.execute(sql)
		end if
		end if
	next
end if

'若點選"會員整合查詢"才要request
if request.QueryString("request") = "yes" then
  session("account") = request("account") 
  session("realname") = request("realname")
  session("nickname") = request("nickname") 
  session("keyword") = request("keyword") 
  session("id_type1") = request("id_type1")
  session("id_type2") = request("id_type2")
  session("scholarValidate") = request("scholarValidate")
  session("id_type3") = request("id_type3")
  session("status") = request("status")
  session("KMcat") = request("KMcat")
  session("remark") = request("remark") 
  session("emailaddress") = request("emailaddress")
  session("homeaddr") = request("homeaddr")
  session("idcard") = request("idcard")
  session("mcode") = request("mcode")
  session("kmintra") = request("kmintra") 
end if

'編修存檔後回原查詢結果頁，故改用session存原查詢的條件  Grace 
'---加入搜尋條件-
 dim account:account=session("account") 
 dim realname:realname=trim(session("realname")) 
 dim nickname:nickname=session("nickname") 
 dim keyword:keyword=session("keyword") 
 dim id_type1:id_type1=session("id_type1")
 dim id_type2:id_type2=session("id_type2")
 dim scholarValidate:scholarValidate=session("scholarValidate")
 dim id_type3:id_type3=session("id_type3")
 dim status:status=session("status")
 dim KMcat:KMcat=session("KMcat")
 dim remark:remark=session("remark")
 dim emailaddress:emailaddress=session("emailaddress")
 dim homeaddr:homeaddr=session("homeaddr")
 dim idcard:idcard=session("idcard")
 dim mcode:mcode=session("mcode")
 dim kmintra:kmintra=session("kmintra") 
 

  if mcode <> ""  then     
        if mcode = "Y"  then
            fSql = fSql & " AND mcode = 'Y'" 
            cSql = cSql & " AND mcode = 'Y'" 
        end if
        if mcode = "Sended"  then
            fSql = fSql & " AND ((mcode <> 'Y' or mcode is null) and ValidCount > 0)" 
            cSql = cSql & " AND ((mcode <> 'Y' or mcode is null) and ValidCount > 0)" 
        end if
        if mcode = "Apply"  then
            fSql = fSql & " AND ((mcode <> 'Y' or mcode is null) and ValidCount = 0)" 
            cSql = cSql & " AND ((mcode <> 'Y' or mcode is null) and ValidCount = 0)" 
        end if        
  end if
  
  if account <> ""  then 
  fSql = fSql & " AND account LIKE '%" & account  & "%'" 
  cSql = cSql & " AND account LIKE '%" & account  & "%' " 
  end if
  
  'sam 移除Unicode查詢 存Unicode用英文查會查到不相干的資料
  if trim(realname) <> "" then '加入Unicode查詢
  fSql = fSql & " AND realname LIKE '%" & trim(realname) & "%' " 
  cSql = cSql & " AND realname LIKE '%" & trim(realname) & "%' " 
  end if
  if nickname <> "" then '加入Unicode查詢
  fSql = fSql & " AND nickname LIKE '%" & nickname & "%' " 
  cSql = cSql & " AND nickname LIKE '%" & nickname & "%' " 
  end if
  if keyword <> "" then 
  fSql = fSql & " AND keyword LIKE '%" & keyword & "%' " 
  cSql = cSql & " AND keyword LIKE '%" & keyword & "%' " 
  end if
  if id_type1="1" then  
  fSql = fSql & " AND id_type1=1" 
  cSql = cSql & " AND id_type1=1" 
  end if
  if id_type2="1" then  
  fSql = fSql & " AND id_type2=1" 
  cSql = cSql & " AND id_type2=1" 
  end if
  if scholarValidate="W" then  
  fSql = fSql & " AND scholarValidate='W'" 
  cSql = cSql & " AND scholarValidate='W'" 
  end if
  if scholarValidate="Y" then  
  fSql = fSql & " AND scholarValidate='Y'" 
  cSql = cSql & " AND scholarValidate='Y'" 
  end if
  if request("scholarValidate")="N" then  
  fSql = fSql & " AND scholarValidate='N'" 
  cSql = cSql & " AND scholarValidate='N'" 
  end if
  if scholarValidate="Z" then  
  fSql = fSql & " AND scholarValidate='Z'" 
  cSql = cSql & " AND scholarValidate='Z'" 
  end if
  if id_type3="1" then  
    'Grace
	fSql = fSql & " AND id_type3=1 " 
    cSql = cSql & " AND id_type3=1 " 
    if kmintra<>"1" then 
      fSql = fSql & " AND member_org <> 'intra'" 
      cSql = cSql & " AND member_org <> 'intra'" 
	end if
  end if
  if id_type3="0" then  
  fSql = fSql & " AND (id_type3=0 or id_type3 is null) " 
  cSql = cSql & " AND (id_type3=0 or id_type3 is null) " 
  end if
  if status="Y" then  
  fSql = fSql & " AND status='Y'" 
  cSql = cSql & " AND status='Y'" 
  end if
  if status="N" then  
  fSql = fSql & " AND status='N'" 
  cSql = cSql & " AND status='N'"
  end if
  if  KMcat<>"" then  
  fSql = fSql & " AND KMcat LIKE '%" & KMcat & "%' " 
  cSql = cSql & " AND KMcat LIKE '%" & KMcat & "%' " 
  end if
  if  remark<>"" then  
  fSql = fSql & " AND remark LIKE '%" & remark & "%' " 
  cSql = cSql & " AND remark LIKE '%" & remark & "%' " 
  end if
  if emailaddress <> ""  then 
  fSql = fSql & " AND email LIKE '%" & emailaddress  & "%'" 
  cSql = cSql & " AND email LIKE '%" & emailaddress  & "%' " 
  end if
  if homeaddr <> ""  then 
  fSql = fSql & " AND homeaddr LIKE '%" & homeaddr  & "%'" 
  cSql = cSql & " AND homeaddr LIKE '%" & homeaddr  & "%' " 
  end if
  if idcard <> "" then
  fSql = fSql & " AND [id] LIKE '%" & idcard  & "%'" 
  cSql = cSql & " AND [id] LIKE '%" & idcard  & "%' " 
  end if


  nowPage = Request.QueryString("nowPage")  '現在頁數
  
  if request("newPerPage") <> "" then
    PerPageSize = cint(request("newPerPage"))
  else
  PerPageSize = cint(Request.QueryString("pagesize"))
  end if
  
  
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


	if Not RSreg.eof then
		if totRec > 0 then 
      RSreg.PageSize = PerPageSize       '每頁筆數
      RSreg.AbsolutePage = nowPage      
		end if    
	end if 
	
Function Chg_UNI(str)        'ASCII轉Unicode
	dim old,new_w,iStr
	old = str
	new_w = ""
	for iStr = 1 to len(str)
		if ascw(mid(old,iStr,1)) < 0 then
			new_w = new_w & "&#" & ascw(mid(old,iStr,1))+65536 & ";"
		elseif        ascw(mid(old,iStr,1))>0 and ascw(mid(old,iStr,1))<127 then
			new_w = new_w & mid(old,iStr,1)
		else
			new_w = new_w & "&#" & ascw(mid(old,iStr,1)) & ";"
		end if
	next
	Chg_UNI=new_w
End Function
	
if  RSreg.eof then
      response.write "<script>alert('找不到資料');history.back();</script>"
      response.end
else

urlParams= "mcode=" & mcode & "&pagesize=" & PerPageSize & "&id_type1=" & id_type1 & "&id_type2=" & id_type2 & "&id_type3=" & id_type3 & "&scholarValidate=" & scholarValidate & "&status=" & status & "&account=" & account & "&realname=" & Server.URLEncode(trim(realname)) & "&nickname=" & Server.URLEncode(trim(nickname)) & "&keyword=" & Server.URLEncode(trim(keyword))  & "&emailaddress=" & Server.URLEncode(trim(emailaddress)) & "&homeaddr=" & Server.URLEncode(trim(homeaddr)) & "&idcard=" & Server.URLEncode(trim(idcard))

%>


<body>
    <div id="FuncName">
        <h1>
            會員整合查詢</h1>
        <div id="Nav">
            <a href="newMember_Query.asp" title="指定查詢條件">會員整合查詢</a> <!--A href="cp_member_new.htm" title="新增資料">新增會員</A-->
        </div>
        <div id="ClearFloat">
        </div>
    </div>
    <div id="FormName">
        會員整合查詢&nbsp; <font size="2">【主題單元:會員整合查詢 / 單元資料:純網頁】</font></div>
    <!--  條列頁簡易查詢功能  -->
    <form id="Form2" name="reg" method="POST" action="newMemberQuery_Act.asp?keep=Y&nowPage=<%=(nowPage)%>&pagesize=<%=PerPageSize%>&id_type1=<%=id_type1%>&id_type2=<%=id_type2%>&id_type3=<%=id_type3%>&scholarValidate=<%=scholarValidate%>&status=<%=status%>&account=<%=account%>&realname=<%=Server.URLEncode(trim(realname))%>&nickname=<%=Server.URLEncode(trim(nickname))%>&keyword=<%=Server.URLEncode(trim(keyword))%>">
        <input type="hidden" name="submitTask" value="">
        <input type="hidden" name="parentIcuitem" value="">
        <input type="hidden" name="selectedItems" value="">
        <div id="Page">
            共<em><%=totRec%></em>筆資料，目前在第<em><%=nowPage%>
                /<%=totPage%></em> 頁 每頁顯示
            <select id="PerPage" size="1" style="color: #FF0000" class="select">
                <option value="15" <%if perpagesize=15 then%> selected<%end if%>>15</option>
                <option value="30" <%if perpagesize=30 then%> selected<%end if%>>30</option>
                <option value="50" <%if perpagesize=50 then%> selected<%end if%>>50</option>
            </select>
            筆 <% if cint(nowPage) <> 1 then %>
            <img src="/images/arrow_previous.gif" alt="上一頁">
              ｜<a href="<%=HTProgPrefix%>.asp?keep=Y&nowPage=<%=(nowPage-1)%>&<%=urlParams %>">
                上一頁</a> ｜ <% end if %>
            跳至第
            <select id="GoPage" size="1" style="color: #FF0000" class="select">
                <% For iPage=1 to totPage %>
                <option value="<%=iPage%>" <%if ipage=cint(nowpage) then %>selected<%end if%>><%=iPage%>
                </option>
                <% Next %>
            </select>
            頁 <% if cint(nowPage)<>totPage then %>
            ｜<a href="<%=HTProgPrefix%>.asp?keep=Y&nowPage=<%=(nowPage+1)%>&<%=urlParams %>">下一頁
                <img src="/images/arrow_next.gif" alt="下一頁"></a> <% end if %>
        </div>
        <!-- 列表 -->
        <table cellspacing="0" id="ListTable">
            <tr>
                <th class="First" scope="col">&nbsp;</th>
                <th scope="col">會員帳號</th>
                <th scope="col">真實姓名</th>
                <th scope="col">電子郵件</th>
                <th class="eTableLable">申請日期</th>
                <th nowrap scope="col">身分</th>
                <th nowrap scope="col">學者審核</th>
                <th nowrap scope="col">專家</th>
                <th nowrap scope="col">狀態</th>
                <!--<th class="First" scope="col">預覽</th>   -->
            </tr>
            <% while not RsREG.eof %>
            <td align="center">
                <span class="eTableContent"><input type="checkbox" name="checkbox" id="<%=trim(RSreg("account"))%>"
                    value="checkbox">
                </span>
            </td>
            <td align="center">
                <a href="newMemberEdit.asp?account=<%=trim(RSreg("account"))%>&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>&from=query">
                    <%=CheckAndReplaceID(RSreg("account"))%>
                </a>
            </td>
            <td align="center">
                <%=CheckAndReplaceID(RSreg("realname"))%>
            </td>
            
            
	        <% If RSreg("mcode") = "Y" Then %>
	        <td align="center"><%=ReplaceEmailId(RSreg("email"))%>&nbsp;</td>
	        <% Else %>
	        <td align="center"><font color='grey'><%=ReplaceEmailId(RSreg("email"))%>&nbsp;(<%=RSreg("ValidCount")%>)</font></td>
	        <% End If %>            
            
            <td align="center">
                <%=RSreg("createtime")%>
            </td>
            <td align="center">
                <% if RSreg("id_type1")="1" then %>
                一般會員 <% end if%>
                <%if RSreg("id_type2")="1" then %>
                , 學者會員 <% end if%>
                &nbsp;
            </td>
            <td>
                <%if  RSreg("scholarValidate")="W" then%>
                <font color="#FF0000">待審核</font> <% end if%>
                <%if RSreg("scholarValidate")="N" then%>
                不通過 <% end if%>
                <%if RSreg("scholarValidate")="Y" then%>
                通過 <% end if%>
                <%if RSreg("scholarValidate")="Z" then%>
                無 <% end if%>
                &nbsp;
            </td>
            <td>
                <% if RSreg("id_type3")="1" then %>
                V <% end if%>
                &nbsp;
            </td>
            <td>
                <% if RSreg("status")="Y" then%>
                正常 <% end if%>
                <% if RSreg("status")="N" then%>
                停權 <% end if%>
                &nbsp;
            </td>
            </tr> <% 
 RSreg.movenext 
  wend
%>
        </table>
        <div align="center">
            <input type="button" name="button" value="整批通過" onclick="AllPassValidate()" class="cbutton" />
            <input type="button" name="button" value="整批不通過" onclick="AllNoPassValidate()" class="cbutton" />
        </div>
        <!-- 程式結束 ---------------------------------->
    </form>
</body>
</html>
<%  end if%>

<script type="text/javascript"  >
	
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

<script language="VBScript">
	
		sub formModSubmit()
			reg.submitTask.value = "DELETE"
			reg.action = "<%=HTProgPrefix%>List.asp?icuitem=<%=request.querystring("icuitem")%>"
			reg.Submit
		end sub
		
		dim gpKey
    sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>.asp?keep=Y&nowPage=" & newPage & "&<%=urlParams%>"
		 
         
		   
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>.asp?keep=Y&nowPage=<%=nowPage%>" & "&<%=urlParams%>&newPerPage=" & newPerPage
     end sub 

     sub shortLongList(sorl)
           document.location.href="<%=HTProgPrefix%>.asp?keep=Y&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>" _
			& "&shortLongList=" & sorl
     end sub   

	</script>

