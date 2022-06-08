<% Response.Expires = 0
HTProgCap="課程"
HTProgFunc="編修"
HTProgCode="PA001"
HTProgPrefix="paAct" %>
<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
 Set RSreg = Server.CreateObject("ADODB.RecordSet")

 fSql="select * from docdownload_attach where pid=" & request("id")
 
 set rs=conn.Execute(fSql)
 

' 取得目前資料表記錄的頁數
pageNo = Request.QueryString("PageNo")
If pageNo = "" Then
   intPageNo = 1
Else
   intPageNo = Convert.ToInt32(pageNo)
End If


strSQL = "Select count(*) From docdownload_attach where pid=" & request("id")

' 建立Command物件的SQL指令     
set rs2=conn.Execute(strSQL)
intMaxRec = rs2(0)
if intMaxRec > 0 then              
   intMaxPageCount=intMaxRec\10
   if (intMaxRec mod 10) <> 0 then
       intMaxPageCount=intMaxPageCount+1
   end if
else
   intMaxPageCount=1
end if    
'***** 計算頁數結束*****

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title>講義下載
</title>
</head>

<body>
<div id="FuncName">
	<h1>講義下載</h1>
	<div id="Nav">
		
			<a href="paActList.asp">回條列</a>
			<A href="docAttachAdd.asp?pid=<% =request("id") %>" title="新增資料">新增講義</A>
		
	</div>
	<div id="ClearFloat"></div>
</div>

<p>講義列表</p>
<Form id="Form2" name=reg method="POST" action=docAttach.asp>
	<!-- 分頁 -->
	<div id="Page">
             
		共<em><% =intMaxRec %></em>筆資料，目前在第 
               <select name="select" style="color:#FF0000"  onchange=Page_Onchange(this.form) class="select">
         <%
           n = 1
           do While n <= intMaxPageCount 
                 Response.Write "<option value='" & n & "'" 
                 If n = intPageNo Then 
                    Response.Write " selected " 
                 End If
                 Response.Write ">" & n & "</option>" 
                 n = n + 1
           loop
      %> 
               </select>
          頁
         
	</div>



	<!-- 列表 -->
	
	<table cellspacing="0" id="ListTable">
		<tr>
			
			<th align="left">　檔案 </th>
<th align="left">　說明 </th></tr>       
<%
  intCount=1
   do while not rs.eof 
    if intCount<=(intPageNo*10) and intCount>(intPageNo-1)*10 then 
%>           
<tr>                  
	
	<TD class=eTableContent>
		<a href="docattachEdit.asp?id=<% =trim(rs("id")) %>"><% =intCount & ". " %><% =trim(rs("filename")) %></a>
	</TD>

	<TD class=eTableContent>
		<% =trim(rs("filename_note"))  %>&nbsp;
</font></td>

    </tr>
<% 
   end if
   intCount=intCount+1
   rs.MoveNext
   loop                      
%>
    </TABLE>
       <!-- 程式結束 ---------------------------------->  
     </form>  
</body>
</html>                                 


<script Language=VBScript>
  dim gpKey
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="docAttach.asp?keep=Y&nowPage=" & newPage & "&pagesize=15"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="docAttach.asp?keep=Y&nowPage=1" & "&pagesize=" & newPerPage                    
     end sub 

     sub setpKey(xv)
     	gpKey = xv
     end sub


Dim CanTarget
Dim followCanTarget

sub popCalendar(dateName,followName)        
 	CanTarget=dateName
 	followCanTarget=followName
	xdate = document.all(CanTarget).value
	if not isDate(xdate) then	xdate = date()
	document.all.calendar1.setDate xdate
	
 	If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 	Else        
       ex=window.event.clientX
       ey=document.body.scrolltop+window.event.clientY+10
       if ex>520 then ex=520
       document.all.calendar1.style.pixelleft=ex-80
       document.all.calendar1.style.pixeltop=ey
       document.all.calendar1.style.visibility=""
 	End If              
end sub     

sub calendar1_onscriptletevent(n,o) 
    	document.all("CalendarTarget").value=n         
    	document.all.calendar1.style.visibility="hidden"    
    if n <> "Cancle" then
        document.all(CanTarget).value=document.all.CalendarTarget.value
        document.all("pcShow"&CanTarget).value=document.all.CalendarTarget.value
        if followCanTarget<>"" then
        	document.all(followCanTarget).value=document.all.CalendarTarget.value
        	document.all("pcShow"&followCanTarget).value=document.all.CalendarTarget.value
        end if
	end if
end sub   
</script>