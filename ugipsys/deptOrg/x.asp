<%@ CodePage = 65001 %>

<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title>系統管理／組織架構</title>
</head>


<script language=VBS>
function blen(xs)              '檢測中英文夾雜字串實際長度
  xl = len(xs)
  for i=1 to len(xs)
	if asc(mid(xs,i,1))<0 then xl = xl + 1
  next
  blen = xl
end function

function xStdTime(dt)  '轉成民國年及時間  999/99/99 10:00 給資料型態為DateTime 使用
   if Len(dt)=0 or isnull(dt) then
   	xStdTime=""
   else
   	xyear = right("000"+ cstr((year(dt)-1911)),3)     '補零
        xmonth = right("00"+ cstr(month(dt)),2)
        xday = right("00"+ cstr(day(dt)),2)
        xhour = right("00" + cstr(hour(dt)),2)
        xminute = right("00" + cstr(minute(dt)),2)
        xStdTime = xyear & "/" & xmonth & "/" & xday & " " & xhour & ":" & xminute
   end if
end function

function d7date(dt)     '轉成民國年  999/99/99 給資料型態為SmallDateTime 使用
	if Len(dt)=0 or isnull(dt) then
		d7date=""
	else
           xy=right("000"+ cstr((year(dt)-1911)),3)     '補零
           xm=right("00"+ cstr(month(dt)),2)
           xd=right("00"+ cstr(day(dt)),2)
           d7date=xy & "/" & xm & "/" & xd
           d7date=dt
        end if
end function

 function ChkIDX(SID)      '-----身份證字號檢驗

   SID = UCase(SID)
   X1 = InStr("ABCDEFGHJKLMNPQRSTUVWXYZIO", Left(SID, 1)) + 9

   ' 初步檢查 ID 的合法性
   If X1 < 10 Then Chk_ID = False: Exit Function
   If Not IsNumeric(Mid(SID, 2)) Then Chk_ID = False: Exit Function
   If Len(SID) <> 10 Then Chk_ID = False: Exit Function

   ' 檢查編號之正確性
   SID = Cstr(X1) + Mid(SID, 2)

   nCheckSum = Cint(Mid(SID, 1, 1))
   For I = 2 To Len(SID) - 1
        nCheckSum = nCheckSum + Cint(Mid(SID, I, 1)) * (11 - I)
   Next
   nCheckSum = nCheckSum + Cint(Mid(SID, 11, 1))

   ChkID = (nCheckSum Mod 10 = 0)
end Function

function ChkID(idno)	'-----身份證字號檢驗(新)
  if len(trim(idno))=0 then exit function

  alpha=UCase(left(idno,1))
  d1=mid(idno,2,1)
  d2=mid(idno,3,1)
  d3=mid(idno,4,1)
  d4=mid(idno,5,1)
  d5=mid(idno,6,1)
  d6=mid(idno,7,1)
  d7=mid(idno,8,1)
  d8=mid(idno,9,1)
  d9=mid(idno,10,1)
  select case alpha
    case "A" : acc=1
    case "B" : acc=10
    case "C" : acc=19
    case "D" : acc=28
    case "E" : acc=37
    case "F" : acc=46
    case "G" : acc=55
    case "H" : acc=64
    case "I" : acc=39
    case "J" : acc=73
    case "K" : acc=82
    case "L" : acc=2
    case "M" : acc=11
    case "N" : acc=20
    case "O" : acc=48
    case "P" : acc=29
    case "Q" : acc=38
    case "R" : acc=47
    case "S" : acc=56
    case "T" : acc=65
    case "U" : acc=74
    case "V" : acc=83
    case "W" : acc=21
    case "X" : acc=3
    case "Y" : acc=12
    case "Z" : acc=30
  end select
  on error resume next
  checksum = acc+8*d1+7*d2+6*d3+5*d4+4*d5+3*d6+2*d7+1*d8+1*d9
  check1 = Int(checksum/10)
  check2 = checksum/10
  check3 = (check2-check1)*10
  if len(idno)>10 then
    ChkID=false
  elseif err.number<>0 then
    ChkID=false
  elseif checksum=check1*10 then
    ChkID=true
  elseif d9=(10-check3) then
    ChkID=true
  else
    ChkID=false
  end if
end function
function ftpDo(FTPIP,FTPPort,FTPID,FTPPWD,fileAction,FTPfilePath,fileDir,fileTarget,fileSource)	'----FTP機制  2004/7/7
response.end
	Set oFtp = CreateObject("FtpCom.FTP")
	oFtp.Connect FTPIP,FTPPort,FTPID,FTPPWD
	
	if left(oFtp.GetMsg,1) = "2" then 
		if FTPfilePath <> "" then oFtp.Execute("CWD " + FTPfilePath)
		if fileAction="CreateDir" then oFtp.CreateDir(fileDir)
		if fileAction="DeleteDir" then oFtp.DeleteDir(fileDir)
		if fileAction="MoveFile" then oFtp.MoveFile fileTarget,fileSource,1,0
		if fileAction="DELE" then oFtp.Execute("DELE " + fileTarget)
		if left(oFtp.GetMsg,1) <> "2" then FTPErrorMSG="  FTP機制出現錯誤,FTP未成功!"
		oFtp.LogOffServer 
		set oFtp = nothing
	else
		FTPErrorMSG="  FTP機制出現錯誤,FTP未成功!"
		oFtp.LogOffServer 
		set oFtp = nothing
	end if	
end function

</script>
<body>
<div id="FuncName">
	<h1>系統管理／組織架構</h1>
	<div id="Nav">
	    <a href="VBScript: history.back">回上一頁</a>
	    <a href="deptAdd.asp?deptID=011">新增從屬部門</a>
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">編修組織</div>


    <script language="vbscript">
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
  
	IF reg.htx_deptName.value = Empty Then 
		MsgBox replace(nMsg,"{0}","部門中文名稱"), 64, "Sorry!"
		reg.htx_deptName.focus
		exit sub
	END IF
	IF blen(reg.htx_deptName.value) > 70 Then
		MsgBox replace(replace(lMsg,"{0}","部門中文名稱"),"{1}","70"), 64, "Sorry!"
		reg.htx_deptName.focus
		exit sub
	END IF
	IF blen(reg.htx_AbbrName.value) > 30 Then
		MsgBox replace(replace(lMsg,"{0}","單位簡稱"),"{1}","30"), 64, "Sorry!"
		reg.htx_AbbrName.focus
		exit sub
	END IF
	IF blen(reg.htx_eDeptName.value) > 60 Then
		MsgBox replace(replace(lMsg,"{0}","英文名稱"),"{1}","60"), 64, "Sorry!"
		reg.htx_eDeptName.focus
		exit sub
	END IF
	IF blen(reg.htx_eAbbrName.value) > 30 Then
		MsgBox replace(replace(lMsg,"{0}","英文簡稱"),"{1}","30"), 64, "Sorry!"
		reg.htx_eAbbrName.focus
		exit sub
	END IF
	IF reg.htx_inUse.value = Empty Then
		MsgBox replace(nMsg,"{0}","是否有效"), 64, "Sorry!"
		reg.htx_inUse.focus
		exit sub
	END IF
  
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub

   sub formDelSubmit()
         chky=msgbox("注意！"& vbcrlf & vbcrlf &"　你確定刪除資料嗎？"& vbcrlf , 48+1, "請注意！！")
        if chky=vbok then
	       reg.submitTask.value = "DELETE"
	      reg.Submit
       end If
  end sub

</script>

<form id="Form1" method="GET" name="reg" action="xx.asp">
<INPUT TYPE=hidden name=submitTask value="">
<input type="hidden" name="htx_deptID">
<CENTER><TABLE  cellspacing="0">
<TR><TD class="Label" align="right"><font color="red">*</font>部門中文名稱：</TD>
<TD class="whitetablebg"><input name="htx_deptName" size="70">
</TD></TR>
<TR><TD class="Label" align="right">單位簡稱：</TD>
<TD class="whitetablebg"><input name="htx_AbbrName" size="30">
</TD></TR>
<TR><TD class="Label" align="right">英文名稱：</TD>
<TD class="whitetablebg"><input name="htx_eDeptName" size="50">
</TD></TR>
<TR><TD class="Label" align="right">英文簡稱：</TD>
<TD class="whitetablebg"><input name="htx_eAbbrName" size="30">
</TD></TR>

<TR><TD class="Label" align="right">顯示次序：</TD>
<TD class="whitetablebg"><input name="htx_seq" size="2">
</TD></TR>
<TR><TD class="Label" align="right">單位層級：</TD>
<TD class="whitetablebg"><Select name="htx_OrgRank" size=1>
<option value="">請選擇</option>
			

			<option value="0">頂級</option>
			

			<option value="1">一級</option>
			

			<option value="2">二級</option>
			
</select></TD></TR>
<TR><TD class="Label" align="right">類別：</TD>
<TD class="whitetablebg"><Select name="htx_kind" size=1>
<option value="">請選擇</option>
			

			<option value="0">本部</option>
			

			<option value="1">所屬單位</option>
			

			<option value="2">所屬機關</option>
			

			<option value="3">所屬機構</option>
			
</select></TD></TR>
<TR><TD class="Label" align="right">是否有效：</TD>
<TD class="whitetablebg"><Select name="htx_inUse" size=1>
<option value="">請選擇</option>
			

			<option value="Y">是</option>
			

			<option value="N">否</option>
			
</select></TD></TR>
<TR><TD class="Label" align="right">資料大類：</TD>
<TD class="whitetablebg"><table width="100%"><TR>
			
			<td><input type="checkbox" name="htx_tDataCat" value="1" ><font size=2>類別一</font></td>
			
			<td><input type="checkbox" name="htx_tDataCat" value="2" ><font size=2>類別二</font></td>
			
			<td><input type="checkbox" name="htx_tDataCat" value="3" ><font size=2>類別三</font></td>
			
			<td><input type="checkbox" name="htx_tDataCat" value="4" ><font size=2>類別四</font></td>
			</TR><TR>
			<td><input type="checkbox" name="htx_tDataCat" value="5" ><font size=2>類別五</font></td>
			
			<td><input type="checkbox" name="htx_tDataCat" value="6" ><font size=2>類別六</font></td>
			
			<td><input type="checkbox" name="htx_tDataCat" value="7" ><font size=2>類別七</font></td>
			
			<td><input type="checkbox" name="htx_tDataCat" value="8" ><font size=2>類別八</font></td>
			
</TR></table>
</TD></TR>
</TABLE>
</CENTER>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">

               <input type=button value ="編修存檔" class="cbutton" onClick="formModSubmit()">
        
            <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
		           
        <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">


 </td></tr>
</table>
<SCRIPT LANGUAGE="vbs">
    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
    end sub
</SCRIPT>
	

<SCRIPT LANGUAGE="vbs">

</SCRIPT>
		
    <script language=vbs>
      sub resetForm 
	       reg.reset()
	       clientInitForm
      end sub

     sub clientInitForm
	reg.htx_deptID.value= "011"
	reg.htx_deptName.value= "台北地檢"
	reg.htx_AbbrName.value= "tpc"
	reg.htx_eDeptName.value= "tpc"
	reg.htx_eAbbrName.value= "tpc"

	reg.htx_OrgRank.value= "1"
	reg.htx_seq.value= "1"
	reg.htx_kind.value= "1"
	reg.htx_inUse.value= "Y"
	initCheckbox "htx_tDataCat",""
    end sub

    sub window_onLoad
         clientInitForm
    end sub
    
    sub initRadio(xname,value)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    end sub

    sub initOtherRadio(xname,value, otherName)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    	if value="" then	exit sub
		reg.all(xname).item(reg.all(xname).length-1).checked = true
		reg.all(otherName).value = value
		reg.all(xname).item(reg.all(xname).length-1).value = value
    end sub

    sub initCheckbox(xname,ckValue)
    	value = ckValue & ","
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    		end if
    	next
    end sub
    
    sub initOtherCheckbox(xname,ckValue,otherName)
    	valueArray = split(ckValue,", ")
    	valueCount = ubound(valueArray) + 1
    	value = ckValue & ","
    	ckCount = 0
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    			ckCount = ckCount + 1
    		end if
    	next
		if ckCount <> valueCount then
			reg.all(xname).item(reg.all(xname).length-1).checked = true
			reg.all(otherName).value = valueArray(ubound(valueArray))
			reg.all(xname).item(reg.all(xname).length-1).value = valueArray(ubound(valueArray))
		end if
    end sub

    sub initImgFile(xname, value)
		reg.all("htImgActCK_"&xname).value=""
		reg.all("htImg_"&xname).style.display="none"
		reg.all("hoImg_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
	   		document.all("logo_"&xname).src= "/public/" & value
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
	end sub

Sub addLogo(xname)	'新增logo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="addLogo"
End sub

Sub chgLogo(xname)	'更換logo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="editLogo"
End sub

Sub orgLogo(xname)	'原圖
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value=""
End sub

Sub delLogo(xname)	'刪除logo
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("noLogo_"&xname).style.display=""
    reg.all("htImg_"&xname).value=""
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value="delLogo"
End sub

    sub initAttFile(xname, value, orgValue)
		reg.all("htFileActCK_"&xname).value=""
		reg.all("htFile_"&xname).style.display="none"
		reg.all("hoFile_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
	   		document.all("logo_"&xname).innerText= orgValue
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
	end sub

Sub addXFile(xname)	'新增logo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="addLogo"
End sub

Sub chgXFile(xname)	'更換logo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="editLogo"
End sub

Sub orgXFile(xname)	'原圖
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value=""
End sub

Sub delXFile(xname)	'刪除logo
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("noLogo_"&xname).style.display=""
    reg.all("htFile_"&xname).value=""
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value="delLogo"
End sub

 </script>	
     
<div id="Explain">
	<h1>說明</h1>
	<ul>
		<li><span class="Must">*</span>為必要欄位</li>
	</ul>
</div>
</body>
</html>
