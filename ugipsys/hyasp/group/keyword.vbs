'----DB連線
DBConnStr="Provider=SQLOLEDB;server=10.10.10.59;User ID=hyGIP;Password=hyweb;Database=GIPhy"
wScript.Echo("open db conn..."&DBConnStr)
	Set Conn = CreateObject("ADODB.Connection")
	Conn.Open(DBConnStr)
wScript.Echo("db conn opened.")	
'----取得WTP資料並放入陣列中
	SQLWTP="Select iKeyword,iKeywordNew,KeywordStatus from CuDTKeywordWTP where KeywordStatus<>'P' order by KeywordStatus"
	Set RSWTP=conn.execute(SQLWTP)
	if RSWTP.EOF then 
		wScript.Echo("Recordset not found!.")
    		WScript.Quit(1)
	end if	
	WTPArray=RSWTP.getrows()
wScript.Echo("Current Issue: SQL資料放入陣列中......")
'----開檔案物件,並逐一比對WTP資料陣列後Match
inFile = "C:\Hysearch\group\keyword.txt"
outFile = "C:\Hysearch\group\keywordtmp.txt"
Set fso = CreateObject("scripting.filesystemobject")

set xfin = fso.openTextFile(inFile)
set xfout = fso.createTextFile(outFile)

wScript.Echo("關鍵字詞txt檔更新中......")

Do while not xfin.AtEndOfStream
    xinStr = xfin.readline
    writeFlag=true
    for i=0 to ubound(WTPArray,2)
	if trim(xinStr)=trim(WTPArray(0,i)) then
	    if WTPArray(2,i)="E" then		'----編修部份
	    	xfout.writeline	WTPArray(1,i)
		SQLDelete="Delete CuDTKeywordWTP where iKeyword='"+WTPArray(0,i)+"'"   
		Set RSE=conn.execute(SQLDelete)	    
	    elseif WTPArray(2,i)="D" then	'----刪除部份
		SQLDelete="Delete CuDTKeywordWTP where iKeyword='"+WTPArray(0,i)+"'"    
		Set RSD=conn.execute(SQLDelete)		
	    end if
	    writeFlag=false
	    exit for
	end if
    next
    if writeFlag then xfout.writeline trim(xinStr)	
Loop

for i=0 to ubound(WTPArray,2)
	if WTPArray(2,i)="A" then 		'----新增部份
		xfout.writeline	WTPArray(0,i)
		SQLDelete="Delete CuDTKeywordWTP where iKeyword='"+WTPArray(0,i)+"'"
		Set RSA=conn.execute(SQLDelete)		
	end if
next
xfin.close
xfout.close

fso.CopyFile outFile, inFile

wScript.Echo("關鍵字詞txt檔更新完成!")

