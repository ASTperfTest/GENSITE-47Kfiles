<%

'// mp 專用的 URL check 副程式.
sub mpSQLinjectionCheck()
    dim bHaveErrorWhenCheck
    dim xf
    dim myqStr
    dim myqStrField

    bHaveErrorWhenCheck = 0     '// no error
    myqStr = request.queryString

    for each xf in request.queryString
        if request(xf) <> "" then
            myqStrField = request(xf)
            if instr(myqStrField,"'") > 0 or instr(myqStrField,"""") > 0 or instr(myqStrField,"@") > 0 or instr(myqStrField,"<") > 0 or instr(myqStrField,">") > 0 or instr(myqStrField,"+") > 0  or instr(myqStrField,"-") > 0 or instr(myqStrField,"*") > 0 or instr(myqStrField," ") > 0 or instr(myqStrField,";") > 0 then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 1
            end if
        end if
    next

    for each xf in request.form
        if request(xf) <> "" then
            myqStrField = request(xf)
            if instr(myqStrField,"'") > 0 or instr(myqStrField,"""") > 0 or instr(myqStrField,"@") > 0 or instr(myqStrField,"<") > 0 or instr(myqStrField,">") > 0 or instr(myqStrField,"+") > 0  or instr(myqStrField,"-") > 0 or instr(myqStrField,"*") > 0 or instr(myqStrField," ") > 0 or instr(myqStrField,";") > 0  then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 2
            end if
        end if
    next

    if bHaveErrorWhenCheck = 1 then
        call EndOfMyASP(mp)
    end if
end sub


'// rss 專用的 URL check 副程式.
sub rssSQLinjectionCheck()
    dim bHaveErrorWhenCheck
    dim xf
    dim myqStr
    dim myFieldName
    dim myFilterSymbols

    bHaveErrorWhenCheck = 0     '// no error
    myqStr = request.queryString

    if bHaveErrorWhenCheck <> 1 then
        '// rss 用參數
        myFieldName = "xNode"
        myFilterSymbols = "'""@<>%andor~!#$^&*()-=+"
        if checkSymbolInField(myFieldName,myFilterSymbols) then
            bHaveErrorWhenCheck = 1
            strErrorNumber = 7
        end if

        myFieldName = "cuitem"
        myFilterSymbols = "'""@<>%andor~!#$^&*()-=+"
        if checkSymbolInField(myFieldName,myFilterSymbols) then
            bHaveErrorWhenCheck = 1
            strErrorNumber = 7
        end if

        if len(request("cuitem")) > 9 then
            bHaveErrorWhenCheck = 1
            strErrorNumber = 128
        end if

    end if

    for each xf in request.form
        if request(xf) <> "" then
            myqStr = request(xf)
            if instr(myqStr,"'") > 0 or instr(myqStr,"""") > 0 or instr(myqStr,"@") > 0 or instr(myqStr,"<") > 0 or instr(myqStr,">") > 0 then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 22
            end if
        end if
    next

    if bHaveErrorWhenCheck = 1 then
        call EndOfMyASP(mp)
    end if
end sub


'// lp 專用的 URL check 副程式.
sub lpSQLinjectionCheck()
    dim bHaveErrorWhenCheck
    bHaveErrorWhenCheck = 0     '// no error

    if bHaveErrorWhenCheck <> 1 then
        bHaveErrorWhenCheck = cpCommonCheck(strErrorNumber)
    end if

    if bHaveErrorWhenCheck <> 1 then
        bHaveErrorWhenCheck = lpCommonCheck(strErrorNumber)
    end if

    if bHaveErrorWhenCheck = 1 then
        call EndOfMyASP(mp)
    end if
end sub

'// cp 專用的 URL check 副程式.
sub cpSQLinjectionCheck()
    dim bHaveErrorWhenCheck
    bHaveErrorWhenCheck = 0     '// no error

    if bHaveErrorWhenCheck <> 1 then
        bHaveErrorWhenCheck = cpCommonCheck(strErrorNumber)
    end if

    if bHaveErrorWhenCheck = 1 then
        call EndOfMyASP(mp)
    end if
end sub

'// sp 專用的 URL check 副程式.
sub spSQLinjectionCheck()
    dim bHaveErrorWhenCheck

    dim myFieldName
    dim myFilterSymbols
    dim iPos

    bHaveErrorWhenCheck = 0     '// no error
    strErrorNumber = 0
    myFieldName = ""
    myFilterSymbols = ""


    '// check xdURL 必需要有值.
    if request("xdurl") = "" then
        bHaveErrorWhenCheck = 1
        strErrorNumber = 0
    else
        if instr(request("xdurl"),".asp") > 0 then
            '// ok.
        else
            '// error.
            bHaveErrorWhenCheck = 1
            strErrorNumber = 1
        end if
    end if

    '// check GIP 常用參數
    if bHaveErrorWhenCheck <> 1 then
        bHaveErrorWhenCheck = cpCommonCheck(strErrorNumber)
    end if

    '// lp 常用檢測項目.
    if bHaveErrorWhenCheck <> 1 then
        bHaveErrorWhenCheck = lpCommonCheck(strErrorNumber)
    end if

    '// sp 常用檢測項目.
    if bHaveErrorWhenCheck <> 1 then
        bHaveErrorWhenCheck = spCommonCheck(strErrorNumber)
    end if

    if bHaveErrorWhenCheck <> 1 then
        '// 單獨對 ap 做判斷.
        if instr(ucase(request("xdurl")),"MAIL/MAIL") > 0 then
            myFieldName = "content"
            if request(myFieldName) <> "" then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 102
            end if

            myFieldName = "event_title"
            if request(myFieldName) <> "" then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 102
            end if

            myFieldName = "library"
            if request(myFieldName) <> "" then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 103
            end if

            myFieldName = "m011_type"
            if request(myFieldName) <> "" then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 104
            end if

            myFieldName = "realname"
            if request(myFieldName) <> "" then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 105
            end if

            myFieldName = "subject"
            if request(myFieldName) <> "" then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 106
            end if

            myFieldName = "subjectid"
            if request(myFieldName) <> "" then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 107
            end if

            myFieldName = "xpostdate_year"
            if request(myFieldName) <> "" then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 107
            end if
        end if

        if request("xdurl") = "mail/mail_act.asp"  then
            myFieldName = "event_title"
            if request(myFieldName) <> "" then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 202
            end if
        end if

        '// 單獨對 ap 做判斷.
        if instr(ucase(request("xdurl")),"LIB_PUB/LIB") > 0 then
            myFieldName = "id"
            if request(myFieldName) <> "" then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 301
            end if

            myFieldName = "event_title"
            if request(myFieldName) <> "" then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 302
            end if

            myFieldName = "pkind"
            if request(myFieldName) <> "" then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 303
            end if

            myFieldName = "pub_id"
            myFilterSymbols = "'""@<>%&()-=+"
            if checkSymbolInField(myFieldName,myFilterSymbols) then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 304
            end if
        end if


        '// 單獨對 ap 做判斷.
        if instr(ucase(request("xdurl")),"ATTEND/PE") > 0 then
            myFieldName = "id"
            myFilterSymbols = "'""@<>%&()-=+*/"
            if checkSymbolInField(myFieldName,myFilterSymbols) then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 304
            end if

            myFieldName = "id"
            if not isnumeric(request(myFieldName)) then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 301
            end if

            myFieldName = "ParentSortPage"
            myFilterSymbols = "'""@<>%&()=-+/"
            if checkSymbolInField(myFieldName,myFilterSymbols) then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 304
            end if

            myFieldName = "yymm"
            myFilterSymbols = "'""@<>%&()=+*andor"
            if checkSymbolInField(myFieldName,myFilterSymbols) then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 304
            end if

            '// 哎... 這個變本, 根本不會出現在程式碼裡, 但還是要判斷一下, 不然會被測出來.
            myFieldName = "JOURNAL_INFO"
            myFilterSymbols = "'""@<>%&()-=+*/"
            if checkSymbolInField(myFieldName,myFilterSymbols) then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 304
            end if

            '//身分證字號
            myFieldName = "Ident"
            myFilterSymbols = "'""@<>%&()-=+*/"
            if checkSymbolInField(myFieldName,myFilterSymbols) then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 304
            end if

            '//姓名
            myFieldName = "Name"
            myFilterSymbols = "'""@<>%&()=+*/"
            if checkSymbolInField(myFieldName,myFilterSymbols) then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 304
            end if

            '//出生年月日
            myFieldName = "BirthDt"
            myFilterSymbols = "'""@<>%&()=+*"
            if checkSymbolInField(myFieldName,myFilterSymbols) then
                bHaveErrorWhenCheck = 1
                strErrorNumber = 304
            end if




        end if
    end if

    if bHaveErrorWhenCheck = 1 then
        call EndOfMyASP(mp)
    end if

end sub


'// symbol in variable return true.
function checkSymbolInField(byval variableName, byval symbols)
    dim returnValue
    dim singleSymbols
    dim iSymbol

    returnValue = False
    singleSymbols = ""
    iSymbol = 0

    if symbols <> "" then
        if len(symbols)=1 then
            if instr(request(variableName),symbols) >0 then
                returnValue = True
            end if
        else
            for iSymbol = 1 to len(symbols)
                if instr(request(variableName),mid(symbols,iSymbol,1)) >0 then
                    returnValue = True
                end if
            next
        end if
    end if

    checkSymbolInField = returnValue
end function


'// purpose: return bHaveErrorWhenCheck
function cpCommonCheck(byref strErrorNumber)
    dim returnValue
    dim myFieldName
    dim myFilterSymbols

    returnValue = 0
    strErrorNumber = 0
    myFieldName = ""
    myFilterSymbols = ""

    myFieldName = "mp"
    myFilterSymbols = "'""@<>%andor~`!#$^&*/\.(){}[]-=+;"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 10
    end if

    myFieldName = "CtNode"
    myFilterSymbols = "'""@<>%andor~`!#$^&*/\.(){}[]-=+;"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 4
    end if

    myFieldName = "CtUnit"
    myFilterSymbols = "'""@<>%andor~`!#$^&*/\.(){}[]-=+;"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 5
    end if

    myFieldName = "BaseDSD"
    myFilterSymbols = "'""@<>%andor~`!#$^&*/\.(){}[]-=+;"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 6
    end if

    myFieldName = "icuitem"
    myFilterSymbols = "'""@<>%andor~`!#$^&*/\.(){}[]-=+;"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 7
    end if

    myFieldName = "xitem"
    myFilterSymbols = "'""@<>%andor~`!#$^&*/\.(){}[]-=+;"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 7
    end if

    '// check isnumeric
    if request("CtNode") <> "" then
        if not isnumeric(request("CtNode")) then
            returnValue = 1
            strErrorNumber = 50
        end if
    end if

    if request("CtUnit") <> "" then
        if not isnumeric(request("CtUnit")) then
            returnValue = 1
            strErrorNumber = 50
        end if
    end if

    if request("BaseDSD") <> "" then
        if not isnumeric(request("BaseDSD")) then
            returnValue = 1
            strErrorNumber = 50
        end if
    end if

    if request("icuitem") <> "" then
        if not isnumeric(request("icuitem")) then
            returnValue = 1
            strErrorNumber = 50
        end if
    end if

    if request("xNode") <> "" then
        if not isnumeric(request("xNode")) then
            returnValue = 1
            strErrorNumber = 50
        end if
    end if

    if request("xitem") <> "" then
        if not isnumeric(request("xitem")) then
            returnValue = 1
            strErrorNumber = 50
        end if
    end if


    '// check length
    if len(request("CtNode")) > 6 then
        returnValue = 1
        strErrorNumber = 125
    end if
    if len(request("CtUnit")) > 5 then
        returnValue = 1
        strErrorNumber = 126
    end if
    if len(request("BaseDSD")) > 4 then
        returnValue = 1
        strErrorNumber = 127
    end if
    if len(request("mp")) > 6 then
        returnValue = 1
        strErrorNumber = 128
    end if

    if len(request("icuitem")) > 9 then
        returnValue = 1
        strErrorNumber = 128
    end if

    cpCommonCheck = returnValue
end function


'// purpose: return bHaveErrorWhenCheck
function lpCommonCheck(byref strErrorNumber)
    dim returnValue
    dim myFieldName
    dim myFilterSymbols

    returnValue = 0
    strErrorNumber = 0
    myFieldName = ""
    myFilterSymbols = ""

    myFieldName = "htx_idept"
    myFilterSymbols = "'""@<>%~!#$^&*()-=+"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 8
    end if

    myFieldName = "htx_topCat"
    myFilterSymbols = "'""@<>%~!#$^&*()-=+"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 9
    end if

    myFieldName = "xq_xCat"
    myFilterSymbols = "'""@<>%~!#$^&*()-=+"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 9
    end if

    myFieldName = "xq_xCat2"
    myFilterSymbols = "'""@<>%~!#$^&*()-=+"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 10
    end if

    myFieldName = "htx_xpostDate"
    myFilterSymbols = "'""@<>%~!#$^&*()-=+andor"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 10
    end if

    myFieldName = "htx_xpostDateS"
    myFilterSymbols = "'""@<>%~!#$^&*()-=+andor"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 10
    end if

    myFieldName = "htx_xpostDateE"
    myFilterSymbols = "'""@<>%~!#$^&*()-=+andor"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 10
    end if

    myFieldName = "nowPage"
    myFilterSymbols = "'""@<>%~!#$^&*()-=+"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 10
    end if

    myFieldName = "pagesize"
    myFilterSymbols = "'""@<>%~!#$^&*()-=+"
    if checkSymbolInField(myFieldName,myFilterSymbols) then
        returnValue = 1
        strErrorNumber = 10
    end if

    '// length check
    if len(request("htx_idept")) > 10 then
        returnValue = 1
        strErrorNumber = 128
    end if

    if len(request("htx_topCat")) > 10 then
        returnValue = 1
        strErrorNumber = 128
    end if

    if len(request("xq_xCat")) > 10 then
        returnValue = 1
        strErrorNumber = 128
    end if

    if len(request("xq_xCat2")) > 10 then
        returnValue = 1
        strErrorNumber = 128
    end if

    if len(request("htx_xpostDate")) > 20 then
        returnValue = 1
        strErrorNumber = 128
    end if

    if len(request("htx_xpostDateS")) > 20 then
        returnValue = 1
        strErrorNumber = 128
    end if

    if len(request("htx_xpostDateE")) > 20 then
        returnValue = 1
        strErrorNumber = 128
    end if

    if len(request("nowPage")) > 4 then
        returnValue = 1
        strErrorNumber = 128
    end if

    if len(request("pagesize")) > 4 then
        returnValue = 1
        strErrorNumber = 128
    end if

    lpCommonCheck = returnValue
end function


'// purpose: return bHaveErrorWhenCheck
function spCommonCheck(byref strErrorNumber)
pArray=array("stamp", "card","rname","remail2","fname","femail","subject","fmsg","s_id","mp","ctNode","realname","email","content","strCAPTCHA","checkstr")

for each str in pArray

p=request(str)

        '要檢查的 pattern

        patern=array("<", ">")



        for each ptn in patern
                if (Instr(p, ptn) >0) then
                   returnValue = 1
                   strErrorNumber = 10
                end if
        next
next

    spCommonCheck = returnValue
end function
sub EndOfMyASP(byval mp)
    response.write "<html>"
    response.write "<head><meta http-equiv='Content-Type' content='text/html; charset="& response.charset &"'>"
    response.write "<title>" & Request.ServerVariables("SERVER_NAME") & "</title>"
    response.write "</head>" & vbCrLf
    response.write "<body>" & vbCrLf
    response.write "<script language='JavaScript'>location.replace('mp.asp?mp=" & mp & "');</script>" & vbCrLf
    response.write "<noscript>this javascrip is used to redirect to <a href='http://" & Request.ServerVariables("SERVER_NAME") & "/mp.asp?mp=" & mp &"'>home</a> page.</noscript>" & vbCrLf
    response.write "<!-- "&strErrorNumber&" -->" & vbCrLf
    response.write "</body>" & vbCrLf
    response.write "</html>"
    Response.End()
end sub
function getMPvalue()
    dim myMP

    myMP = request("mp")
    if myMP = "" then
        myMP = session("mpTree")
    else
        if isnumeric(myMP) then
            myMP = int(myMP)
        else
            myMP = session("mpTree")
            call EndOfMyASP(myMP)
        end if
    end if


	getMPvalue = myMP
end function
%>