﻿<%
function maxEncrypt(byval inputString)
    dim myOffsetList 
    dim i
    dim newStr
    
    '// initial
    myOffsetList = "32165498765432165498765432165216549876543216549876543216498765454321654987654321654987654321654987654321654654321654987654321654987654321316549876543216549876543216549876543216548796543216549876543213216549876543216549876543216549876545432165498765432165498765432165498765432165465432165498765432165498765432131654987654321654987654321654987654321216549876543216549876543216654879654321654987654321321654987654321654987654321652165498765432165498765432164987654543216549876543216549876543216549876543216546543216549876543216549876543213165498765432165498765432165498765432165487965432165498765432132165498765432165498765432165498765454321654987654321654987654321654987654321654654321654987654321654987654321316549876543216549876543216549876543212165498765432165498765432166548796543216549876543213216549876543216549876543216521654987654321654987654321649876545432165498765432165498765432165498765432165465432165498765432165498765432131654987654321654987654321654987654321654879654321654987654321321654987654321654987654321654987654543216549876543216549876543216549876543216546543216549876543216549876543213165498765432165498765432165498765432121654987654321654987654321665487965432165498765432132165498765432165498765432165216549876543216549876543216498765454321654987654321654987654321654987654321654654321654987654321654987654321316549876543216549876543216549876543216548796543216549876543213216549876543216549876543216549876545432165498765432165498765432165498765432165465432165498765432165498765432131654987654321654987654321654987654321216549876543216549876543216654879654321654987654321"
    newStr = ""
    i=0
    
    if inputString <> "" and len(myOffsetList)>=len(inputString) then
        myOffsetList = right(myOffsetList,len(myOffsetList)-len(inputString))
        for i = 1 to len(inputString)
            newStr = newStr & chr((asc(mid(inputstring,i,1))+int(mid(myOffsetList,i,1))))
        next
    end if
    
    maxEncrypt = newStr
end function

function maxDecrypt(byval inputString)
    dim myOffsetList
    dim i
    dim oldStr
    
    '// initial
    myOffsetList = "32165498765432165498765432165216549876543216549876543216498765454321654987654321654987654321654987654321654654321654987654321654987654321316549876543216549876543216549876543216548796543216549876543213216549876543216549876543216549876545432165498765432165498765432165498765432165465432165498765432165498765432131654987654321654987654321654987654321216549876543216549876543216654879654321654987654321321654987654321654987654321652165498765432165498765432164987654543216549876543216549876543216549876543216546543216549876543216549876543213165498765432165498765432165498765432165487965432165498765432132165498765432165498765432165498765454321654987654321654987654321654987654321654654321654987654321654987654321316549876543216549876543216549876543212165498765432165498765432166548796543216549876543213216549876543216549876543216521654987654321654987654321649876545432165498765432165498765432165498765432165465432165498765432165498765432131654987654321654987654321654987654321654879654321654987654321321654987654321654987654321654987654543216549876543216549876543216549876543216546543216549876543216549876543213165498765432165498765432165498765432121654987654321654987654321665487965432165498765432132165498765432165498765432165216549876543216549876543216498765454321654987654321654987654321654987654321654654321654987654321654987654321316549876543216549876543216549876543216548796543216549876543213216549876543216549876543216549876545432165498765432165498765432165498765432165465432165498765432165498765432131654987654321654987654321654987654321216549876543216549876543216654879654321654987654321"
    oldStr = ""
    i=0
    
    if inputString <> "" and len(myOffsetList)>=len(inputString) then
        myOffsetList = right(myOffsetList,len(myOffsetList)-len(inputString))
        for i = 1 to len(inputString)
            oldStr = oldStr & chr((asc(mid(inputstring,i,1))-int(mid(myOffsetList,i,1))))
        next
    end if
    
    maxDecrypt = oldStr
end function

function maxEncrypt2(byval inputString)
    dim myPrivateKey
    dim i
    dim newStr
    
    '// initial
    myPrivateKey = 18
    newStr = ""
    i=0
    
    if inputString <> "" then
        for i = 1 to lenb(inputString)
            newStr = newStr & chrb(ascb(midb(inputstring,i,1)) xor myPrivateKey)
        next
    end if
    
    maxEncrypt2 = newStr
end function

function maxDecrypt2(byval inputString)
    dim myPrivateKey
    dim i
    dim newStr
    
    '// initial
    myPrivateKey = 18
    newStr = ""
    i=0
    
    if inputString <> "" then
        for i = 1 to lenb(inputString)
            newStr = newStr & chrb(ascb(midb(inputstring,i,1)) xor myPrivateKey)
        next
    end if
    
    maxDecrypt2 = newStr
end function
%>