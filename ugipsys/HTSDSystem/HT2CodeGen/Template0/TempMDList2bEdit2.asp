  
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub

   sub formDelSubmit()
         chky=msgbox("�`�N�I"& vbcrlf & vbcrlf &"�@�A�T�w�R����ƶܡH"& vbcrlf , 48+1, "�Ъ`�N�I�I")
        if chky=vbok then
	       reg.submitTask.value = "DELETE"
	      reg.Submit
       end If
  end sub

     sub clientInitForm
