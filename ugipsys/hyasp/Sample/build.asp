<html>
<title>Hysearch �����˯����� ASP ���սd��</title>
<body>
<center>
<br>
<h3>HySearch �����˯��t��  ASP �������պ���</h3><br><br>
<%
   ' step 1: �ϥ� CreateObject ���� Object. DLL name: hysdk, class name: hyft
   Set obj = Server.CreateObject("hysdk.hyft.1")

   ' step 2: �I�s hyft_connect �s�� hyftd
   x = obj.hyft_connect("localhost", 2816, "WRITER", "", "")

   ' hyft_connect() �Ǧ^�Ȭ��s�u id, �s��� obj.ret_val ��
   id = obj.ret_val
   if id = -1 then
      '�Y�O�s�u����, �h���~�T���s��� obj.errmsg
      response.write "�s�u hyftd �o�Ϳ��~, ���~�T��: " & obj.errmsg & "<br><br>"
      response.end
   else
      response.write "�s�u hyftd ���\! id = " & id & "<br><br>"
   end if

   ' step 3: ��l�Ư��ޫإ�����
   x = obj.hyft_initbuild(id)
   ' �Ǧ^�Ȭ� obj.ret_val
   ret = obj.ret_val
   if ret = -1 then
      '�Y�O��l�ƥ���, �h���~�T���s��� obj.errmsg
      response.write "initbuild �o�Ϳ��~, ���~�T��: " & obj.errmsg
      response.end
   end if

   response.write "�د��ޤ�......<br><br>"

   ' step 4: �}�l�I�s hyft_addbuild, �N�n�د��ޤ�����ƶ�J
   response.write "au = �k���J�L<br>"
   x = obj.hyft_addbuild(id, "au", "�k���J�L")
   ret = obj.ret_val
   if ret = -1 then
      response.write "��J au ���ɵo�Ϳ��~, ���~�T��: " & obj.errmsg
      response.end
   end if

   response.write "sex = �k<br>"
   x = obj.hyft_addbuild(id, "sex", "�k")
   ret = obj.ret_val
   if ret = -1 then
      response.write "��J sex ���ɵo�Ϳ��~, ���~�T��: " & obj.errmsg
      response.end
   end if

   response.write "birth = 19750413<br>"
   x = obj.hyft_addbuild(id, "birth", "19750413")
   ret = obj.ret_val
   if ret = -1 then
      response.write "��J birth ���ɵo�Ϳ��~, ���~�T��: " & obj.errmsg
      response.end
   end if
   
   response.write "text = ������<br>"
   x = obj.hyft_addbuild(id, "text", "������")
   ret = obj.ret_val
   if ret = -1 then
      response.write "��J text ���ɵo�Ϳ��~, ���~�T��: " & obj.errmsg
      response.end
   end if

   ' step 5: �I�s hyft_build �öǤJ sysid �i��د��ްʧ@
   x = obj.hyft_build(id, "000000000001")
   ret = obj.ret_val
   if ret = -1 then
      response.write "�د��ޮɵo�Ϳ��~, ���~�T��: " & obj.errmsg
      response.end
   end if

   ' step 6: �����P hyftd ���s�u
   x = obj.hyft_close(id)
   ret = obj.ret_val
   if ret = -1 then
      response.write "�����s�u�ɵo�Ϳ��~, ���~�T��: " & obj.errmsg
      response.end
   end if
   response.write "<br>�إ߯��ާ���!<br>"
   set obj=Nothing	
%>

<br><br>
<a href="javascript:history.go(-1)">�^�W��</a><br>
</center>
</body>
</html>
