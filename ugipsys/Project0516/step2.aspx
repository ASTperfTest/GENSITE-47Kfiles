<%@ Page Language="C#" AutoEventWireup="true" CodeFile="step2.aspx.cs" Inherits="step2" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">

<title>���R�W����</title>
<link href="./css/list.css" rel="stylesheet" type="text/css"/>
<link href="./css/layout.css" rel="stylesheet" type="text/css"/>
<link href="./css/theme.css" rel="stylesheet" type="text/css"/>
</head>
<body>

    <div id="FuncName">
	    <h1>�D�D�]�޲z</h1>
	    <div id="ClearFloat"></div>
    </div>

    <div class="step">�B�J�G 1. ��g�򥻸�� > <span>2. �]�w��������</span> > 3. �����[�c�]�w > 4. �����t�m</div>

        <form id="form1" runat="server">
        

        <table cellspacing="0" class="setting" >
    <caption>�i�п�����������t�m�j�]�w�]���Ѥ��Ӱ϶��A���ذt�m�覡�^
    </caption>
    <tr>
    <td>
      <asp:RadioButtonList ID="outlet" runat="server" AutoPostBack="false" RepeatColumns="4" RepeatDirection="Horizontal">
          </asp:RadioButtonList>
     </td>     
      

    </tr>
    </table>
            &nbsp;
        <table cellspacing="0" class="setting">
    <caption>�i�п���D�D�]����t��j�]�C�ӭ���N�|���ѤT�i���D�Ϥ��^
    </caption>

    <tr>
    <td>
      <asp:RadioButtonList ID="wstyle" runat="server" RepeatDirection="Horizontal" RepeatColumns="4">
        </asp:RadioButtonList></td></tr>
    </table>

    <table cellspacing="0" class="setting">
    <caption>�i�����Ϥ��]�w�j</caption>
    <tr>
    <td rowspan="2" width="30%"><img src="./images/snap1842.jpg" width="170"></td>
    <td>1. �W�Ǽ��D�Ϥ��G<br>
        <asp:FileUpload ID="Topic_upload" runat="server" Width="269px" /><br>(��ĳ�Ϥ��j�p��200 �� 35px)<br><img src="./images/snap1841.jpg" width="250"></td>
    </tr>
    <tr>
    <td style="height: 117px">2. �����T�Ϥ��G<br>(��ĳ�Ϥ��j�p��700 �� 160px, �i����w�]��T�αN�ۦ�]�p����T�Ϥ��W�ǫ����ϥ�) 
        <asp:RadioButtonList ID="bannerlist" runat="server">
        </asp:RadioButtonList><%--<p><input name="" type="radio" value=""> �A����T<br><img src="./images/snap1757.jpg" width="250"></p>
    <p><input name="" type="radio" value=""> ������T<br><img src="./images/snap1757.jpg" width="250"></p>
    <p><input name="" type="radio" value=""> �b���~��T<br><img src="./images/snap1757.jpg" width="250"></p>
    <p><input name="" type="radio" value=""> �ۭq�Ϥ���T�@/ <a href="#">�i�R���j</a><br><img src="./images/snap1757.jpg" width="250"></p>
    --%><p>>> <a href="./step2.aspx" onclick="window.open('new_web_pic.aspx','','width=450,height=200')">�i�ϥΦۦ�]�p���D�j </a></p>
    </td>
    </tr>
    </table>
    <div class="settingbutton">
    <asp:Button ID="back" runat="server" Text="�W�@�B" OnClick="back_Click" />
    <asp:Button ID="next" runat="server" Text="�U�@�B" OnClick="next_Click" />

    <asp:Button ID="save" runat="server" Text="�x�s�]�w" OnClick="save_Click" />
    <asp:Button ID="delete" runat="server" Visible="false" Text="�R���D�D�]" OnClientClick="return confirm('�O�_�R��?')" OnClick="delete_Click"  />
    <asp:Button ID="cancel" runat="server" Visible="true" Text="����(�^����)" OnClick="cancel_Click" />
    <%--<input name="delete" type="submit" onClick=" " value="�x�s�]�w">
    <input name="button" type="button" onClick="location.href='new_web_list.htm'" value="���� (�^����)">
    --%></div>
   
    </form>
</body>
</html>
