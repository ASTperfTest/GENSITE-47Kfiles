<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Step1_2.aspx.cs" Inherits="Step1_2" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">   
  <link href="./css/list.css" rel="stylesheet" type="text/css" />
  <link href="./css/layout.css" rel="stylesheet" type="text/css"/>
  <link href="./css/theme.css" rel="stylesheet" type="text/css"/>
  <title>主題館基本資料</title>
  <script language="javascript" type="text/javascript">
   
    function check_login() {      
      var Title = document.getElementById("Txt_Name");
      var description = document.getElementById("Txt_Description");
	    var class2 = document.form1.DDL_Class1.value;
	    var file = document.form1.Fileup.value;
	    var subtitle = document.getElementById("Txt_SubTitle");
     
	    if(Title.value == "") {	
		    alert("請輸入主題");
		    event.returnValue = false;		
	    }
	    else if(description.value == "") {
	      alert("請輸入主題館描述");
		    event.returnValue = false;
	    }	
	    else if(class2 == "--請選擇--") {
		    alert("請選擇分類");
		    event.returnValue = false;    
	    }
	    else if(Title.value.length > 20) {
	      alert("主題不得超過20字!");
	      event.returnValue = false;
	    }
	    else if(description.value.length > 250) {
	      alert("描述不得超過250字!");
	      event.returnValue = false;
	    }
	    else if(subtitle.value.length > 100) {
	      alert("副標題不得超過100字!");
	      event.returnValue = false;
	    }		
    }    
  </script>
</head>
<body>
  <div id="FuncName">
	  <h1>主題館管理</h1>
	  <div id="ClearFloat"></div>
  </div>

  <div class="step">步驟: <span>1. 填寫基本資料</span> > 2. 設定版面風格 > 3. 導覽架構設定 > 4. 首頁配置</div>

  <form id="form1" runat="server">
    <div>
    <table cellspacing="0" class="setting">
    <caption>【主題館基本資料】</caption>
    <tr>
      <th style="width: 20%"><em>*</em> 主題館名稱：</th>
      <td><asp:TextBox ID="Txt_Name" runat="server" Width="200px" CssClass="box" ></asp:TextBox></td>
    </tr>
    <tr>
      <th style="width: 489px; height: 39px;">主題館副標題：</th>
      <td style="height: 39px"><asp:TextBox ID="Txt_SubTitle" runat="server" Width="200px" CssClass="box"></asp:TextBox></td>
    </tr>
    <tr>
      <th style="width: 489px">外部主題URL連結：</th>
      <td>
        <asp:TextBox ID="Txt_URL" runat="server" Width="275px" CssClass="box"></asp:TextBox><br /> 
        (若所增的主題館屬於外部現有的網站，請於欄位中填寫連結網址，填寫完畢後按下一步，即完成新增主題館的動作)
      </td>
    </tr>
    <tr>
      <th style="width: 489px; height: 52px;"><em>*</em> 主題館圖片：</th>
      <td style="height: 52px">
        <asp:FileUpload ID="Fileup" runat="server" CssClass="box" Width="259px" /><br />
        (建議圖片大小為142×62px)<br />
        <asp:Image ID="DBimage" runat="server" Visible="False" Height="62px" />
      </td>
    </tr>
    <tr>
      <th style="width: 489px"><em>*</em> 主題館描述：</th>
      <td><textarea name="" cols="48" rows="3" class="box" id="Txt_Description" runat="server"></textarea></td>
    </tr>
    <tr>
      <th style="width: 489px; height: 31px;">是否公開：</th>
      <td style="height: 31px">
        <asp:RadioButton ID="Rb_yes" runat="server" GroupName="Yesorno" Text="是" />
        <asp:RadioButton ID="Rb_no" runat="server" GroupName="Yesorno" Text="否" />
        (系統預設不公開)
      </td>
    </tr>
    <tr>
      <th style="width: 489px; height: 52px;"><em>*</em> 主題館分類：</th>
      <td style="height: 52px">
        <asp:DropDownList ID="DDL_Class1" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DDL_Class1_SelectedIndexChanged"></asp:DropDownList>
        <asp:DropDownList ID="DDL_Class2" runat="server" Visible="true"></asp:DropDownList>
      </td>
    </tr>
    </table>
    </div>
    
    <div class="settingbutton">
      <asp:Button ID="back" Text="回首頁" runat="server" Visible="true" />
      <asp:Button ID="next" OnClick="next_Click" Text="下一步" OnClientClick="check_login()" runat="server" Visible="true" />
      <asp:Button ID="save" runat="server" Text="儲存設定" OnClientClick="check_login()" OnClick="save_Click" />
      <asp:Button ID="delete" runat="server" Text="刪除主題館" OnClientClick="return confirm('是否刪除?');" OnClick="delete_Click" Visible="true" EnableTheming="True" />
    </div>
  </form>
</body>
</html>
