<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Step1_Edit.aspx.cs" Inherits="Edit_Step1_Edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
<link href="../css/list.css" rel="stylesheet" type="text/css" />
<link href="../css/layout.css" rel="stylesheet" type="text/css"/>
<link href="../css/theme.css" rel="stylesheet" type="text/css"/>
    <title>主題館基本資料</title>
    <script language="javascript" type="text/javascript">
   
    function check_login()
{
    
     var Title = document.getElementById("Txt_Name");
     var description = document.getElementById("Txt_Description");
	 var class2 = document.form1.DDL_Class1.value;
	 var file = document.form1.Fileup.value;
	 var subtitle = document.getElementById("Txt_SubTitle");
     var D_URL = document.getElementById("Txt_dept_URL");
	 var dept = document.getElementById("Txt_dept");
	 
	if(Title.value == "")
	{	
		alert("請輸入主題");
		event.returnValue = false;
		
	}
	else if(description.value == "")
	{
	    alert("請輸入主題館描述");
		event.returnValue = false;
	}	
	else if(class2 == "--請選擇--")
	{
		 alert("請選擇分類");
		 event.returnValue = false;    
	}
	//else if(dept.value == "")
	//{	
		//alert("請輸入維護單位");
		//event.returnValue = false;
		
	//}
	//else if(D_URL.value == "")
	//{	
		//alert("請輸入維護單位URL連結");
		//event.returnValue = false;
		
	//}
	else if(Title.value.length > 20)
	{
	     alert("主題不得超過20字!");
	     event.returnValue = false;
	}
	else if(description.value.length > 250)
	{
	     alert("描述不得超過250字!");
	     event.returnValue = false;
	}
	else if(subtitle.value.length > 100)
	{
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

    <div class="step">步驟: <span>1. 填寫基本資料</span> > 2. <a href="step2_edit.aspx">設定版面風格</a> > 3. <a href="../GIP/web/step3.aspx">導覽架構設定</a> > 4. <a href="../HomePageEdit.aspx">首頁配置</a></div>


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
        <td><asp:TextBox ID="Txt_URL" runat="server" Width="275px" CssClass="box"></asp:TextBox><br>
        <font color="red">(若所新增的主題館屬於外部現有的網站，請於欄位中填寫連結網址，填寫完畢後按下一步，即完成新增主題館的動作)</font>
        </td>
    </tr>
    <tr>
        <th style="width: 489px; height: 52px;"><em>*</em> 主題館圖片：</th>
        <td style="height: 52px">
            <asp:FileUpload ID="Fileup" runat="server" CssClass="box" Width="259px" /><br />
        (建議圖片大小為142×62px)<br />
            <asp:Image ID="DBimage" runat="server" Visible="False" Height="62px" /></td>
    </tr>
    <tr>
        <th style="width: 489px"><em>*</em> 主題館描述：</th>
        <td><textarea name="" cols="48" rows="3" class="box" id="Txt_Description" runat="server"></textarea></td>
    </tr>
    <tr>
    <th style="width: 489px; height: 31px;">是否公開：</th>
    <td style="height: 31px">
        <asp:RadioButton ID="Rb_yes" runat="server" GroupName="Yesorno" Text="是" />
        <asp:RadioButton ID="Rb_no" runat="server" GroupName="Yesorno"
            Text="否" />
     (系統預設不公開)</td>
    </tr>
    <tr>
    <th style="width: 489px; height: 52px;"><em>*</em> 主題館分類：</th>
    <td style="height: 52px">
        <asp:DropDownList ID="DDL_Class1" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DDL_Class1_SelectedIndexChanged">
        </asp:DropDownList>
        <asp:DropDownList ID="DDL_Class2" runat="server" Visible="true">
        </asp:DropDownList>
        </td>
    </tr>
    <tr>
    <th style="width: 489px; height: 52px;"><em>*</em> 是否鎖右鍵：</th>
    <td style="height: 52px">
        <asp:DropDownList ID="lockRightBtnDDL" runat="server">
          <asp:ListItem Value="Y" Text="是" />
          <asp:ListItem Value="N" Text="否" />
        </asp:DropDownList>
        
        </td>
    </tr>
	<tr>
        <th style="width: 489px"><em>*</em> 維護單位：</th>
        <td><textarea name="" cols="48" rows="3" class="box" id="Txt_dept" runat="server"></textarea></td>
    </tr>
	<tr>
        <th style="width: 489px"><em>*</em>維護單位URL連結：</th>
        <td><asp:TextBox ID="Txt_dept_URL" runat="server" Width="275px" CssClass="box"></asp:TextBox><br>
        </td>
    </tr>
	<tr>
        <th style="width: 489px">關鍵字：</th>
        <td>
		    <%=kmKeyword%>
		    <br><asp:TextBox ID="Txt_Keywords" runat="server" Width="275px" CssClass="box"></asp:TextBox>
		    &nbsp;&nbsp;(使用,分隔不同關鍵字)
        </td>
    </tr>
	<tr id="orderv" style="display:block">
        <th style="width: 489px">主題館排序：</th>
        <td><asp:TextBox ID="Txt_order" runat="server" Width="50px" CssClass="box"></asp:TextBox>(只能填數字)<br>
        </td>
    </tr>
	<tr>
        <th style="width: 489px">關閉網站導覽：</th>
        <td><asp:CheckBox ID="CatMemo_Disable" runat="server" Text="(勾選後網站導覽將不顯示)" /><br>
        </td>
    </tr>
    </table>
    </div>
    
    <div class="settingbutton">
<asp:Button ID="back" Text="回首頁" runat="server" Visible="true" OnClick="back_Click" />
<asp:Button ID="next" OnClick="next_Click" Text="下一步" OnClientClick="check_login()" runat="server" Visible="true" />
<asp:Button ID="save" runat="server" Text="暫存主題館" OnClientClick="check_login()" OnClick="save_Click" />
<asp:Button ID="delete" runat="server" Text="刪除主題館" OnClientClick="return confirm('是否刪除?');" OnClick="delete_Click" Visible="true" EnableTheming="True" />
<!--<input name="delete" type="submit" onClick=" " value="儲存設定">
<input name="button" type="button" onClick="location.href='new_web_list.htm'" value="取消 (回首頁)">-->
</div>
    </form>
</body>
</html>
