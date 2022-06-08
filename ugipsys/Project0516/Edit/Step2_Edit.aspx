<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Step2_Edit.aspx.cs" Inherits="Edit_Step2_Edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
<link href="../css/list.css" rel="stylesheet" type="text/css"/>
<link href="../css/layout.css" rel="stylesheet" type="text/css"/>
<link href="../css/theme.css" rel="stylesheet" type="text/css"/>
</head>
<body>

    <div id="FuncName">
	    <h1>主題館管理</h1>
	    <div id="ClearFloat"></div>
    </div>

    <div class="step">步驟： 1. <a href="step1_edit.aspx">填寫基本資料</a> > <span>2. 設定版面風格</span> > 3. <a href="../GIP/web/step3.aspx">導覽架構設定</a> > 4. <a href="../HomePageEdit.aspx">首頁配置</a></div>

        <form id="form1" runat="server">
        

        <table cellspacing="0" class="setting" >
    <caption>【請選取首頁版面配置】（預設提供五個區塊，六種配置方式）
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
    <caption>【請選取主題館風格配色】（每個風格將會提供三張標題圖片）
    </caption>

    <tr>
    <td>
      <asp:RadioButtonList ID="wstyle" runat="server" RepeatDirection="Horizontal" RepeatColumns="4">
        </asp:RadioButtonList></td></tr>
    </table>

    <table cellspacing="0" class="setting">
    <caption>【網站圖片設定】</caption>
    <tr>
    <td rowspan="2" width="30%"><img src="../images/snap1842.jpg" width="170"></td>
    <td>1. 上傳標題圖片：<br>
        <asp:FileUpload ID="Topic_upload" runat="server" Width="269px" /><br />(建議圖片大小為200 × 35px)<br/><asp:Image ID="Topic_pic" ImageUrl="../images/snap1841.jpg" Width="250" Height="40px" runat="server" /></td>
    </tr>
    <tr>
    <td style="height: 117px">2. 選取橫幅圖片：<br>(建議圖片大小為700 × 160px, 可選取預設橫幅或將自行設計的橫幅圖片上傳後選取使用)
        <asp:RadioButtonList ID="bannerlist" runat="server">
        </asp:RadioButtonList><%--<p><input name="" type="radio" value=""> 農產橫幅<br><img src="./images/snap1757.jpg" width="250"></p>
    <p><input name="" type="radio" value=""> 漁產橫幅<br><img src="./images/snap1757.jpg" width="250"></p>
    <p><input name="" type="radio" value=""> 畜產品橫幅<br><img src="./images/snap1757.jpg" width="250"></p>
    <p><input name="" type="radio" value=""> 自訂圖片橫幅　/ <a href="#">【刪除】</a><br><img src="./images/snap1757.jpg" width="250"></p>
    --%><p>>> <a href="./step2_Edit.aspx" onclick="window.open('../new_web_pic.aspx','','width=450,height=200')">【使用自行設計標題】 </a></p>
    </td>
    </tr>
    </table>
    <div class="settingbutton">
    <asp:Button ID="back" runat="server" Text="上一步" OnClick="back_Click" />
    <asp:Button ID="next" runat="server" Text="下一步" OnClick="next_Click" />

    <asp:Button ID="save" runat="server" Text="儲存設定" OnClick="save_Click" />
    <asp:Button ID="delete" runat="server" Visible="false" Text="刪除主題館" OnClick="delete_Click" OnClientClick="return confirm('是否刪除?');" />
    <asp:Button ID="cancel" runat="server" Visible="true" Text="取消(回首頁)" OnClick="cancel_Click" />
    <%--<input name="delete" type="submit" onClick=" " value="儲存設定">
    <input name="button" type="button" onClick="location.href='new_web_list.htm'" value="取消 (回首頁)">
    --%></div>
    </form>
</body>
</html>
