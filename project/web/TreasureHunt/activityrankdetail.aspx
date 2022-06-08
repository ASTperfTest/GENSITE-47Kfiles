<%@ Page Language="C#" AutoEventWireup="true" CodeFile="activityrankdetail.aspx.cs" Inherits="TreasureHunt_activityrankdetail" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../js/jquery-1.3.2.min.js"></script>

    <script type="text/javascript" src="../js/datepicker.js"></script>

    <link rel="Stylesheet" href="../css/datepicker.css" />
    <link type="text/css" href="../css/jquery.css" rel="stylesheet" />
    <script type="text/javascript" language="JavaScript">
 
    
    $(document).ready(function(e){
		$("#TxtStartDate").datepicker({
        dateFormat:"yy/mm/dd"
        }); 
        $("#TxtEndDate").datepicker({
        dateFormat:"yy/mm/dd"
        });
        $("#ReverseChecked").click(function(){
            if($("#ReverseChecked").attr("checked"))
            if($("#TreasureCategory").val()==0){
                $('#TreasureCategory')[0].selectedIndex =1;
            }
             $("#OnlyGetBox").attr("checked",false);
        });
        $("#OnlyGetBox").click(function(){
            $("#ReverseChecked").attr("checked",false);
            $('#TreasureCategory')[0].selectedIndex =0;
        });
    })
    
    function QueryTop(){
        $("#FroQuery").val("true");
        document.forms['form1'].submit();
    }
    
    function NextPage(){
        index = $('#PageNumberDDL').get(0).selectedIndex;
        $('#PageNumberDDL').get(0).selectedIndex = index +1;
        document.forms['form1'].submit();
    }
    function PrePage(){
        index = $('#PageNumberDDL').get(0).selectedIndex;
        $('#PageNumberDDL').get(0).selectedIndex = index -1;
        document.forms['form1'].submit();
    }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
             <td class="content">
             <div class="content_mid">
        <div class="Page">
            <br/>
            已發出寶物數量：<br/>
            <asp:Label ID="AllFindedTreasure" runat="server"></asp:Label>
            <br/><br/>
            所有使用者擁有寶物數：<br/>
            <asp:Label ID="AllUserTreasure" runat="server"></asp:Label>
            <div >
			<br/>
		   篩選條件：<br/>
		   
		   <table><tr>
		                <td>會員:
		                </td>
		                <td><asp:TextBox ID="TextBoxMember" runat="server" Width="100px"></asp:TextBox>&nbsp;(帳號、姓名、暱稱)
		                </td>
		          </tr>
		          <tr>
		                <td>已蒐集寶物種類:
		                </td>
		                <td><asp:DropDownList ID="TreasureCategory" AutoPostBack="false" runat="server" /> &nbsp;<asp:CheckBox ID="ReverseChecked" runat="server" Text="未蒐集"></asp:CheckBox>
		                </td>
		          </tr>
		          <tr>
		                <td>已蒐集寶物套數:
		                </td>
		                <td><asp:TextBox ID="TextBoxSuits" runat="server" Width="30px"></asp:TextBox>~<asp:TextBox ID="TextBoxSuitsUpper" runat="server" Width="30px"></asp:TextBox>
		                 <asp:RegularExpressionValidator id="RegularExpressionValidator3" ValidationExpression="\d*" runat="server" ControlToValidate="TextBoxSuits" ErrorMessage="已蒐集寶物套數區間只能輸入數字" Display="None" SetFocusOnError="true" />
		                <asp:RegularExpressionValidator id="validatorTextBoxScoreS" ValidationExpression="\d*" runat="server" ControlToValidate="TextBoxSuitsUpper" ErrorMessage="已蒐集寶物套數區間只能輸入數字" Display="None" SetFocusOnError="true" />
		                </td>
		          </tr>
		          <tr>
		                <td>獲得寶物數:
		                </td>
		                <td><asp:TextBox ID="TextTreasureSetsLowerBound" runat="server" Width="30px"></asp:TextBox>~<asp:TextBox ID="TextTreasureSetsUpperBound" runat="server" Width="30px"></asp:TextBox>
		                <asp:RegularExpressionValidator id="RegularExpressionValidator1" ValidationExpression="\d*" runat="server" ControlToValidate="TextTreasureSetsLowerBound" ErrorMessage="獲得寶物數只能輸入數字" Display="None" SetFocusOnError="true"/>
		                <asp:RegularExpressionValidator id="RegularExpressionValidator2" ValidationExpression="\d*" runat="server" ControlToValidate="TextTreasureSetsUpperBound" ErrorMessage="獲得寶物數只能輸入數字" Display="None" SetFocusOnError="true"/>
		                </td>
		          </tr>
		          <tr>
		                <td>尋寶機率小於:
		                </td>
		                <td><asp:TextBox ID="TextBoxProbability" runat="server" Width="30px"></asp:TextBox>/100<asp:RegularExpressionValidator id="validatorTextBoxScoreE" ValidationExpression="\d*" runat="server" ControlToValidate="TextBoxProbability" ErrorMessage="尋寶機率只能輸入數字" Display="None" SetFocusOnError="true"/>
		                </td>
		          </tr>
		          <tr>
		                <td>日期區間:
		                </td>
		                <td>
		                <asp:TextBox ID="TxtStartDate" runat="server" Width="80px" ></asp:TextBox>~<asp:TextBox ID="TxtEndDate" runat="server" Width="80px"></asp:TextBox>
		                
		                </td>
		          </tr>
		          <tr>
		                <td>寶物領取查詢:
		                </td>
		                <td>
		                <asp:CheckBox ID="OnlyGetBox" runat="server" Text="未領取寶物"></asp:CheckBox>
		                
		                </td>
		          </tr>
		          <tr>
		                <td colspan="2"><input type="button" id="btnQuery" value="查詢" onclick="QueryTop();"/>&nbsp;&nbsp;<input type="button" id="clearAll" value="回首頁" onclick="window.location='/treasurehunt/activityrankdetail.aspx?avtivityid=<%=avtivityid %>'"/>
		                </td>
		          </tr>
		   </table>
                
                <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="true" ShowSummary="false"/>
                <br/>
    		  
                
                <asp:HyperLink ID="linkExport"  runat="server" Target="_blank">匯出報表</asp:HyperLink> &nbsp;&nbsp;&nbsp;  <asp:HyperLink ID="HyperLink1" Visible="false"  runat="server" Target="_self" NavigateUrl="checkcheatuser.aspx" >驗證錯誤清單</asp:HyperLink>   
		   </div>
        第<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>頁，
   	    共<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />筆資料，     
        <asp:HyperLink ID="PreviousLink" runat="server">
          <asp:image runat="server" ImageUrl="/xslGip/style1/images3/arrow_left.gif" ID="PreviousImg" AlternateText="上一頁" />
          <asp:Label ID="PreviousText" runat="server" >上一頁 &nbsp;</asp:Label>
        </asp:HyperLink>
        跳至第<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" /> 頁 &nbsp;
        <asp:HyperLink ID="NextLink" runat="server">
          <asp:image runat="server" ImageUrl="/xslGip/style1/images3/arrow_right.gif" ID="NextImg" AlternateText="下一頁"/>
          <asp:Label ID="NextText" runat="server">下一頁 &nbsp;</asp:Label>
        </asp:HyperLink>，每頁                      
        <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
          <asp:ListItem Selected="True" Value="15">15</asp:ListItem>                    
          <asp:ListItem Value="30">30</asp:ListItem>
          <asp:ListItem Value="50">50</asp:ListItem>
        </asp:DropDownList>筆資料
      </div>
      <asp:Label ID="TableText" runat="server" Text="" />  
      <div class="top">
        <a href="#" title="top">top</a>
      </div>
      <br/>
      </div>
      </td>
      </tr>
      </table>
    </div>
    <input type="hidden" value="false" id="FroQuery" name="FroQuery" />
    <input type="hidden" value="<%=avtivityid %>" id="avtivityid" name="avtivityid" />
</form>
</body>
</html>
