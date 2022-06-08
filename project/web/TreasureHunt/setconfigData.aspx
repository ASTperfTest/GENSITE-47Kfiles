<%@ Page Language="C#" AutoEventWireup="true" CodeFile="setconfigData.aspx.cs" Inherits="TreasureHunt_setconfigData" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>未命名頁面</title>
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../js/jquery-1.3.2.min.js"></script>
    <script type="text/javascript" src="../js/datepicker.js"></script>
     <link rel="Stylesheet" href="../css/datepicker.css" />
    <link type="text/css" href="../css/jquery.css" rel="stylesheet" />
    <script type="text/javascript" language="JavaScript">
    $(document).ready(function(e){
		    $("#TextBoxChatModeStartTime").datepicker({
            dateFormat:"yy/mm/dd"
            }); 
            $("#TextBoxChatModeEndTime").datepicker({
            dateFormat:"yy/mm/dd"
            }); 
            $("#TextVoteStartDate").datepicker({
            dateFormat:"yy/mm/dd"
            }); 
            $("#TextVoteEndDate").datepicker({
            dateFormat:"yy/mm/dd"
            }); 
            $("#savesubmit").click(function(){
                SaveData();
            });
        }); 
        function SaveData(){
            $("#saveconfigData").val("true");
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
        <br/>
        活動名稱:<asp:DropDownList ID="activityList" AutoPostBack="true" runat="server" />
        <br/>
        <br/>
        <br/>
        <table>
            <tr>
                <td>
                    備援程式起始日:
                </td>
                <td>
                    <asp:TextBox ID="TextBoxChatModeStartTime" runat="server" Width="80px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    備援程式結束日:
                </td>
                <td>
                    <asp:TextBox ID="TextBoxChatModeEndTime" runat="server" Width="80px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    投套數起始日:
                </td>
                <td>
                    <asp:TextBox ID="TextVoteStartDate" runat="server" Width="80px" ></asp:TextBox>&nbsp;&nbsp;<asp:DropDownList ID="TextBoxTxtVoteStartHours" AutoPostBack="false" runat="server" />  
                </td>
            </tr>
            <tr>
                <td>
                    投套數結束日:
                </td>
                <td>
                    <asp:TextBox ID="TextVoteEndDate" runat="server" Width="80px"></asp:TextBox>&nbsp;&nbsp;<asp:DropDownList ID="TextBoxTxtVoteEndHours" AutoPostBack="false" runat="server" />
                </td>
            </tr>
            
        </table>
        <div style="float:left; width:200px;">&nbsp;</div>
            <input type="button" id="savesubmit" name ="savesubmit" value="儲存"  />
                <input type="hidden" id="saveconfigData" name ="saveconfigData" value="false" />
        </td>
        </tr>
      </table>
    </div>
    
    </form>
</body>
</html>
