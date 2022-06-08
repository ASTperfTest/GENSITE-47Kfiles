<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="myknowledge_trace_lp.aspx.vb" Inherits="myknowledge_trace_lp" title="農業知識入口網-農業知識家" ValidateRequest="false" %>

<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

	  <!--path star (style B)-->
	  <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a>&gt;<a href=""#"">我的追蹤</a></div>	
	
	  <h3>我的個人紀錄</h3>
  	
  	<!-- 我的發問TAB LIST／開始 -->
  	<div id="Magazine">
		<div id="MagTabs"><asp:Label ID="TabText" runat="server" Text=""></asp:Label></div>	
		<div class="Event">
		
		    <!--資料篩選/開始 -->
			<div>
			    <label>資料篩選：</label>
                <asp:DropDownList ID="QuestionTypeDDL" AutoPostBack="true" runat="server">                
                    <asp:ListItem Value="" Text="全部" />
                    <asp:ListItem Value="A" Text="農" />
                    <asp:ListItem Value="B" Text="林" />
                    <asp:ListItem Value="C" Text="漁" />
                    <asp:ListItem Value="D" Text="牧" />
                    <asp:ListItem Value="E" Text="其他" />             
                </asp:DropDownList>			      		        
		    </div>		    

		    <!--瀏覽排序/開始 -->
			<div class="View">
                瀏覽排序依：<asp:Label ID="ArticleTypeText" runat="server" Text=""></asp:Label>
			</div>
		    
            <!--分頁 -->
		    <div class="Page">
        	    第<asp:Label ID="PageNumberText" runat="server" Text="" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server" Text="" CssClass="Number" />頁，
        	    共<asp:Label ID="TotalRecordText" runat="server" Text="" CssClass="Number" />筆資料，     
        	    <asp:HyperLink ID="PreviousLink" runat="server">
        	        <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg" Visible="false" AlternateText="上一頁"></asp:image>        	            
        	        <asp:Label ID="PreviousText" runat="server" Visible="false" Text="Label">上一頁 &nbsp;</asp:Label>            
        	    </asp:HyperLink>
                到第 <asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" /> 頁 &nbsp;
                <asp:HyperLink ID="NextLink" runat="server">
                    <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg" Visible="false" AlternateText="下一頁"></asp:image>                     	
                    <asp:Label ID="NextText" runat="server" Visible="false" Text="Label">下一頁 &nbsp;</asp:Label>
                </asp:HyperLink>，每頁顯示                      
                <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
                    <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
                    <asp:ListItem Value="20">20</asp:ListItem>
                    <asp:ListItem Value="30">30</asp:ListItem>
                    <asp:ListItem Value="50">50</asp:ListItem>
                </asp:DropDownList> 筆
            </div>
            
            <!--表格條列 /開始 -->
		    <div class="lp">
                <asp:Label ID="TableText" runat="server" Text="" />  
                <asp:Button ID="DeleteTraceBtn" CssClass="btn2" runat="server" Text="刪除選取的項目" />                     
                <asp:TextBox ID="DeleteItem" Height="0" Width="0" BorderWidth="0"  runat="server"></asp:TextBox>
		    </div>
                
    		<!--搜尋我的知識/開始 -->
			<div class="mysearch">
			    <label for="textfield">搜尋我的追蹤：</label>
			    <asp:TextBox ID="KeywordText" runat="server" onfocus="this.value=''">請輸入關鍵字</asp:TextBox>
                <asp:Button ID="SearchBtn" CssClass="btn2" runat="server" Text="搜尋" />
		        <asp:Button ID="AdSearchBtn" CssClass="btn2" runat="server" Text="進階查詢" Visible="false" />		        		  
		    </div>
	    </div>
  	</div>
  	<div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td>	
	
	<script language="javascript" type="text/javascript">
	
	function CheckDeleteBtn(total) {
		
		var item = "";
	    for(var i = 0; i < total; i++) {
	        var checkboxname = "tracecheckbox" + i;
	        if(document.getElementById(checkboxname).checked)
	        {
	            item += document.getElementById(checkboxname).value + ",";
	        }	       
	    }
        document.aspnetForm.ctl00_ContentPlaceHolder1_DeleteItem.value = item;
	}
	
	</script>

</asp:Content>

