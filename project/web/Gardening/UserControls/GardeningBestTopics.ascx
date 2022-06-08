<%@ Control Language="C#" AutoEventWireup="true" CodeFile="GardeningBestTopics.ascx.cs" Inherits="GardeningBestTopics" %>
<link href="~/Gardening/css/index.css" rel="stylesheet" type="text/css" />

<table class="title2" cellpadding="0" cellspacing="0">
            <tr>
                <td style=" width:155px;">
					<span><asp:Label runat="server" Text="園藝栽培" CssClass="TitleStyle1"></asp:Label> </span>  
                    <span><asp:Label runat="server" Text="主題館" ></asp:Label> </span>                   
                </td>
            </tr>
			<tr style="heigth:10px;">
				<td class="line01">&nbsp;</td>
				<td class="line02">&nbsp;</td>
			</tr>
</table>

<div id="table_best_topics">
<asp:DataList ID="BestTopicsDataList" runat="server" Height="197px" 
    Width="150px">
    
    <ItemTemplate>
    
          <table runat="server" >
            <tr>
                <td id="topic_img_td" rowspan="2" width="100px">
                    <asp:Image runat="server" ImageUrl='<%# Eval("ImageUri") %>' Width="100px" CssClass="topic_img"/>
                </td>
                <td>
                    <asp:HyperLink ID="hlkTitle" runat="server" NavigateUrl='<%# Eval("TopicUri") %>' CssClass="topic_title"><%# Eval("Title") %></asp:HyperLink>              
                </td>
            </tr>
            <tr>
                <td class="topic_context">
                    <asp:Label ID="lbDescription" runat="server" Text='<%# Eval("Description") %>' ></asp:Label>                
                </td>
            </tr>
         </table>        
    </ItemTemplate>
    
</asp:DataList>
</div>