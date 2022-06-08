<%@ Control Language="C#" AutoEventWireup="true" CodeFile="GardeningExpert.ascx.cs" Inherits="GardeningExpert" %>

<asp:DataList ID="gardenExpertDataList" runat="server" 
    RepeatColumns="3" 
    RepeatDirection="Horizontal">
    <ItemTemplate>
        <table style="float:left;">
            <tr>
                <td style="height:90px;">
                    <asp:Image ID="gardenExprtyImage" runat="server" ImageUrl='<%# Eval("Image") %>' Width="80px" />
                </td>
            </tr>
            <tr>
                <td style="width:100px;">
                    <asp:Label ID="gardebExpertName" runat="server" Text='<%# Eval("Name") %>' ></asp:Label>
                </td>
            </tr>
            <tr>
                <td style="height:200px;width:100px;">
                    <asp:Label ID="gardenExpertIntro" runat="server" Text='<%# Eval("Intro") %>'  Font-Size="Smaller"></asp:Label>
                </td>
            </tr>
        </table>		
    </ItemTemplate>
	
    <HeaderTemplate>
        <table>
            <tr>
                <td style="width:100%;">
                        <asp:Label ID="lbHeaderTitle" runat="server" Text="園藝達人" ></asp:Label>                    
                </td>
                <td>
                        <asp:Label ID="lbMore" runat="server" Text="More..."></asp:Label>                    
                </td>
            </tr>
        </table>
    </HeaderTemplate>	
	
</asp:DataList>

