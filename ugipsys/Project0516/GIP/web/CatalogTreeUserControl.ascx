<%@ Control Language="C#" AutoEventWireup="true" CodeFile="CatalogTreeUserControl.ascx.cs"
	Inherits="GIP_web_CatalogTreeUserControl" %>
<h5>
	目錄架構 》</h5>
<dl>
	<asp:Repeater ID="CatelogRepeater" runat="server" OnItemDataBound="CatelogRepeater_ItemDataBound"
		OnItemCommand="CatelogRepeater_ItemCommand">
		<ItemTemplate>
			<dt id="Dt1" runat="server">
				<asp:Button ID="MoveUpButton" runat="server" CssClass="button right" Text="↑" CausesValidation="false" />
				<asp:Button ID="MoveDownButton" runat="server" CssClass="button right" Text="↓" CausesValidation="false" />
				<asp:LinkButton ID="EditButton" runat="server" CausesValidation="false"></asp:LinkButton>
			</dt>
			<asp:Repeater ID="NodeRepeater" runat="server" OnItemCommand="CatelogRepeater_ItemCommand">
				<ItemTemplate>
					<dd id="Dd1" runat="server">
						<asp:Button ID="MoveUpButton" runat="server" CssClass="button right" Text="↑" CausesValidation="false" />
						<asp:Button ID="MoveDownButton" runat="server" CssClass="button right" Text="↓" CausesValidation="false" />
						<asp:LinkButton ID="EditButton" runat="server" CausesValidation="false"></asp:LinkButton>
					</dd>
				</ItemTemplate>
			</asp:Repeater>
		</ItemTemplate>
	</asp:Repeater>
</dl>