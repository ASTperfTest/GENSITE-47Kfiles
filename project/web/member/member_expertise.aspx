<%@ Page Language="VB" AutoEventWireup="false" CodeFile="member_expertise.aspx.vb" Inherits="member_member_expertise" Title="領域及專長" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<link type="text/css" rel="stylesheet" href="../css/treeview.css">
<link type="text/css" rel="stylesheet" href="../css/tree.css">
<style type="text/css">
table  {
    font-size:13px;
}
table {
    border-collapse: collapse;
    border-spacing: 1;
    padding:0;
}
a {
    color:Black;
    text-decoration: none;
}
</style>
<script type="text/javascript" src="../js/jquery.js"></script>
<script type="text/javascript" src="../js/yahoo-min.js"></script>
<script type="text/javascript" src="../js/event-min.js"></script>
<script type="text/javascript" src="../js/treeview-min.js"></script>
<script type="text/javascript" src="../js/tasknode.js"></script>
<script type="text/javascript" src="../js/lambda-utilities.js"></script>
    <title></title>
</head>
<script type="text/javascript">
    var tree,currentIconMode;
    function changeIconMode() {
        var newVal = parseInt(this.value);
        if (newVal != currentIconMode) {
            currentIconMode = newVal;
        }
        treeInit();
    }
    function treeInit() {
        tree = new YAHOO.widget.TreeView("treeContainer");
        tree.setDynamicLoad(loadNodeData, currentIconMode);
        var root = tree.getRoot();
        var myobj, tmpNode;
        myobj = { label: "領域及專長", id: "2" };
        tmpNode = new YAHOO.widget.TextNode(myobj, root, false);
        tmpNode.labelStyle = "icon-root";
        tree.subscribe("checkClick", onCheckClick);
        tree.draw();
        $("#ygtvt1").click();
        //loadNodeData(tmpNode, function () { tmpNode.loadComplete(); });
    }
    function loadNodeData(node, fnLoadComplete) {
            if (node.data.source == null) {
                var url = "../services/categoryservices.aspx";
                var param = "cmd=expandtreenode&nodeid=2";
                $.get(url, param, function (data) {
                    //alert(data);
                    var json = eval('(' + data + ')');
                    if (json.Success)
                    $.each(json.Data, function (i, n) {
                        newobj = { label: n.DisplayName, id: n.Id, source: n };
                        var newNode = new YAHOO.widget.TaskNode(newobj, node, false);
                    });
                    fnLoadComplete();
                });
        }
        else {
            var url = "../services/categoryservices.aspx";
            var param = "cmd=expandtreenode&nodeid=" + node.data.source.Id;
            $.get(url, param, function (data) {
                var json = eval('(' + data + ')');
                if (json.Success)
                $.each(json.Data, function (i, n) {
                    newobj = { label: n.DisplayName, id: n.Id, source: n };
                    var newNode = new YAHOO.widget.TaskNode(newobj, node, false);
                });
                fnLoadComplete();
            });
        }
    }
    function onCheckClick(node) {
        text = $("#hidevalue").val();
        if (text.indexOf(node.data.source.DisplayName) != -1) {
            text=text.replace(node.data.source.DisplayName + ",", "");
        } else {
            text += node.data.source.DisplayName + ",";
        }
        $("#hidevalue").val(text);
    }
    YAHOO.util.Event.on(window, "load", treeInit);
    function SendMessage() {
        var obj = window.dialogArguments;
        temp = $("#hidevalue").val();
        if (temp != "") {
            if (temp.substring(temp.length - 1, temp.length) == ",") {
                temp = temp.substring(0, temp.length - 1)
            }
            obj.Express = temp;
            window.returnValue = "Y";
        }
        window.close();
    }
</script>
<body>
    <form id="form1" runat="server">
    <div>
    <div id="treeContainer" style="margin:5px;">
    </div>
    </div>
    <br/>
    &nbsp;&nbsp;<input type="button" ud="state_close" value="關閉" onclick="self.close();" />
     &nbsp;&nbsp;<input type="button" id="state_sumbit" value="確定" onclick="SendMessage();"/>
     <input type="hidden" id="hidevalue" value=""  />
    </form>
</body>
</html>
