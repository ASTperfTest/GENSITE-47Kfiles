
YAHOO.widget.TaskNode = function(oData, oParent, expanded, checked) {
    this.oData = oData;
    
    if (YAHOO.widget.LogWriter) {
        this.logger = new YAHOO.widget.LogWriter(this.toString());
    } else {
        this.logger = YAHOO;
    }

    if (oData) { 
        this.init(oData, oParent, expanded);
        this.setUpLabel(oData);
        this.setUpCheck(checked);
    }

};

YAHOO.extend(YAHOO.widget.TaskNode, YAHOO.widget.TextNode, {
    checked: false,
    checkState: 0,

    taskNodeParentChange: function() {
        //this.updateParent();
    },

    setUpCheck: function(checked) {
        if (checked && checked === true) {
            this.check();
        } else if (this.parent && 2 == this.parent.checkState) {
            this.updateParent();
        }
        if (this.tree && !this.tree.hasEvent("checkClick")) {
            this.tree.createEvent("checkClick", this.tree);
        }

        this.subscribe("parentChange", this.taskNodeParentChange);

    },

    getCheckElId: function() {
        return "ygtvcheck" + this.index;
    },

    getCheckEl: function() {
        return document.getElementById(this.getCheckElId());
    },
    getCheckStyle: function() {
        if (this.oData.source == undefined) {
            return "ygtvcheck" + this.checkState;
        }
        else {
            return "ygtvcheck" + this.checkState + " " + this.oData.source.SubjectId;
        }
    },
    getCheckLink: function() {
        return "YAHOO.widget.TreeView.getNode(\'" + this.tree.id + "\'," +
            this.index + ").checkClick()";
    },

    checkClick: function() {
        this.logger.log("previous checkstate: " + this.checkState);
        if (this.checkState === 0) {
            this.check();
        } else {
            this.uncheck();
        }

        this.onCheckClick();
        this.tree.fireEvent("checkClick", this);
    },

    onCheckClick: function() {
        this.logger.log("onCheckClick: " + this);
    },

    updateParent: function() {
        return;
        var p = this.parent;

        if (!p || !p.updateParent) {
            this.logger.log("Abort udpate parent: " + this.index);
            return;
        }

        var somethingChecked = false;
        var somethingNotChecked = false;

        for (var i = 0; i < p.children.length; ++i) {
            if (p.children[i].checked) {
                somethingChecked = true;
                // checkState will be 1 if the child node has unchecked children
                if (p.children[i].checkState == 1) {
                    somethingNotChecked = true;
                }
            } else {
                somethingNotChecked = true;
            }
        }

        if (somethingChecked) {
            p.setCheckState((somethingNotChecked) ? 1 : 2);
        } else {
            p.setCheckState(0);
        }

        p.updateCheckHtml();
        p.updateParent();
    },

    updateCheckHtml: function() {
        if (this.parent && this.parent.childrenRendered) {
            this.getCheckEl().className = this.getCheckStyle();
        }
    },
    setCheckState: function(state) {
        this.checkState = state;
        this.checked = (state > 0);
    },

    check: function() {
        this.logger.log("check");
        this.setCheckState(2);
        for (var i = 0; i < this.children.length; ++i) {
            //this.children[i].check();
        }
        this.updateCheckHtml();
        this.updateParent();
    },

    uncheck: function() {
        this.setCheckState(0);
        for (var i = 0; i < this.children.length; ++i) {
            //this.children[i].uncheck();
        }
        this.updateCheckHtml();
        this.updateParent();
    },

    getNodeHtml: function() {
        this.logger.log("Generating html");
        var sb = [];

        var getNode = 'YAHOO.widget.TreeView.getNode(\'' +
                        this.tree.id + '\',' + this.index + ')';


        sb[sb.length] = '<table border="0" cellpadding="0" cellspacing="0">';
        sb[sb.length] = '<tr>';

        for (var i = 0; i < this.depth; ++i) {
            sb[sb.length] = '<td class="' + this.getDepthStyle(i) + '"><div class="ygtvspacer"></div></td>';
        }

        sb[sb.length] = '<td';
        sb[sb.length] = ' id="' + this.getToggleElId() + '"';
        sb[sb.length] = ' class="' + this.getStyle() + '"';
        if (this.hasChildren(true)) {
            sb[sb.length] = ' onmouseover="this.className=';
            sb[sb.length] = 'YAHOO.widget.TreeView.getNode(\'';
            sb[sb.length] = this.tree.id + '\',' + this.index + ').getHoverStyle()"';
            sb[sb.length] = ' onmouseout="this.className=';
            sb[sb.length] = 'YAHOO.widget.TreeView.getNode(\'';
            sb[sb.length] = this.tree.id + '\',' + this.index + ').getStyle()"';
        }
        sb[sb.length] = ' onclick="javascript:' + this.getToggleLink() + '">&#160;';
        //sb[sb.length] = '</td>';
        sb[sb.length] = '<div class="ygtvspacer"></div></td>';

        // check box
        sb[sb.length] = '<td';
        sb[sb.length] = ' id="' + this.getCheckElId() + '"';
        sb[sb.length] = ' class="' + this.getCheckStyle() + '"';
        sb[sb.length] = ' onclick="javascript:' + this.getCheckLink() + '">';
        //sb[sb.length] = '&#160;</td>';
        sb[sb.length] = '<div class="ygtvspacer"></div></td>';


        sb[sb.length] = '<td>';
        sb[sb.length] = '<a';
        sb[sb.length] = ' id="' + this.labelElId + '"';
        sb[sb.length] = ' class="' + this.labelStyle + '"';
        sb[sb.length] = ' href="' + this.href + '"';
        sb[sb.length] = ' target="' + this.target + '"';
        sb[sb.length] = ' onclick="return ' + getNode + '.onLabelClick(' + getNode + ')"';
        if (this.hasChildren(true)) {
            sb[sb.length] = ' onmouseover="document.getElementById(\'';
            sb[sb.length] = this.getToggleElId() + '\').className=';
            sb[sb.length] = 'YAHOO.widget.TreeView.getNode(\'';
            sb[sb.length] = this.tree.id + '\',' + this.index + ').getHoverStyle()"';
            sb[sb.length] = ' onmouseout="document.getElementById(\'';
            sb[sb.length] = this.getToggleElId() + '\').className=';
            sb[sb.length] = 'YAHOO.widget.TreeView.getNode(\'';
            sb[sb.length] = this.tree.id + '\',' + this.index + ').getStyle()"';
        }
        sb[sb.length] = (this.nowrap) ? ' nowrap="nowrap" ' : '';
        sb[sb.length] = ' >';
        sb[sb.length] = this.label;
        sb[sb.length] = '</a>';
        sb[sb.length] = '</td>';
        sb[sb.length] = '</tr>';
        sb[sb.length] = '</table>';

        return sb.join("");

    },

    toString: function() {
        return "TaskNode (" + this.index + ") " + this.label;
    }

});
