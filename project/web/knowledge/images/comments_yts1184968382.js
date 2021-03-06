// The following functions rely on javascript found in templates/includes/_comments2.tmpl

function approveComment(comment_id, comment_type, entity_id, token)
{
	if (CheckLogin() == false)
		return false;
	
	//if(!confirm("Really approve this comment?"))
	//	return true;
	
	//postFormByForm(form, true, execOnSuccess(commentApproved));
	//postUrl("/comment_servlet",  urlEncodeDict(formVars), true, execOnSuccess(commentApproved));
	
	postUrlXMLResponse("/comment_servlet", "&field_approve_comment=1&comment_id=" + comment_id + "&comment_type=" + comment_type + "&entity_id=" + entity_id + "&AXT=" + token, self.commentApproved);

	return false;
}
	


function commentApproved(xmlHttpRequest)
{
	alert("Comment approved.")
}


function removeComment(div_id, deleter_user_id, comment_id, comment_type, entity_id, token)
{
	self.div_id = div_id
	self.commentRemoved = commentRemoved
	if (CheckLogin() == false)
		return;

	//if (!confirm("Really remove comment?"))
	//	return;
	
	postUrlXMLResponse("/comment_servlet", "deleter_user_id=" + deleter_user_id + "&remove_comment&comment_id=" + comment_id + "&comment_type=" + comment_type + "&entity_id=" + entity_id + "&AXT=" + token, self.commentRemoved);

	return false;
}
function commentRemoved(xmlHttpRequest)
{
	toggleVisibility(self.div_id, false);
	return;
}
		
function hideCommentReplyForm(form_id) {
	var div_id = "div_" + form_id;
	var reply_id = "reply_" + form_id;
	toggleVisibility(reply_id, true);
	toggleVisibility(div_id, false);
	//setInnerHTML(div_id, "");
}

function handleStateChange(xmlHttpReq) {
	document.getElementById("all_comments_content").innerHTML=getNodeValue(xmlHttpReq.responseXML, "html_content");
	
	style2 = document.getElementById("recent_comments").style;
	style2.display = "none";
	
	var style2 = document.getElementById("all_comments").style;
	style2.display = "";
}
