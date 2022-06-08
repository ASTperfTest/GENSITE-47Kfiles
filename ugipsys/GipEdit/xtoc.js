//TABLE OF CONTENTS TREE SCRIPT
//last updated: 2/04/05 7:25 PM

//warning: any element in html document that has its id == TOC_CONTAINER_ID
//	and is not a ul will be replaced

/*
constants or default values
*/

//depth of XML returned from the server via a XML request
//can be set by a public method
var DEPTH_PER_XML = 2

//when entry in TOC is this number away from maximum depth,
//an XML request is sent to the server, requesting more info
//make sure MIN_DEPTH_LEFT_PER_XML < DEPTH_PER_XML
var MIN_DEPTH_LEFT_PER_XML = 1

//url of xml servlet
//can be set by a public method
var SERVLET_URL = 'http://localhost:8080/toc/servlet/com.hyweb.gip.maian.DBToXMLServlet'

//ID of root of TOC tree
var TREE_ROOT_ID = 'treeRoot'

//ID of TOC tree root container
var TOC_CONTAINER_ID = 'toc'

//if true, whole TOC is requested recursively
//if false, only the first level of TOC is requested
var bLOAD_ALL = false

//whenever a TOC category is opened and a XML request needs to be sent:
//	if true, displays category before and during XML request
//	if false, displays category after XML request
//no effect if bLOAD_ALL is true
var bOPEN_CATEGORY_ASYNC_WITH_XML_REQUEST = true

//functions that define what happens when you mouseover and mouseout a line in the toc
//this is necessary since for some stupid reason IE doesn't support :hover in css well
var origStyleColor
function MOUSEOVER_FUNC() {
	//this.style.backgroundColor = 'yellow'
	origStyleColor = this.style.color
	this.style.color = 'blue'
}
function MOUSEOUT_FUNC() {
	//this.style.backgroundColor = 'transparent'
	this.style.color = origStyleColor
}

//if true, displays loading status above the TOC
var bDISPLAY_STATUS = false

//image sources
function IMG_SRC () {}
var urlPath = location.href.substring(0, location.href.lastIndexOf('/') + 1)
IMG_SRC.CATEGORY_OPEN = urlPath + 'image/ftv2folderopen.gif'	//can be set by a public method
IMG_SRC.CATEGORY_CLOSED = urlPath + 'image/ftv2folderclosed.gif'	//can be set by a public method
IMG_SRC.DOCUMENT = urlPath + 'image/ftv2doc.gif'	//can be set by a public method
IMG_SRC.VERTICAL_LINE = urlPath + 'image/ftv2vertline.gif'
IMG_SRC.NODE = urlPath + 'image/ftv2node.gif'
IMG_SRC.NODE_PLUS = urlPath + 'image/ftv2pnode.gif'
IMG_SRC.NODE_MINUS = urlPath + 'image/ftv2mnode.gif'
IMG_SRC.LAST_NODE = urlPath + 'image/ftv2lastnode.gif'
IMG_SRC.LAST_NODE_PLUS = urlPath + 'image/ftv2plastnode.gif'
IMG_SRC.LAST_NODE_MINUS = urlPath + 'image/ftv2mlastnode.gif'
IMG_SRC.BLANK = urlPath + 'image/ftv2blank.gif'

/*
public methods
*/

function setDepthPerXML (depth) {
	DEPTH_PER_XML = depth
	//make sure MIN_DEPTH_LEFT_PER_XML is always less than DEPTH_PER_XML
	if (MIN_DEPTH_LEFT_PER_XML >= depth)
		MIN_DEPTH_LEFT_PER_XML = depth - 1
}

function setCategoryOpenImgSrc(imgsrc) {
	if (imgsrc.search('http://') == -1)
		IMG_SRC.CATEGORY_OPEN = urlPath + imgsrc
	else
		IMG_SRC.CATEGORY_OPEN = imgsrc
}

function setCategoryClosedImgSrc(imgsrc) {
	if (imgsrc.search('http://') == -1)
		IMG_SRC.CATEGORY_CLOSED = urlPath + imgsrc
	else
		IMG_SRC.CATEGORY_CLOSED = imgsrc
}

function setDocumentImgSrc(imgsrc) {
	if (imgsrc.search('http://') == -1)
		IMG_SRC.DOCUMENT = urlPath + imgsrc
	else
		IMG_SRC.DOCUMENT = imgsrc
}

function setServletURL(url) {
	SERVLET_URL = url
}

/*
init
*/

window.onload = function () {

	cleanWhitespace(document)
	if (bDISPLAY_STATUS) {
		var p = document.getElementById('status')
		if (p == null) {
			p = document.createElement('p')
			p.id = 'status'
			p.appendChild(document.createTextNode('Status: '))
			var toc = document.getElementById('toc')
			toc.parentNode.insertBefore(p, toc)
		}
		statusText = document.createTextNode('')
		p.appendChild(statusText)
		p.style.display = 'list-item'
	} else {
		var p = document.getElementById('status')
		if (p != null)
			p.style.display = 'none'
	}		

	getXML(function () {
		var li = document.getElementById(TREE_ROOT_ID)
		click(li)
	}, TREE_ROOT_ID)	
}

/*
XML request functions
*/

var statusText
var xmlHttpReqQueue = new Array()

//func: function called after XML is loaded
//arguments[1+]: IDs
function getXML(func) {	
	if (arguments.length < 2)
		return
	var ids = new Array(arguments.length - 1)
	for (var i = 1; i < arguments.length; i++) {
		ids[i - 1] = arguments[i]
	}
	var idstr = ids[0]	
	for (var i = 1; i < ids.length; i++) {
		idstr += ',' + ids[i]
	}

	var xmlHttpReq
	//if Mozilla
	if (document.implementation && document.implementation.createDocument) {
		xmlHttpReq = new XMLHttpRequest()
		xmlHttpReq.onload = function () {
			onloadFunc(xmlHttpReq, func)
		}
	// if IE
	} else if (window.ActiveXObject) {
		xmlHttpReq = new ActiveXObject('Msxml2.XMLHTTP')
		xmlHttpReq.onreadystatechange = function () {
			if (xmlHttpReq.readyState != 4)
				return
			onloadFunc(xmlHttpReq, func)
		}
 	} else {
		alert('error: your browser cannot handle this script')
		return
	}

	var url = SERVLET_URL + '?id=' + idstr +
		'&depth=' + DEPTH_PER_XML	
	
	xmlHttpReq.open('GET', url, true)

	xmlHttpReqQueue.push(xmlHttpReq)
	//if queue was previously empty, there is no request being sent
	//so send this request immediately
	//alert(url);	
	if (xmlHttpReqQueue.length == 1) {
		//alert("123");	
		if (bDISPLAY_STATUS)
			statusText.data = "loading..."
		//alert("1232" + xmlHttpReq);		
		xmlHttpReq.send(null)
		//alert("-23");	
	}
	//alert("123__");
	
}

//IE doesn't accept null for a value, so need to check for null
function setRequestHeader(xmlHttpReq, header, value) {
	if (value != null)
		xmlHttpReq.setRequestHeader(header, value)
}

function onloadFunc(xmlHttpReq, func) {
	if (xmlHttpReq.status == 200) {
		buildBaseNodes(xmlHttpReq.responseXML)
		if (func != null) {
			func()
		}
		xmlHttpReqQueue.shift()
		//send the next request if queue isn't empty
		if (xmlHttpReqQueue.length > 0) {
			xmlHttpReqQueue[0].send(null)
		} else {
			if (bDISPLAY_STATUS)
				statusText.data = "done loading"
		}
	} else
		alert(xmlHttpReq.status + ': ' + xmlHttpReq.statusText)
}

/*
build functions
*/

//too bad javascript doesn't support enumerations...
function NODE_STYLE() {}
NODE_STYLE.NO_NODE = 0
NODE_STYLE.NODE = 1
NODE_STYLE.LAST_NODE = 2

function buildBaseNodes(xmlDoc) {
	var root = xmlDoc.documentElement
	for (var node = root.firstChild; node != null; node = node.nextSibling) {
		var ID = getFirstChildElementByTagName(node, 'ID').firstChild.data
		if (ID == TREE_ROOT_ID) {
			var toc = document.getElementById(TOC_CONTAINER_ID)
			var ul
			if (toc == null) {	//if not present, create it and append it to HTML body
				ul = document.createElement('ul')
				getFirstChildElementByTagName(document.documentElement, 'BODY').appendChild(ul)
				ul.id = TOC_CONTAINER_ID
			} else if (toc.nodeName != 'UL') {	//if not ul, make it ul
				ul = document.createElement('ul')
				toc.parentNode.replaceChild(ul, toc)
				ul.id = TOC_CONTAINER_ID
			} else {
				ul = toc
			}
			var drawList = document.createDocumentFragment()
			buildNode(node, ul, drawList, DEPTH_PER_XML, NODE_STYLE.NO_NODE, ID)
		} else {
			var li = document.getElementById(ID)
			if (li) {
				//get drawList, extract nodeStyle,
				//	then remove last draw node, since it will be replaced in buildNode
				var drawList = getDrawList(li)
				var nodeStyle = getNodeStyle(drawList)
				drawList.removeChild(drawList.lastChild)
				buildNode(node, li.parentNode, drawList, DEPTH_PER_XML, nodeStyle, ID)
			}
		}
	}
}

//last argument is put in for efficiency reasons
//it can be omitted
function buildNode(node, ul, drawList, depthLeft, nodeStyle, ID) {
	var li, nodeType
	if (ID == null)
		ID = getFirstChildElementByTagName(node, 'ID').firstChild.data
	nodeType = getFirstChildElementByTagName(node, 'nodeType').firstChild.data
	if (nodeType == 'U') {	//is a document
		li = document.getElementById(ID)
		//create description row if it doesn't exist
		if (li == null) {
			var newDrawList = drawList.cloneNode(true)
			if (nodeStyle == NODE_STYLE.NODE)
				newDrawList.appendChild(createImg(IMG_SRC.NODE))
			else if (nodeStyle == NODE_STYLE.LAST_NODE)
				newDrawList.appendChild(createImg(IMG_SRC.LAST_NODE))
			newDrawList.appendChild(createImg(IMG_SRC.DOCUMENT))
			li = buildDocumentRow(node, ul, newDrawList, ID)
			//for some stupid reason, IE doesn't support :hover in css well,
			//	so I have to script it
			li.onmouseover = MOUSEOVER_FUNC
			li.onmouseout = MOUSEOUT_FUNC
		//make sure node is updated correctly
		//(e.g. in the case that this node is no longer the last node)
		} else {
			/*if (nodeStyle != NODE_STYLE.NO_NODE) {
				var img = getLastChildElementByTagName(li, 'IMG')
				img.previousSibling.setAttribute('src', IMG_SRC.NODE)
			}*/
		}
	} else if (nodeType == 'C') {	//is a category
		var children = getFirstChildElementByTagName(node, 'children')
		li = document.getElementById(ID)
		//create category header row if it doesn't exist
		if (li == null) {
			//determine how drawList is to be adjusted
			//then adjust it and create the header row
			if (children) {
				//if there are children, after the row is created,
				//	adjust drawList for children and sibling nodes
				if (nodeStyle == NODE_STYLE.NODE) {
					var img = createImg(IMG_SRC.NODE_PLUS)
					drawList.appendChild(img)
					var newDrawList = drawList.cloneNode(true)
					newDrawList.appendChild(createImg(children == null && depthLeft > 0 ? IMG_SRC.CATEGORY_OPEN : IMG_SRC.CATEGORY_CLOSED))
					li = buildCategoryHeaderRow(node, ul, newDrawList, ID,
						new Array(newDrawList.lastChild, newDrawList.lastChild.previousSibling))
					img.setAttribute('src', IMG_SRC.VERTICAL_LINE)
				} else if (nodeStyle == NODE_STYLE.LAST_NODE) {
					var img = createImg(IMG_SRC.LAST_NODE_PLUS)
					drawList.appendChild(img)
					var newDrawList = drawList.cloneNode(true)
					newDrawList.appendChild(createImg(children == null && depthLeft > 0 ? IMG_SRC.CATEGORY_OPEN : IMG_SRC.CATEGORY_CLOSED))
					li = buildCategoryHeaderRow(node, ul, newDrawList, ID,
						new Array(newDrawList.lastChild, newDrawList.lastChild.previousSibling))
					img.setAttribute('src', IMG_SRC.BLANK)
				} else {	//NODE_STYLE.NO_NODE
					var newDrawList = drawList.cloneNode(true)
					newDrawList.appendChild(createImg(children == null && depthLeft > 0 ? IMG_SRC.CATEGORY_OPEN : IMG_SRC.CATEGORY_CLOSED))
					li = buildCategoryHeaderRow(node, ul, newDrawList, ID,
						new Array(newDrawList.lastChild))
				}
			} else {
				if (nodeStyle == NODE_STYLE.NODE) {
					var newDrawList = drawList.cloneNode(true)
					newDrawList.appendChild(createImg(depthLeft == 0 ? IMG_SRC.NODE_PLUS : IMG_SRC.NODE))
					newDrawList.appendChild(createImg(children == null && depthLeft > 0 ? IMG_SRC.CATEGORY_OPEN : IMG_SRC.CATEGORY_CLOSED))
					li = buildCategoryHeaderRow(node, ul, newDrawList, ID,
						new Array(newDrawList.lastChild, newDrawList.lastChild.previousSibling))
				} else if (nodeStyle == NODE_STYLE.LAST_NODE) {
					var newDrawList = drawList.cloneNode(true)
					newDrawList.appendChild(createImg(depthLeft == 0 ? IMG_SRC.LAST_NODE_PLUS : IMG_SRC.LAST_NODE))
					newDrawList.appendChild(createImg(children == null && depthLeft > 0 ? IMG_SRC.CATEGORY_OPEN : IMG_SRC.CATEGORY_CLOSED))
					li = buildCategoryHeaderRow(node, ul, newDrawList, ID,
						new Array(newDrawList.lastChild, newDrawList.lastChild.previousSibling))
				} else {	//NODE_STYLE.NO_NODE
					var newDrawList = drawList.cloneNode(true)
					newDrawList.appendChild(createImg(children == null && depthLeft > 0 ? IMG_SRC.CATEGORY_OPEN : IMG_SRC.CATEGORY_CLOSED))
					li = buildCategoryHeaderRow(node, ul, newDrawList, ID,
						new Array(newDrawList.lastChild))
				}
			}
			//for some stupid reason, IE doesn't support :hover in css well,
			//	so I have to script it
			li.onmouseover = MOUSEOVER_FUNC
			li.onmouseout = MOUSEOUT_FUNC
		} else {
			if (children) {
				//adjust drawList for children nodes
				if (nodeStyle == NODE_STYLE.NODE) {
					drawList.appendChild(createImg(IMG_SRC.VERTICAL_LINE))
				} else if (nodeStyle == NODE_STYLE.LAST_NODE) {
					drawList.appendChild(createImg(IMG_SRC.BLANK))
				}	//else NODE_STYLE.NO_NODE
			} else {
				//there is a case where an empty category show as if they did have children
				//	(see "benefit of doubt" case near end of function)
				//so update it accordingly and remove its onclick events
				var img = getLastChildElementByTagName(li, 'IMG')
				img.setAttribute('src', IMG_SRC.CATEGORY_OPEN)
				img = img.previousSibling
				if (img.getAttribute('src') == IMG_SRC.NODE_PLUS) {
					img.setAttribute('src', IMG_SRC.NODE)
				} else if (img.getAttribute('src') == IMG_SRC.LAST_NODE_PLUS) {
					img.setAttribute('src', IMG_SRC.LAST_NODE)
				}
				setOnclick(li, null)
			}
		}
		//category may be empty, so test if there are children nodes
		if (children != null) {
			var newUl
			//create list of category's contents and the row that will contain it for child nodes if it doesn't exist
			if (li.contents == null) {
				newUl = document.createElement('ul')
				var newLi = document.createElement('li')
				//li.contents primary purpose is to easily show/hide the list of category's contents
				//however, setting ul's display to none won't work,
				//	since IE will display the li that contains it as a blank line,
				//	so li.contents must be the containing li,
				//	and this li's display will be set to none
				li.contents = newLi
				newLi.appendChild(newUl)
				if (li == ul.lastChild) {
					ul.appendChild(newLi)
				} else {
					ul.insertBefore(newLi, li.nextSibling)
				}
				//assign event handlers now that li.contents exists
				li.contents.style.display = 'none'
				var func
				if (depthLeft <= MIN_DEPTH_LEFT_PER_XML && !bLOAD_ALL) {	//need to send a XML request
					if (bOPEN_CATEGORY_ASYNC_WITH_XML_REQUEST && MIN_DEPTH_LEFT_PER_XML > 0) {
						func = function () {
							openCategory(li, nodeStyle)
							getXML(null, li.id)
						}
					} else {
						func = function () {
							getXML(function () {
								openCategory(li, nodeStyle)
							}, li.id)
						}
					}
				} else {	//don't need to send a XML request
					func = function () {
						openCategory(li, nodeStyle)
					}
				}
				setOnclick(li, func)
			} else {
				//li.contents is the li containing the contents list (ul)
				newUl = li.contents.firstChild
			}
			//recurse through child nodes
			var child = children.firstChild, childLi
			for (i = 0; child != children.lastChild; child = child.nextSibling, i++) {
				childLi = buildNode(child, newUl, drawList, depthLeft - 1, NODE_STYLE.NODE)
				childLi.up = li
			}
			childLi = buildNode(child, newUl, drawList, depthLeft - 1, NODE_STYLE.LAST_NODE)
			childLi.up = li
			//if there were children, an img was appended to drawList,
			//	so remove last draw img to remove it
			if (nodeStyle != NODE_STYLE.NO_NODE) {
				drawList.removeChild(drawList.lastChild)
			}
		} else {
			//if 0 depthLeft, give benefit of doubt,
			//	and assume that once an XML request is sent,
			//	they will have children
			//if they actually don't have children, they will be updated accordingly
			if (depthLeft == 0) {
				if (bLOAD_ALL) {
					getXML(null, li.id)
					func = function () {
						openCategory(li, nodeStyle)
					}
				} else {
					if (bOPEN_CATEGORY_ASYNC_WITH_XML_REQUEST) {
						func = function () {
							openCategory(li, nodeStyle)
							getXML(null, li.id)
						}
					} else {
						func = function () {
							getXML(function () {
								openCategory(li, nodeStyle)
							}, li.id)
						}
					}
				}
				setOnclick(li, func)
			}
		}
	}
	return li
}

//TODO: fix function to make compatible with Mozilla
function buildDocumentRow(node, ul, drawList, ID) {
	var li = document.createElement('li')
	li.id = ID
	//append drawList
	li.appendChild(drawList)
	//var catRootID = getFirstChildElementByTagName(node, 'catRootID')
	var desc = getFirstChildElementByTagName(node, 'desc')
	//note: following not Mozilla compliant
	li.insertAdjacentHTML("beforeEnd", desc.firstChild.xml)
	ul.appendChild(li)
	return li
}

function buildCategoryHeaderRow(node, ul, drawList, ID, onclickList) {
	var li = document.createElement('li')
	li.id = ID
	//append drawList
	li.appendChild(drawList)
	//var catRootID = getFirstChildElementByTagName(node, 'catRootID')
	var desc = getFirstChildElementByTagName(node, 'desc')
	//note: following not Mozilla compliant
	li.insertAdjacentHTML("beforeEnd", desc.firstChild.xml)
	//sets the list of elements that will have their onclick funcs set
	//does not set those onclick funcs yet
	li.onclickList = onclickList
	ul.appendChild(li)
	return li
}

/*
event functions
*/

function setOnclick(li, func) {
	for (var i = 0; i < li.onclickList.length; i++)
		li.onclickList[i].onclick = func
}

function click(li) {
	if (li.contents) {
		li.onclickList[0].onclick()
	}
}

function closeCategory(li, nodeStyle) {
	var func = function () {
		openCategory(li, nodeStyle)
	}
	setOnclick(li, func)
	li.contents.style.display = 'none'
	//change last img
	var img = li.lastChild
	while (img != null && img.nodeName != 'IMG')
		img = img.previousSibling
	img.setAttribute('src', IMG_SRC.CATEGORY_CLOSED)
	//change second last img if it exists
	img = img.previousSibling
	if (img) {
		if (nodeStyle == NODE_STYLE.NODE)
			img.setAttribute('src', IMG_SRC.NODE_PLUS)
		else if (nodeStyle == NODE_STYLE.LAST_NODE)
			img.setAttribute('src', IMG_SRC.LAST_NODE_PLUS)
	}
}

function openCategory(li, nodeStyle) {
	var func = function () {
		closeCategory(li, nodeStyle)
	}
	setOnclick(li, func)
	li.contents.style.display = 'list-item'
	//change last img
	var img = li.lastChild
	while (img != null && img.nodeName != 'IMG')
		img = img.previousSibling
	img.setAttribute('src', IMG_SRC.CATEGORY_OPEN)
	//change second last img if it exists
	img = img.previousSibling
	if (img) {
		if (nodeStyle == NODE_STYLE.NODE)
			img.setAttribute('src', IMG_SRC.NODE_MINUS)
		else if (nodeStyle == NODE_STYLE.LAST_NODE)
			img.setAttribute('src', IMG_SRC.LAST_NODE_MINUS)
	}
}

/*
drawList functions
*/

//returns clone of li's drawList, including all img elements except the last
function getDrawList(li) {
	var drawList = document.createDocumentFragment()
	if (li.hasChildNodes()) {
		for (var child = li.firstChild; child.nextSibling != null && child.nextSibling.nodeType == 1; child = child.nextSibling)
			drawList.appendChild(child.cloneNode(true))
	}
	return drawList
}

function getNodeStyle(drawList) {
	if (drawList.hasChildNodes()) {
		var lastDrawImgSrc = drawList.lastChild.getAttribute('src')
		if (lastDrawImgSrc == IMG_SRC.NODE ||
			lastDrawImgSrc == IMG_SRC.NODE_PLUS ||
			lastDrawImgSrc == IMG_SRC.NODE_MINUS) {
			return NODE_STYLE.NODE
		} else if (lastDrawImgSrc == IMG_SRC.LAST_NODE ||
			lastDrawImgSrc == IMG_SRC.LAST_NODE_PLUS ||
			lastDrawImgSrc == IMG_SRC.LAST_NODE_MINUS) {
			return NODE_STYLE.LAST_NODE
		}
	} else {
		return NODE_STYLE.NO_NODE
	}
}

function createImg(imgsrc) {
	var img = document.createElement('img')
	img.setAttribute('src', imgsrc)
	return img
}

/*
misc functions for convenience and compatibility between Mozilla and IE
*/

//removes all useless text nodes that Mozilla keeps between element nodes
function cleanWhitespace(node) {
	for (var child = node.firstChild; child != null; child = child.nextSibling) {
		if ((child.nodeType == 3) && (!/\S/.test(child.nodeValue))) {	//if whitespace text node
			node.removeChild(child)
		}
		if (child.nodeType == 1) {	//elements can have text child nodes of their own
			cleanWhitespace(child)
		}
	}
}

//note: case-sensitive
function getFirstChildElementByTagName(node, tagName) {
	var child = node.firstChild
	while (child != null && child.tagName != tagName)
		child = child.nextSibling
	return child
}

//note: cast-sensitive
function getLastChildElementByTagName(node, tagName) {
	var child = node.lastChild
	while (child != null && child.tagName != tagName)
		child = child.previousSibling
	return child
}
