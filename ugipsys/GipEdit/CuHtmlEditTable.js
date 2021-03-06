//last modified: 11/15/05 11:32 PM

//uses IE-specific modal dialog properties, TextRange, and getBoundingClientRect,
//	so this script is not cross-browser

//read dialog arguments

var args = window.dialogArguments;
var oldTable = null, oldHtmlText = null;
var table, doc, range;
//if table element was passed
if (args.tagName && args.tagName.toLowerCase() == 'table') {
	table = args;
	range = null;
	doc = table.ownerDocument;
//else TextRange or ControlRange object was passed
} else {
	table = null;
	range = args;
	doc = args.length? args(0).ownerDocument : args.parentElement().ownerDocument;
}

//grid data structure
//note: there's a distinction between data cells and table cells;
//	data cells correspond to this grid data structure,
//	while table cells correspond to the grid table (which is the presented view of the grid data)

var data = [];
data.numRows = 0;
data.numCols = 0;

function DataCell(tableRef, gridRef, type, colSpan, rowSpan) {
	this.tableRef = tableRef || null;
	this.gridRef = gridRef || null;
	this.type = type || 'td';
	this.colSpan = colSpan || 1;
	this.rowSpan = rowSpan || 1;
}

function DataRow(tableRef, gridRef) {
	var self = [];
	self.tableRef = tableRef || null;
	self.gridRef = gridRef || null;
	return self;
}

//editor buttons

function InsertColButton(isLast) {
	var self = document.createElement('td');
	//if (!isLast)
	//	self.style.borderRight = '1px solid blue';
	//TODO: replace span with some img
	if (isLast)
		self.className = 'last';
	var img = document.createElement('span');
	img.onclick = insertCol;
	img.innerHTML = '+';
	self.appendChild(img);
	return self;
}

function DeleteColButton(isLast) {
	var self = document.createElement('td');
	//if (!isLast)
	//	self.style.borderRight = '1px solid blue';
	//TODO: replace span with some img
	if (isLast)
		self.className = 'last';
	var img = document.createElement('span');
	img.onclick = deleteCol;
	img.innerHTML = 'x';
	self.appendChild(img);
	return self;
}

function InsertRowButton(isLast) {
	var self = document.createElement('tr');
	var cell = document.createElement('td');
	//if (!isLast)
	//	cell.style.borderBottom = '1px solid blue';
	//TODO: replace span with some img
	if (isLast)
		cell.className = 'last';
	var img = document.createElement('span');
	img.onclick = insertRow;
	img.innerHTML = '+';
	cell.appendChild(img);
	self.appendChild(cell);
	return self;
}

function DeleteRowButton(isLast) {
	var self = document.createElement('tr');
	var cell = document.createElement('td');
	//if (!isLast)
	//	cell.style.borderBottom = '1px solid blue';
	//TODO: replace span with some img
	if (isLast)
		cell.className += 'last';
	var img = document.createElement('span');
	img.onclick = deleteRow;
	img.innerHTML = 'x';
	cell.appendChild(img);
	self.appendChild(cell);
	return self;
}

//init

var editor, editorHorz, editorVert;
var editorBarThickness, footerHeight, editorOffsetTop, bodyPadding;
var grid, gridCellHeight, gridCellWidth;
var caption, summary, cellSpacing, cellPadding, border, width, layoutMode;

var reNonDigit = /\D/;

function init() {
	caption = document.getElementById('caption');
	summary = document.getElementById('summary');
	cellSpacing = document.getElementById('cellSpacing');
	cellPadding = document.getElementById('cellPadding');
	border = document.getElementById('border');
	width = document.getElementById('width');
	editor = document.getElementById('editor');

	//retrieve table data and populate grid data structure
	if (table) {
		initTextInputs();
		initData();
		initLayoutMode();
	} else {
		initNewTable();
	}
	initEditor();
	initResize();
}

function initTextInputs() {
	caption.value = table.caption? table.caption.innerText : '';
	caption.originalValue = caption.value;
	caption.onchange = verifyOption;
	caption.onchange();
	summary.value = table.summary;
	summary.originalValue = summary.value;
	summary.onchange = verifyOption;
	summary.onchange();
	cellSpacing.value = table.cellPadding;
	cellSpacing.originalValue = cellSpacing.value;
	cellSpacing.onchange = verifyOptionInteger;
	cellSpacing.onchange();
	cellPadding.value = table.cellPadding;
	cellPadding.originalValue = cellPadding.value;
	cellPadding.onchange = verifyOptionInteger;
	cellPadding.onchange();
	border.value = table.border;
	border.originalValue = border.value;
	border.onchange = verifyOptionInteger;
	border.onchange();
	width.value = table.width;
	width.onchange = verifyOptionPercentage;
	width.originalValue = width.value;
	if (width.value) {
		if (width.value.indexOf('%') == -1) {
			if (width.value.search(reNonDigit) == -1)
				width.value += 'px';
		} else {
			width.value = width.value.substr(0, width.value.length - 1);
		}
		width.originalValue = width.value;
	}
	width.onchange();
}

function initData() {
	var rows = table.rows;
	var dataRow, prevDataRow;
	data.numRows = rows.length;
	//row loop
	for (var i = 0; i < data.numRows; i++) {
		var row = rows[i];
		var cells = row.cells;
		prevDataRow = dataRow;
		dataRow = new DataRow(row, null);
		//reserve cells for cells with rowSpan > 1 in above rows
		if (prevDataRow) {
			//i = index of current data row
			for (var j = 0; j < prevDataRow.length;) {
				var prevDataCell = prevDataRow[j];
				if (prevDataCell && prevDataCell.rowRem) {
					var colSpan = prevDataCell.colSpan;
					for (var l = 0; l < colSpan; l++)
						dataRow[j++] = prevDataCell;
					prevDataCell.rowRem--;
				} else
					j++;
			}
		}
		//column loop
		for (var j = 0, k = 0; j < cells.length; j++) {
			var cell = cells[j];
			var colSpan = cell.colSpan;
			var rowSpan = cell.rowSpan;
			//make sure no cells overlap with cells that have rowSpan > 1 in above rows
			for (var l = 0; l < colSpan;) {
				//if cell already exists, skip it
				if (dataRow[k + l]) {
					k = k + l + 1;
					//restart loop (check for each colSpan again)
					l = 0;
				} else
					l++;
			}
			//create data cell
			var dataCell = new DataCell(cell, null, cell.tagName.toLowerCase(), colSpan, rowSpan);
			//temporary property representing how many vertical "subcells"
			//	haven't been visited yet:
			dataCell.rowRem = rowSpan - 1;
			for (var l = 0; l < colSpan; l++)
				dataRow[k++] = dataCell;
		}
		data[i] = dataRow;
	}
	//each row should have the same # columns
	data.numCols = dataRow.length;
	//account for any rowSpan > 1 cells in above rows
	//i = index of last data row; k = # cells in last data row
	for (j = 0; j < k; j++) {
		var dataCell = dataRow[j];
		if (dataCell) {
			var maxRow = i + dataCell.rowRem;
			dataCell.rowRem = 0;
			//append new rows to data as needed
			while (maxRow > data.numRows) {
				data[data.numRows++] = [];
			}
			for (l = i; l < maxRow; l++) {
				data[l][j] = dataCell;
			}
		}
	}
}

function initLayoutMode() {
	//calculate the table's layout mode (top header, left header, both header, no header, or other)
	//default to "other"
	layoutMode = 'Other';
	lblLayout: {
		var topRow = data[0];
		if (data.numRows && data.numCols) {
			//top row is all TH, no TH, or some TH?
			//all TH => "top header" or "both header", depending on left column
			//no TH => "no header" or "left header", depending on left column
			//some TH => "other"
			var headerTop = false;
			//if top row's right-most cell is TH, either "all TH" or "some TH"
			//the reason top row's left-most cell isn't checked is that:
			//	1)	if there is more than 1 cell and left-most cell is TH, it could be the case that
			//		that cell is part of the left column and not the top row, in which case,
			//		it wouldn't be "all TH"
			//	2)	the only case this wouldn't be ambiguous would be if there were only 1 cell on
			//		the row, in which case it would also be "all TH"
			if (topRow[topRow.length - 1].type == 'th') {
				//all cells must be TH to be "all TH"
				//if a cell is not TH, "some TH"
				for (var i = 0; i < topRow.length;) {
					var dataCell = topRow[i];
					if (dataCell.type != 'th') {
						break lblLayout;
					}
					i += dataCell.colSpan;
				}
				headerTop = true;
			}

			//left column is all TH, no TH, or some TH?
			//all TH => "left header" or "both header", depending on top header
			//no TH => "no header" or "top header", depending on top row
			//some TH => "other"
			var headerLeft = false;
			//if bottom row's left-most cell is TH, either "all TH" or "some TH"
			//the reason top row's left-most cell isn't checked is the same reason as before
			if (data[data.numRows - 1][0].type == 'th') {
				//all cells must be TH to be "all TH"
				//if a cell is not TH, "some TH"
				for (var i = 0; i < data.numRows;) {
					var dataCell = data[i][0];
					if (dataCell.type != 'th') {
						break lblLayout;
					}
					i += dataCell.rowSpan;
				}
				headerLeft = true;
			}

			//if both "top header" and "left header", then "both header"
			if (headerTop && headerLeft)
				layoutMode = 'Both';
			else if (headerTop)
				layoutMode = 'Top';
			else if (headerLeft)
				layoutMode = 'Left';
			//if neither, then either "no header" or "other"
			else {
				//if there are any TH in top row, then "other"
				//don't need to check the right-most cell was already checked,
				//	since it was already checked
				var i = data.numRows - 1;
				i -= data[i][0].colSpan;
				while (i >= 0) {
					var dataCell = topRow[i];
					if (dataCell.type == 'th') {
						break lblLayout;
					}
					i -= dataCell.colSpan;
				}
				//ditto for left column
				var i = data.numRows - 1;
				i -= data[i][0].rowSpan;
				while (i >= 0) {
					var dataCell = data[i][0];
					if (dataCell.type == 'th') {
						break lblLayout;
					}
					i -= dataCell.rowSpan;
				}
				//if it's not "other", then "no header"
				layoutMode = 'None';
			}
		}
	}
	if (layoutMode != 'Other') {
		document.getElementById('layoutH' + layoutMode).checked = true;
	}
}

function initNewTable() {
	data.numRows = 1;
	data.numCols = 1;
	data[0] = [new DataCell()];
	layoutMode = 'None';
}

function initEditor() {
	initEditorHorz();
	initEditorVert();
	initEditorGrid();

	//editor top-left corner img (intersection of editor bars)
	var editorTopLeft = document.createElement('div');
	editorTopLeft.id = 'editorTopLeft';
	editor.appendChild(editorTopLeft);
	//since the img, which is providing the diagonal line on the top-left corner of the grid,
	//	is overlapping both the first buttons of the vertical and horizontal editor bars
	//	need to make sure clicks are passed onto those buttons
	editorTopLeft.onclick = editorBarOverlapClickThru;
}

function initEditorHorz() {
	//horizontal editor bar
	editorHorz = document.createElement('table');
	editorHorz.id = 'editorHorz';
	editorHorz.cellSpacing = 0;
	editorHorz.cellPadding = 0;
	var tbody = document.createElement('tbody');
	var row = document.createElement('tr');
	for (var i = 0; i < data.numCols; i++) {
		row.appendChild(new InsertColButton());
		row.appendChild(new DeleteColButton());
	}
	row.appendChild(new InsertColButton(true));
	tbody.appendChild(row);
	editorHorz.appendChild(tbody);
	editor.appendChild(editorHorz);
}

function initEditorVert() {
	//vertical editor bar
	editorVert = document.createElement('table');
	editorVert.id = 'editorVert';
	editorVert.cellSpacing = 0;
	editorHorz.cellPadding = 0;
	var tbody = document.createElement('tbody');
	for (var i = 0; i < data.numRows; i++) {
		tbody.appendChild(new InsertRowButton());
		tbody.appendChild(new DeleteRowButton());
	}
	tbody.appendChild(new InsertRowButton(true));
	editorVert.appendChild(tbody);
	editor.appendChild(editorVert);
}

function initEditorGrid() {
	//calc grid "margin"
	editorBarThickness = parseInt(editorHorz.currentStyle.height);

	//grid table
	//not an actual grid since it may have spans,
	//	but it's the live representation of the data object, which has a grid structure
	grid = document.createElement('table');
	grid.id = 'grid';
	grid.cellSpacing = 0;
	grid.cellPadding = 0;
	var gridCellWidth = Math.round((editor.offsetWidth - editorBarThickness * 2) / data.numCols);
	var gridCellHeight = Math.round((editor.offsetHeight - editorBarThickness * 2) / data.numRows);
	var tbody = document.createElement('tbody');
	for (var i = 0; i < data.numRows; i++) {
		var row = document.createElement('tr');
		var dataRow = data[i];
		for (var j = 0; j < data.numCols;) {
			var cell = document.createElement('td');
			var dataCell = dataRow[j];
			if (dataCell) {
				if (dataCell.tableRef) {
					//if dataCell's gridRef is already defined, dataCell has already been visited
					if (dataCell.gridRef) {
						j += dataCell.colSpan;
						continue;
					}
					cell.innerHTML = dataCell.tableRef.innerHTML;
					cell.className = 'old-' + dataCell.type;
					cell.colSpan = dataCell.colSpan;
					cell.rowSpan = dataCell.rowSpan;
					j += dataCell.colSpan;
				} else {
					cell.className = 'new-' + dataCell.type;
					j++;
				}
				dataCell.gridRef = cell;
			} else {
				//if execution ever reaches here, table's HTML is incorrect
				//TODO: replace the following with some sort of dialog that asks whether new
				//	cells should be added to fill in the gap
				alert(document.getElementByid('INVALID_TABLE_HTML').innerHTML);
				throw 'INVALID_TABLE_HTML';
			}
			row.appendChild(cell);
		}
		dataRow.gridRef = row;
		tbody.appendChild(row);
	}
	grid.appendChild(tbody);
	editor.appendChild(grid);
}

function initResize() {
	//for editor vertical resizing
	footerHeight = document.getElementById('buttons').offsetHeight;
	editorOffsetTop = editor.offsetTop;

	//editor resize
	window.onresize = windowOnresize;
	windowOnresize();
}

window.onload = init;

//IE-only:
//since the grid's top-left diagonal line img is overlapping both the first buttons of the
//	vertical and horizontal editor bars,
//	need to make sure clicks are passed onto those buttons
//(assigned to editorVert.onclick on window.onload)
function editorBarOverlapClickThru() {
	var x = event.clientX;
	var y = event.clientY;
	//get horizontal bar's first cell's child
	var button = editorHorz.firstChild.firstChild.firstChild.firstChild;
	var rect = button.getBoundingClientRect();
	//only need to check left, top, and bottom boundaries, since right boundary is not within
	//	the top-left img
	if (x >= rect.left && y >= rect.top && y <= rect.bottom) {
		button.fireEvent('onclick');
	}
	//get vertical bar's first cell's child
	var button = editorVert.firstChild.firstChild.firstChild.firstChild;
	var rect = button.getBoundingClientRect();
	//only need to check left, right, and top boundaries, since bottom boundary is not within
	//	the top-left img
	if (x >= rect.left && x <= rect.right && y >= rect.top) {
		button.fireEvent('onclick');
	}
}

//options input verification

function verifyOption() {
	var container = document.getElementById(this.id + 'Container');
	var desc = document.getElementById(this.id + 'Desc');
	if (this.originalValue != this.value) {
		container.className = 'changed';
		desc.innerHTML = document.getElementById('ORIGINAL_VALUE').innerHTML + ': "' + this.originalValue + '"';
		return;
	}
	container.className = '';
	desc.innerHTML = '';
}

function verifyOptionInteger() {
	var container = document.getElementById(this.id + 'Container');
	var desc = document.getElementById(this.id + 'Desc');
	//if value isn't an integer or is empty (string has non-digit character),
	//	error
	if (this.value.search(reNonDigit) != -1) {
		container.className = 'error';
		desc.innerHTML = document.getElementById('NOT_AN_INTEGER').innerHTML;
		return;
	}
	verifyOption.call(this);
}

function verifyOptionPercentage() {
	var container = document.getElementById(this.id + 'Container');
	var desc = document.getElementById(this.id + 'Desc');
	//allow value to be empty
	//if value is non-empty...
	if (this.value != '') {
		var intVal = parseInt(this.value);
		//if value isn't an integer (String(parseInt(str)) != str)
		//	or isn't between 0 and 100, error
		if (String(intVal) != this.value) {
			container.className = 'error';
			desc.innerHTML = document.getElementById('NOT_AN_INTEGER').innerHTML;
			return;
		}
		if (intVal < 1 || intVal > 100) {
			container.className = 'error';
			desc.innerHTML = document.getElementById('NOT_A_PERCENTAGE').innerHTML;
			return;
		}
	}
	verifyOption.call(this);
}

//editor resizing funcs
//note: col elements aren't necessary, since "table-layout: fixed" makes sure all columns
//	have the same width if no individual widths are defined;
//	however, row elements need to have their heights defined, since there are some cases
//	involving multi-row cells where rows don't have the same height even with fixed table
//	layout

function updateHorzStyle() {
	if (data.numCols == 0) {
		return;
	}
	var gridWidth = editor.offsetWidth - editorBarThickness * 2;
	grid.style.width = gridWidth + 'px';
	var gridCellWidth = gridWidth / data.numCols;

	var horzPadding = Math.round(gridCellWidth / 4);
	editorHorz.style.left = editorBarThickness - horzPadding + 'px';
	editorHorz.style.width = gridWidth + horzPadding * 2 + 'px';
}

function updateVertStyle() {
	if (data.numRows == 0) {
		return;
	}
	var gridHeight = editor.offsetHeight - editorBarThickness * 2;
	grid.style.height = gridHeight + 'px';
	var gridCellHeight = gridHeight / data.numRows;

	//adjust row heights
	var rows = grid.rows;
	var rowHeight = Math.round(gridCellHeight);
	rowHeight += 'px';
	for (var i = 0; i < data.numRows; i++) {
		rows[i].style.height = rowHeight;
	}

	var vertPadding = Math.round(gridCellHeight / 4);
	editorVert.style.top = editorBarThickness - vertPadding + 'px';
	editorVert.style.height = gridHeight + vertPadding * 2 + 'px';
}

var prevHeights = [];
var cancelResize = false;
function windowOnresize() {
	//see below comments
	if (cancelResize)
		return;

	var editorHeight = Math.round((document.documentElement.clientHeight - editorOffsetTop - footerHeight) * 0.95);
	if (editorHeight < grid.offsetTop * 2) {
		return;
	}

	//resize event seems to be triggered multiple times with slightly different
	//	window size values, as if the size is being recalculated several times
	//	(can't pinpoint what exactly is causing this, but its also sometimes triggered
	//	by changing dimensions of some elements)
	//in some cases, window switches back and forth between two sizes indefinitely;
	//	this seems to be caused by scrollbars appearing and disappearing,
	//	so in the CSS, html's overflow is set to hidden to prevent scrollbars from
	//	appearing
	//workaround:
	//1)	remember previous heights
	//2)	if current height matches one of the previous heights, then
	//		set height to the max of those previous and current heights
	//		and set a flag (cancelResize) to prevent this method from running again
	//3)	clear the flag and the stored previous heights afterwards
	//		in a 0-delay timeout
	for (var i = 0; i < prevHeights.length; i++) {
		if (prevHeights[i] == editorHeight) {
			editorHeight = Math.max(Math.max.apply(null, prevHeights), editorHeight);
			cancelResize = true;
		}
	}
	setTimeout(function() {
		prevHeights = [];
		cancelResize = false;
	}, 0);

	prevHeights.push(editorHeight);
	editor.style.height = editorHeight + 'px';
	updateHorzStyle();
	updateVertStyle();
}

//layout mode funcs

function layoutModeTop() {
	layoutMode = 'Top';
	if (!data.numRows || !data.numCols)
		return;
	//change all TDs on the top row to THs
	var topRow = data[0];
	for (var i = 0; i < topRow.length;) {
		var dataCell = topRow[i];
		if (dataCell.type != 'th') {
			dataCell.type = 'th';
			var className = dataCell.gridRef.className;
			dataCell.gridRef.className = className.substr(0, className.length - 2) + 'th';
		}
		i += dataCell.colSpan;
	}
	//change all THs on the left column to TDs
	//skip the first cell since it's already been visited
	for (var i = topRow[0].rowSpan; i < data.numRows;) {
		var dataCell = data[i][0];
		if (dataCell.type != 'td') {
			dataCell.type = 'td';
			var className = dataCell.gridRef.className;
			dataCell.gridRef.className = className.substr(0, className.length - 2) + 'td';
		}
		i += dataCell.rowSpan;
	}
}

function layoutModeLeft() {
	layoutMode = 'Left';
	if (!data.numRows || !data.numCols)
		return;
	//change all the TDs on the left column to THs
	for (var i = 0; i < data.numRows;) {
		var dataCell = data[i][0];
		if (dataCell.type != 'th') {
			dataCell.type = 'th';
			var className = dataCell.gridRef.className;
			dataCell.gridRef.className = className.substr(0, className.length - 2) + 'th';
		}
		i += dataCell.rowSpan;
	}
	//change all THs on the top row to TDs
	//skip the first cell since it's already been visited
	var topRow = data[0];
	for (var i = topRow[0].colSpan; i < topRow.length;) {
		var dataCell = topRow[i];
		if (dataCell.type != 'td') {
			dataCell.type = 'td';
			var className = dataCell.gridRef.className;
			dataCell.gridRef.className = className.substr(0, className.length - 2) + 'td';
		}
		i += dataCell.colSpan;
	}
}

function layoutModeBoth() {
	layoutMode = 'Both';
	if (!data.numRows || !data.numCols)
		return;
	//change all TDs on the top row to THs
	var topRow = data[0];
	for (var i = 0; i < topRow.length;) {
		var dataCell = topRow[i];
		if (dataCell.type != 'th') {
			dataCell.type = 'th';
			var className = dataCell.gridRef.className;
			dataCell.gridRef.className = className.substr(0, className.length - 2) + 'th';
		}
		i += dataCell.colSpan;
	}
	//change all the TDs on the left column to THs
	//skip the first cell since it's already been visited
	for (var i = topRow[0].rowSpan; i < data.numRows;) {
		var dataCell = data[i][0];
		if (dataCell.type != 'th') {
			dataCell.type = 'th';
			var className = dataCell.gridRef.className;
			dataCell.gridRef.className = className.substr(0, className.length - 2) + 'th';
		}
		i += dataCell.rowSpan;
	}
}

function layoutModeNone() {
	layoutMode = 'None';
	if (!data.numRows || !data.numCols)
		return;
	//change all THs on the top row to TDs
	var topRow = data[0];
	for (var i = 0; i < topRow.length;) {
		var dataCell = topRow[i];
		if (dataCell.type != 'td') {
			dataCell.type = 'td';
			var className = dataCell.gridRef.className;
			dataCell.gridRef.className = className.substr(0, className.length - 2) + 'td';
		}
		i += dataCell.colSpan;
	}
	//change all the THs on the left column to TDs
	//skip the first cell since it's already been visited
	for (var i = topRow[0].rowSpan; i < data.numRows;) {
		var dataCell = data[i][0];
		if (dataCell.type != 'td') {
			dataCell.type = 'td';
			var className = dataCell.gridRef.className;
			dataCell.gridRef.className = className.substr(0, className.length - 2) + 'td';
		}
		i += dataCell.rowSpan;
	}
}

//insertion/deletion funcs

function insertCol() {
	//find index of targeted col
	var target = event.srcElement.parentNode;
	var colIndex;
	var editorRow = editorHorz.firstChild.firstChild;
	var editorRowCols = editorRow.childNodes;
	for (var i = 0; i < editorRowCols.length; i++) {
		if (editorRowCols[i] == target) {
			colIndex = i / 2;
			break;
		}
	}

	//expand editorHorz
	editorRow.insertBefore(new InsertColButton(), target);
	editorRow.insertBefore(new DeleteColButton(), target);

	//insert new col before selected col
	//strategy: first, splice in empty data cells for each row
	//	(so comparing two adjacent rows works, which is necessarily for finding next
	//	adjacent table cells to insertBefore); then replace those empty data cells with
	//	the new data cells and insert/expand table cells into the grid table
	//	case 1: single-column cell: new cell
	//	case 2: part of multi-column cell, but not start of it: increment cell's colSpan
	//	case 3: start of multi-column cell: new cell
	var curDataRow, prevDataRow;
	//splice in empty data cells for each row
	for (var i = 0; i < data.numRows; i++) {
		data[i].splice(colIndex, 0, 0);
	}
	data.numCols++;
	//replace those empty data cells with new data cells and insert/expand table cells
	//	into the grid table
	for (var i = 0; i < data.numRows;) {
		prevDataRow = curDataRow;
		curDataRow = data[i];
		//curDataRow[colIndex + 1] is the current data cell, since an empty data cell was
		//	already spliced to curDataRow[colIndex]
		var curDataCell = curDataRow[colIndex + 1];
		//if new col intersects a multi-col cell, increment its colSpan
		//end cases: dataRow[-1] is undefined while dataRow[0] is not, and
		//	dataRow[dataRow.length] is undefined while dataRow[dataRow.length - 1] is not,
		//	so there are no cases of undefined == undefined
		//optimization: increment j by colSpan
		if (curDataCell == curDataRow[colIndex - 1]) {
			curDataCell.colSpan++;
			curDataCell.gridRef.colSpan++;
			curDataRow[colIndex] = curDataCell;
			i++;
			var rowSpan = curDataCell.rowSpan;
			for (var k = 1; k < rowSpan; k++)
				data[i++][colIndex] = curDataCell;
		//else if start of multi-col cell or single-col cell, insert new cell
		} else {
			//find next adjacent cell that is an actual table cell (rather than a grid cell,
			//	which is a table cell split among its spans) on the same row
			//unlike deleteRow (which shares code similar to here), j cannot be cached,
			//	since j must be recalculated for every row (j is the index of the cell
			//	to insert before)
			if (prevDataRow && curDataCell) {
				var j = colIndex + curDataCell.colSpan;
				while (j < data.numCols && curDataCell == prevDataRow[j]) {
					j += curDataCell.colSpan;
					curDataCell = curDataRow[j];
				}
			}
			var cell = document.createElement('td');
			//calc whether type is td or th
			var type = 'td';
			//if inserting before left-most column and layout mode involves left-most column
			if (colIndex == 0 && (layoutMode == 'Both' || layoutMode == 'Left')) {
				//change cells of the current left-most column to td
				if (curDataCell && curDataCell.type != 'td' && (i != 0 || layoutMode != 'Both')) {
					curDataCell.type = 'td';
					var className = curDataCell.gridRef.className;
					curDataCell.gridRef.className = className.substr(0, className.length - 2) + 'td';
				}
				//all new cells will be th
				type = 'th';
			//else if layout mode involves top row, only the top new cell will be th
			} else if (i == 0 && (layoutMode == 'Both' || layoutMode == 'Top')) {
				type = 'th';
			}
			var newDataCell = new DataCell(null, cell, type);
			cell.className = 'new-' + type;
			curDataRow.gridRef.insertBefore(cell, curDataCell? curDataCell.gridRef : null);
			curDataRow[colIndex] = newDataCell;
			i++;
		}
	}

	//recalc style
	updateHorzStyle();
}

function deleteCol() {
	//find index of targeted col
	var target = event.srcElement.parentNode;
	var colIndex;
	var editorRow = editorHorz.firstChild.firstChild;
	var editorRowCols = editorRow.childNodes;
	for (var i = 0; i < editorRowCols.length; i++) {
		if (editorRowCols[i] == target) {
			colIndex = (i - 1) / 2;
			break;
		}
	}

	//truncate editorHorz
	editorRow.removeChild(target.previousSibling);
	editorRow.removeChild(target);

	//remove col
	//strategy: remove/reduce cells for each row
	//	case 1: single-column cell: remove cell
	//	case 2: multi-column cell: decrement cell's colSpan
	for (var i = 0; i < data.numRows;) {
		var dataRow = data[i];
		var dataCell = dataRow[colIndex];
		//if multi-col cell, decrement its colSpan
		if (dataCell.colSpan > 1) {
			dataCell.colSpan--;
			dataCell.gridRef.colSpan--;
		//else if single-col cell, remove it
		} else {
			dataRow.gridRef.removeChild(dataCell.gridRef);
		}
		//remove data cells for each row span
		var rowSpan = dataCell.rowSpan;
		dataRow.splice(colIndex, 1);
		i++;
		for (var j = 1; j < rowSpan; j++)
			data[i++].splice(colIndex, 1);
	}
	data.numCols--;
	//if left-most column is being removed and layout mode involves left-most column,
	//	change the cells of the now left-most column to th
	if (data.numCols && colIndex == 0 && (layoutMode == 'Both' || layoutMode == 'Left')) {
		for (var i = 0; i < data.numRows;) {
			var dataCell = data[i][0];
			if (dataCell.type != 'th') {
				dataCell.type = 'th';
				var className = dataCell.gridRef.className;
				dataCell.gridRef.className = className.substr(0, className.length - 2) + 'th';
			}
			i += dataCell.rowSpan;
		}
	}

	//recalc style
	updateHorzStyle();
}

function insertRow() {
	//find index of targeted row
	var target = event.srcElement.parentNode.parentNode;
	var rowIndex;
	var editorCol = editorVert.firstChild;
	var editorColRows = editorCol.childNodes;
	for (var i = 0; i < editorColRows.length; i++) {
		if (editorColRows[i] == target) {
			rowIndex = i / 2;
			break;
		}
	}

	//expand editorVert
	editorCol.insertBefore(new InsertRowButton(), target);
	editorCol.insertBefore(new DeleteRowButton(), target);

	//insert new row before selected row
	//strategy: create new row and populate it, then splice it into data
	//	case 1: single-row cell: new cell
	//	case 2: part of multi-row cell, but not start of it: increment cell's rowSpan
	//	case 3: start of multi-row cell: new cell
	var row = document.createElement('tr');
	var newDataRow = new DataRow(null, row);
	var curDataRow = data[rowIndex];
	var prevDataRow = data[rowIndex - 1];
	var bFormerTopRowToTD = false;
	for (var i = 0; i < data.numCols;) {
		//if new row intersects a multi-row cell, increment its rowSpan
		if (prevDataRow && curDataRow) {
			var dataCell = curDataRow[i];
			if (dataCell == prevDataRow[i]) {
				dataCell.rowSpan++;
				dataCell.gridRef.rowSpan++;
				var colSpan = dataCell.colSpan;
				for (var j = 0; j < colSpan; j++)
					newDataRow[i++] = dataCell;
				continue;
			}
		}
		//else if start of multi-row cell or single-row cell, insert new cell
		var cell = document.createElement('td');
		//if inserting before top row and layout mode involves top row, all new cells will be th
		//	(the cells of the current top row will be changed to td later)
		//or if layout mode involves left-most column, only the left-most new cell will be th
		var type = ((rowIndex == 0 && (layoutMode == 'Both' || layoutMode == 'Top')) ||
					(i == 0 && (layoutMode == 'Both' || layoutMode == 'Left')))?
			'th' : 'td';
		var dataCell = new DataCell(null, cell, type);
		cell.className = 'new-' + type;
		row.appendChild(cell);
		newDataRow[i] = dataCell;
		i++;
	}
	//if inserting before top row and layout mode involves top row,
	//	change cells of the current top row to td
	if (rowIndex == 0 && (layoutMode == 'Both' || layoutMode == 'Top')) {
		if (data.numRows && data.numCols) {
			curDataRow = data[0];
			//if layout mode is 'both', skip the first cell
			var i = (layoutMode == 'Both')? curDataRow[0].colSpan : 0;
			while (i < curDataRow.length) {
				var curDataCell = curDataRow[i];
				if (curDataCell != 'td') {
					curDataCell.type = 'td';
					var className = curDataCell.gridRef.className;
					curDataCell.gridRef.className = className.substr(0, className.length - 2) + 'td';
				}
				i += curDataCell.colSpan;
			}
		}
	}
	var tbody = grid.tBodies[0];
	tbody.insertBefore(row, tbody.childNodes[rowIndex] || null);
	data.splice(rowIndex, 0, newDataRow);
	data.numRows++;

	//recalc style
	updateVertStyle();
}

function deleteRow() {
	//find index of targeted row
	var target = event.srcElement.parentNode.parentNode;
	var rowIndex;
	var editorCol = editorVert.firstChild;
	var editorColRows = editorCol.childNodes;
	for (var i = 0; i < editorColRows.length; i++) {
		if (editorColRows[i] == target) {
			rowIndex = (i - 1) / 2;
			break;
		}
	}

	//truncate editorVert
	editorCol.removeChild(target.previousSibling);
	editorCol.removeChild(target);

	//remove row
	//strategy: remove/reduce/move all cells from row, then remove row
	//	case 1 - single-row cell: remove cell
	//	case 2 - part of multi-row cell, but not start of it: decrement cell's rowSpan
	//	case 3 - start of multi-row cell: decrement cell's rowSpan and move cell down to next row
	var tbody = grid.tBodies[0];
	var row = tbody.rows[rowIndex];
	var curDataRow = data[rowIndex];
	var prevDataRow = data[rowIndex - 1];
	var nextDataRow = data[rowIndex + 1];
	var j = 0;
	var curDataCell, nextDataCell;
	for (var i = 0; i < data.numCols; i += curDataCell.colSpan) {
		curDataCell = curDataRow[i];
		//if multi-row cell, decrement its rowSpan
		if (curDataCell.rowSpan > 1) {
			curDataCell.rowSpan--;
			curDataCell.gridRef.rowSpan--;
			//don't need to splice out data cells, since they will be removed
			//	when the data row is removed
			//if start of multi-row cell, move that cell to the next row
			//don't need to check if next row exists (e.g. on last row),
			//	because any cells starting on the last row should never have rowSpan > 1
			if (!prevDataRow || curDataCell != prevDataRow[i]) {
				//find next adjacent cell that is an actual table cell (rather than a grid cell,
				//	which is a table cell split among its spans) on the next row
				//nextDataRow is assumed to exist,
				//	since otherwise the table HTML would be incorrect
				//optimization: don't need to recalc nextDataCell if i + 1
				//	still hasn't caught up to j, which is the index of the cell to insert before
				//optimization: increment j by colSpan;
				//	j should always be at the start of a cell, so incrementing by colSpan
				//	should be safe for first iteration
				if (i >= j) {
					j = i + curDataCell.colSpan;
					nextDataCell = nextDataRow[j];
					while (j < data.numCols && nextDataCell == curDataRow[j]) {
						j += nextDataCell.colSpan;
						nextDataCell = nextDataRow[j];
					}
				}
				//move cell
				row.nextSibling.insertBefore(curDataCell.gridRef, nextDataCell ? nextDataCell.gridRef : null);
			}
		}
	}
	tbody.removeChild(row);
	data.splice(rowIndex, 1);
	data.numRows--;
	//if top row is being removed and layout mode involves top column,
	//	change the cells of the now top row to th
	if (data.numRows && rowIndex == 0 && (layoutMode == 'Both' || layoutMode == 'Top')) {
		var topRow = data[0];
		for (var i = 0; i < topRow.length;) {
			var curDataCell = topRow[i];
			if (curDataCell.type != 'th') {
				curDataCell.type = 'th';
				var className = curDataCell.gridRef.className;
				curDataCell.gridRef.className = className.substr(0, className.length - 2) + 'th';
			}
			i += curDataCell.colSpan;
		}
	}

	//recalc style
	updateVertStyle();
}

//button UI/grid serialization funcs

function doOK() {
	doApply();
	window.returnValue = table;
	window.close();
}

function doApply() {
	//validate table properties
	var optionTable = document.getElementById('options');
	var options = document.getElementsByTagName('input');
	for (var i = 0; i < options.length; i++) {
		var option = options[i];
		var container = document.getElementById(option.id + 'Container');
		if (container && container.className == 'error') {
			alert(
				document.getElementById(option.id + 'Label').innerHTML + ' (' + option.id + ') ' +
				document.getElementById('INVALID_VALUE').innerHTML + ': "' + option.value + '"\n\n' +
				document.getElementById(option.id + 'Desc').innerHTML
			);
			option.focus();
			return;
		}
	}

	//create table if needed
	var tbody;
	if (table) {
		tbody = table.tBodies[0];
		if (!oldTable)
			oldTable = table.cloneNode(true);
	} else {
		table = doc.createElement('table');
		tbody = doc.createElement('tbody');
		table.appendChild(tbody);
	}

	//write table properties
	if (caption.value) {
		if (!table.caption)
			table.createCaption();
		table.caption.innerHTML = caption.value;
	} else {
		if (table.caption)
			table.deleteCaption();
	}
	if (summary.value)
		table.summary = summary.value;
	else
		table.removeAttribute('summary');
	if (cellSpacing.value)
		table.cellSpacing = cellSpacing.value;
	else
		table.removeAttribute('cellSpacing');
	if (cellPadding.value)
		table.cellPadding = cellPadding.value;
	else
		table.removeAttribute('cellPadding');
	if (border.value)
		table.border = border.value;
	else
		table.removeAttribute('border');
	if (width.value)
		table.width = width.value + '%';
	else
		table.removeAttribute('width');

	//remove all current rows from table
	while (tbody.firstChild)
		tbody.removeChild(tbody.firstChild);

	//"serialize" data into table
	//TODO: what if table has columns?
	var curDataRow, prevDataRow;
	for (var i = 0; i < data.numRows; i++) {
		prevDataRow = curDataRow;
		curDataRow = data[i];
		//reuse row if it's already available
		var row = curDataRow.tableRef;
		if (row) {
			//remove all current cells from reused row
			while (row.firstChild)
				row.removeChild(row.firstChild);
		} else {
			//otherwise, create a new row
			row = doc.createElement('tr');
		}
		var dataCell;
		for (var j = 0; j < data.numCols; j += dataCell.colSpan) {
			dataCell = curDataRow[j];
			if (prevDataRow && dataCell == prevDataRow[j])
				continue;
			var cell = dataCell.tableRef;
			//if cell was in the original table and is of the same type, reuse it
			if (!cell) {
				cell = doc.createElement(dataCell.type);
				cell.innerHTML = '&nbsp;';
			} else if (cell.type != dataCell.type) {
				var oldCell = cell;
				cell = doc.createElement(dataCell.type);
				//following is slow but is the only way to ensure everything is copied,
				//	including innerHTML, attributes, and custom properties
				for (var p in oldCell) {
					try {
						cell[p] = oldCell[p];
					} catch (e) {}
				}
			}
			cell.colSpan = dataCell.colSpan;
			cell.rowSpan = dataCell.rowSpan;
			row.appendChild(cell);
		}
		tbody.appendChild(row);
	}

	//if a range was passed to this window instead of a table,
	//	overwrite the range with the new table
	if (range) {
		if (range.length) {
			var item = range(0);
			item.parentNode.replaceChild(table, item);
		} else {
			oldHtmlText = range.htmlText;
			range.pasteHTML(table.outerHTML);
			var curParent = range.parentElement(), parent;
			do {
				range.moveStart('character', -1);
				parent = range.parentElement();
			} while (curParent == parent);
			while (parent.tagName.toLowerCase() != 'table')
				parent = parent.parentNode;
			table = parent;
		}
		range = null;
	}
}

function doCancel() {
	window.returnValue = null;
	window.close();
}

window.onunload = function() {
	if (!window.returnValue) {
		//check oldHtmlText first before checking oldTable,
		//	since both oldHtmlText and oldTable can be set
		if (oldHtmlText != null) {
			//don't rely on builtin undo, since it isn't reliable
			table.outerHTML = oldHtmlText;
		} else if (oldTable) {
			table.parentNode.replaceChild(oldTable, table);
		}
	}
}

document.onkeypress = function() {
	switch (event.keyCode) {
		case 27:	//esc
			doCancel();
			break;
		case 13:	//enter
			doSave();
			break;
	}
}
