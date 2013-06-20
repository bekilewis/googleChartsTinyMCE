<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=10" />
	<title>Insert a Graph</title>
	<script type="text/javascript" src="http://code.jquery.com/jquery-1.8.3.min.js"></script>
	<script type="text/javascript" src="/keypoint/database/scripts/jquery.contextMenu.js"></script>
	<script type="text/javascript" src="/keypoint/database/scripts/jquery.ui.position.js"></script>
	<script type="text/javascript" src="js/jquery.handsontable.full.js"></script>
	<script type="text/javascript" src="/keypoint/database/js/tinymce/plugins/compat3x/tiny_mce_popup.js"></script>
	<script type="text/javascript" src="js/dialog.js"></script>
	<script type="text/javascript" src="https://www.google.com/jsapi"></script>
	<link rel="stylesheet" type="text/css" href="/keypoint/database/scripts/jquery.contextMenu.css" />
	<!-- <link rel="stylesheet" type="text/css" href="http://sysdev/keypoint/database/js/tinymce/themes/advanced/skins/default/dialog.css" /> -->
	<link rel="stylesheet" type="text/css" href="/keypoint/database/styles/handsontable.css" />
	
</head>
<body onload="setUpDataTable();">
<%	
	set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open "<path to database>"

	set rs=Server.CreateObject("ADODB.recordset")
	set rs2=Server.CreateObject("ADODB.recordset")
	
	if Request.querystring("save") = "true" then
		rs.Open "SELECT graphID FROM tblGraph WHERE graphID = " & Request.querystring("id"), conn
		if rs.EOF then
			sqlStr =  "INSERT INTO tblGraph (reportID, rowData, colData, graphName, graphType, hAxisLabel, vAxisLabel) VALUES (" & Request.Cookies("auditID") & ", '" & Request.Cookies("rowString") & "', '" & Request.Cookies("colString") & "', '" & Request.Cookies("graphName") & "', '" & Request.Cookies("graphType") & "', '" & Request.Cookies("hAxisLabel") & "', '" & Request.Cookies("vAxisLabel") & "');"
			rs2.Open sqlStr, conn
		else
			sqlStr =  "UPDATE tblGraph SET rowData = '" & Request.Cookies("rowString") & "', colData = '" & Request.Cookies("colString") & "', graphName = '" & Request.Cookies("graphName") & "', graphType = '" & Request.Cookies("graphType") & "', hAxisLabel = '" & Request.Cookies("hAxisLabel") & "', vAxisLabel = '" & Request.Cookies("vAxisLabel") & "' WHERE graphID = " & Request.querystring("id") & ";"
			rs2.Open sqlStr, conn
		end if
		rs.close
	end if
	
	currID = 0
	loadInt = 0
	
	if Request.querystring("id") > 0 then
		loadInt = 1
		
		rs.Open "SELECT * FROM tblGraph WHERE graphID = " & Request.querystring("id"), conn
		
		rowString = rs.Fields("rowData").value
		colString = rs.Fields("colData").value
		graphType = rs.Fields("graphType").value
		graphName = rs.Fields("graphName").value
		hAxisLabel = rs.Fields("hAxisLabel").value
		vAxisLabel = rs.Fields("vAxisLabel").value
		currID = rs.Fields("graphID").value
		rs.close
	else	
		rs.Open "SELECT graphID FROM tblGraph ORDER BY graphID DESC", conn
		currID = rs.Fields("graphID").value + 1
		rs.close
	end if
	
	if Request.querystring("save") = "true" then
	%>
		<script>
		GoogleChartDialog.insert('<%= currID %>');
		</script>
	<%
	end if	
	
%>

<script type="text/javascript">
<% 	if Request.querystring("save") = "true" then %>
	tinyMCE.execCommand('Repaint');
<% end if %>
var currID = <%= currID %>
var loadInt = <%= loadInt %>
var textStr = tinyMCEPopup.editor.selection.getContent({format : 'html'});
//var textStr = tinyMCE.activeEditor.selection.getNode().nodeName;
//alert(textStr);
<% 	if not(Request.querystring("id") >0) then %>
if (textStr == "" ){
} else {
	var textArr = textStr.split("id=");
	var textID = textArr[1].split("\"");
	currID = textID[0];
	loadInt = 1;
	window.location = "dialog.asp?id=" + currID;
}
<% end if %>

function setCookie(c_name,value,exdays) {
	var exdate=new Date();
	exdate.setDate(exdate.getDate() + exdays);
	var c_value=escape(value) + ((exdays==null) ? "" : "; expires="+exdate.toUTCString());
	document.cookie=c_name + "=" + c_value;
}

function setUpDataTable() {
	var data = [
		[" ", "Kia", "Nissan", "Toyota", "Honda"],
		["2008", 10, 11, 12, 13],
		["2009", 20, 11, 14, 13],
		["2010", 30, 15, 12, 13]
	];

	$("#exampleGrid").handsontable({
		data: data,
		minRows: 15,
		minCols: 10,
		minSpareCols: 1,
		minSpareRows: 1,
		height: 300,
		width: 500,
		rowHeaders: true,
		colHeaders: true,
		colWidths: [70, 70, 70, 70, 70],
		manualColumnMove: true,
		manualColumnResize: true,
		stretchH: 'all',
		scrollH: 'auto',
		scrollV: 'auto',
		contextMenu: true
	});
	loadData();
}

function loadData() {

//alert(currID);
//alert(loadInt);
<% 	if loadInt > 0 then %>
	
	$("#exampleGrid").handsontable('clear');

	var rowString = '<%= rowString %>';
	var colString = '<%= colString %>';
	var graphType = '<%= graphType %>';
	var graphName = '<%= graphName %>';
	var hAxisLabel = '<%= hAxisLabel %>';
	var vAxisLabel = '<%= vAxisLabel %>';
	
	if (graphType == "Column") {
		document.getElementById("DropdownList3").selectedIndex = 0;
	} else if (graphType == "Pie") {
		document.getElementById("DropdownList3").selectedIndex = 1;
	} else if (graphType == "Line") {
		document.getElementById("DropdownList3").selectedIndex = 2;
	} else if (graphType == "Area") {
		document.getElementById("DropdownList3").selectedIndex = 3;
	} else if (graphType == "Bar") {
		document.getElementById("DropdownList3").selectedIndex = 4;
	} else {
		document.getElementById("DropdownList3").selectedIndex = 0;
	}
	
	document.getElementById("graphName").value = graphName;
	document.getElementById("hAxisLabel").value = hAxisLabel;
	document.getElementById("vAxisLabel").value = vAxisLabel;
	
	var tempColString = colString.split("}");
	var tempRowString = rowString.split("\"c\"");
	
	for (c=0; c<tempColString.length; c++){
		tempString = tempColString[c];
		tempStringArray = tempString.split("\"");							
		if ((tempStringArray[7] == " undefined") || (tempStringArray[7] == "undefined")) {
			tempStringArray[7] = "";
		}
		$("#exampleGrid").handsontable('setDataAtCell', 0, c, tempStringArray[7])
	}
	
	for (r=0; r<tempRowString.length; r++){
		tempString = tempRowString[r];
		tempStringArray = tempString.split(":");
		for (c=0; c<tempStringArray.length; c++){
			if (c<=1){
			} else {
				cellStringArray = tempStringArray[c].split("}");
				cellString = cellStringArray[0].replace(/\"/g,"");
				//cellString = cellString.replace(/ /g,"");
				if ((cellString == " undefined") || (cellString == "undefined")) {
					cellString = "";
				}
				if (cellString == "null") {
					//cellString = "";
				} else {
					$("#exampleGrid").handsontable('setDataAtCell', r, c-2, cellString);
				}
			}
		}
	}
<% end if %>
}


// Load the Visualization API and the piechart package.
google.load('visualization', '1.0', {'packages':['corechart']});

// Set a callback to run when the Google Visualization API is loaded.
google.setOnLoadCallback(drawVisualization);

function drawVisualization() {

	rowCount = $("#exampleGrid").handsontable('countRows');
	colCount = $("#exampleGrid").handsontable('countCols');
	
	colString = "[";
	rowString = "[";
	for (i=1;i<=colCount;i++){
		if (($("#exampleGrid").handsontable('getDataAtCell', 0, i-1) == "") || ($("#exampleGrid").handsontable('getDataAtCell', 0, i-1) == null)) {
		} else {
			if (colString != "[") {
				colString = colString + ",";
			}
			if (i==1) {
				colString = colString + "{\"id\": \"" + i + "\", \"label\": \"" + $("#exampleGrid").handsontable('getDataAtCell', 0, i-1) + "\", \"type\": \"string\"}";
			} else {
				colString = colString + "{\"id\": \"" + i + "\", \"label\": \"" + $("#exampleGrid").handsontable('getDataAtCell', 0, i-1) + "\", \"type\": \"number\"}";
			}
		}
	}
	colString = colString + "]";

	for (r=2;r<=rowCount;r++){
		if (($("#exampleGrid").handsontable('getDataAtCell', r-1, 0) == null)||($("#exampleGrid").handsontable('getDataAtCell', r-1, 0) == "")) {
		} else {
			if (rowString != "[") {
				rowString = rowString + ",";
			}
			rowString = rowString + "{\"c\":["
			for (c=1;c<=colCount;c++){
				if (c==1) {
					rowString = rowString + "{\"v\": \"" + $("#exampleGrid").handsontable('getDataAtCell', r-1, c-1) + "\"}";
				} else {
					if ($("#exampleGrid").handsontable('getDataAtCell', r-1, c-1) == "") {
					} else {
						rowString = rowString + "{\"v\": " + $("#exampleGrid").handsontable('getDataAtCell', r-1, c-1) + "}";
					}
				}
				if (c!= colCount) {
					rowString = rowString + ",";
				}
			}
			rowString = rowString + "]}";
		}
	}
	rowString = rowString + "]";
	
	colString=colString.replace(/undefined/g,"null");
	rowString=rowString.replace(/undefined/g,"null");
	
	setCookie("colString", colString, 1);
	setCookie("rowString", rowString, 1);
	
	//graphType = document.getElementById("graphType").InnerText;
	var graphName = document.getElementById("graphName").value;
	var hAxisLabel = document.getElementById("hAxisLabel").value;
	var vAxisLabel = document.getElementById("vAxisLabel").value;
	
	var dataTab = "{\"cols\": " + colString + ", \"rows\": " + rowString + "}";
	var data = new google.visualization.DataTable(dataTab, 0.6);

	//Create and draw the visualization.
	htitle = hAxisLabel
	vtitle = vAxisLabel
	var options = {
		title: graphName,
		titleTextStyle: {color: 'black', fontName: 'Arial', fontSize: 18},
		width: 680,
		height: 380,
		fontsize: 14,
		vAxis: {title: vtitle,  titleTextStyle: {color: 'black', fontName: 'Arial', fontSize: 16}, textStyle: {color: 'black', fontName: 'Arial', fontSize: 11}},
		hAxis: {title: htitle,  titleTextStyle: {color: 'black', fontName: 'Arial', fontSize: 16}, textStyle: {color: 'black', fontName: 'Arial', fontSize: 11}}
	};
	var x=document.getElementById("DropdownList3").selectedIndex;
	var y=document.getElementById("DropdownList3").options;
	selected = y[x].text;
	
	setCookie("graphType", selected, 1);
	
	//selected = graphType;
	if (selected == "Table") {
		new google.visualization.Table(document.getElementById('graphHolder')).draw(data,options);
	} else if (selected == "Column") {
		new google.visualization.ColumnChart(document.getElementById('graphHolder')).draw(data,options);
	} else if (selected == "Pie") {
		new google.visualization.PieChart(document.getElementById('graphHolder')).draw(data,options);
	} else if (selected == "Line") {
		new google.visualization.LineChart(document.getElementById('graphHolder')).draw(data,options);
	} else if (selected == "Area") {
		new google.visualization.AreaChart(document.getElementById('graphHolder')).draw(data,options);
	} else if (selected == "Bar") {
		new google.visualization.BarChart(document.getElementById('graphHolder')).draw(data,options);
	} else {
		new google.visualization.ColumnChart(document.getElementById('graphHolder')).draw(data,options);
	} 
	
	setCookie("graphName", document.getElementById("graphName").value, 1);
	setCookie("hAxisLabel", document.getElementById("hAxisLabel").value, 1);
	setCookie("vAxisLabel", document.getElementById("vAxisLabel").value, 1);
}
function saveAndInsert() {
	drawVisualization();
	//GoogleChartDialog.insert('<%= currID %>');
}
</script>

<div id="exampleGrid" class="dataTable" style="height: 300px; width: 500px; overflow: scroll; float:left; padding: 15px;"></div>

<div style="float:right;">
	Graph type:<select id="DropdownList3">
	<option selected="selected" value="Column">Column</option>
	<option value="Pie">Pie</option>
	<option value="Line">Line</option>
	<option value="Area">Area</option>
	<option value="Bar">Bar</option>
	</select>
	<br />
	Graph Title:<input type="text" id="graphName" /><br />
	x-Axis Label:<input type="text" id="hAxisLabel" /><br />
	y-Axis Label:<input type="text" id="vAxisLabel" /><br />
	<button type="button" id="draw" name="draw" value="Draw" onclick="drawVisualization();" >Draw</button><br />
	<form action="dialog.asp" method="get">
		<input name="save" value="true" type="hidden" />
		<input name="id" value="<%= currID %>" type="hidden" />
		<div class="mceActionPanel">
			<input type="submit" id="testinsert" value="Save & Insert" onclick="saveAndInsert();" />
		</div>
		<br />
	</form>	
</div>
<div class="clearboth"></div>
<span id="graphHolder" style="float:left;"></span>

</body>
</html>
