<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<script type="text/javascript" src="https://www.google.com/jsapi"></script>
	<script type="text/javascript" src="http://code.jquery.com/jquery-1.8.3.min.js"></script>
</head>
<body>
<%	
	set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open "<path to database>"

	set rs=Server.CreateObject("ADODB.recordset")
	
	if Request.querystring("id") > 0 then
		rs.Open "SELECT * FROM tblGraph WHERE graphID = " & Request.querystring("id"), conn
		
		rowString = rs.Fields("rowData").value
		colString = rs.Fields("colData").value
		graphType = rs.Fields("graphType").value
		graphName = rs.Fields("graphName").value
		hAxisLabel = rs.Fields("hAxisLabel").value
		vAxisLabel = rs.Fields("vAxisLabel").value
		
		rs.close
		
	end if
%>
<script type="text/javascript">
// Load the Visualization API and the piechart package.
google.load('visualization', '1.0', {'packages':['corechart']});

// Set a callback to run when the Google Visualization API is loaded.
google.setOnLoadCallback(drawVisualization);

function drawVisualization() {

<% 	if Request.querystring("id") > 0 then %>
	rowString = '<%= rowString %>';
	colString = '<%= colString %>';
	graphType = '<%= graphType %>';
	graphName = '<%= graphName %>';
	hAxisLabel = '<%= hAxisLabel %>';
	vAxisLabel = '<%= vAxisLabel %>';
		
	colString=colString.replace(/undefined/g,"null");
	rowString=rowString.replace(/undefined/g,"null");
	//alert(rowString);

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
		fontsize: 12,
		vAxis: {title: vtitle,  titleTextStyle: {color: 'black', fontName: 'Arial', fontSize: 16}, textStyle: {color: 'black', fontName: 'Arial', fontSize: 11}},
		hAxis: {title: htitle,  titleTextStyle: {color: 'black', fontName: 'Arial', fontSize: 16}, textStyle: {color: 'black', fontName: 'Arial', fontSize: 11}}
	};

	selected = graphType;
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
			
<% else %>		
		
    var data = new google.visualization.DataTable();
    data.addColumn('string', 'Year');
    data.addColumn('number', 'Score');
    data.addRows([
      ['2005',3.6],
      ['2006',4.1],
      ['2007',3.8],
      ['2008',3.9],
      ['2009',4.6],
    ]);
	var options = {
		title: graphName,
		titleTextStyle: {color: 'black', fontName: 'Arial', fontSize: 18},
		width: 680,
		height: 380,
		fontsize: 14,
		vAxis: {title: vtitle,  titleTextStyle: {color: 'black', fontName: 'Arial', fontSize: 16}, textStyle: {color: 'black', fontName: 'Arial', fontSize: 11}},
		hAxis: {title: htitle,  titleTextStyle: {color: 'black', fontName: 'Arial', fontSize: 16}, textStyle: {color: 'black', fontName: 'Arial', fontSize: 11}}
	};
	
	new google.visualization.ColumnChart(document.getElementById('graphHolder')).draw(data,options);		
	
<% end if %>
}
</script>

<span id="graphHolder">graph</span>

</body>
</html>

