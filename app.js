const express = require('express');
const app = express();
const port = 3000;
const path = require('path');

const XLSXChart = require('xlsx-chart');
const xlsxChart = new XLSXChart();
const opts = {
    file: "chart.xlsx",
	// chart: "column",
    chart: "line",
	titles: [
		"Title 1"
	],
	fields: [
		"Apple",
		"Blackberry",
		"Strawberry",
		"Cowberry"
	],
	data: {
		"Price": {
			"Apple": 10,
			"Blackberry": 5,
			"Strawberry": 15,
			"Cowberry": 20
		}
	},
	chartTitle: "Line chart"
}

app.get("/", (req, res)=>{
    console.log("/");
    res.sendFile(__dirname + "/html/index.html");
});

app.get("/excel", (req, res) => {
    console.log("/excel");
    // console.log(req.query.sl_id);
    xlsxChart.generate (opts, function (err, data) {
        res.set ({
          "Content-Type": "application/vnd.ms-excel",
          "Content-Disposition": "attachment; filename=line.xlsx",
          "Content-Length": data.length
        });
        res.status (200).send (data);
    });
});

app.listen(port, () => {
    console.log(`http://localhost:${port}`);
});