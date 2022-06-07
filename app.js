const express = require('express');
const app = express();
const port = 3000;
const path = require('path');

const XLSXChart = require('xlsx-chart');
const xlsxChart = new XLSXChart();
const opts = {
    file: "chart.xlsx",
	chart: "column",
	titles: [
		"Title 1",
		"Title 2",
		"Title 3"
	],
	fields: [
		"Field 1",
		"Field 2",
		"Field 3",
		"Field 4"
	],
	data: {
		"Title 1": {
			"Field 1": 5,
			"Field 2": 10,
			"Field 3": 15,
			"Field 4": 20 
		},
		"Title 2": {
			"Field 1": 10,
			"Field 2": 5,
			"Field 3": 20,
			"Field 4": 15
		},
		"Title 3": {
			"Field 1": 20,
			"Field 2": 15,
			"Field 3": 10,
			"Field 4": 5
		}
	}
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
          "Content-Disposition": "attachment; filename=chart.xlsx",
          "Content-Length": data.length
        });
        res.status (200).send (data);
    });
});

app.listen(port, () => {
    console.log(`http://localhost:${port}`);
});