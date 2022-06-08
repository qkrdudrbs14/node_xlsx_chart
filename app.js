const express = require('express');
const app = express();
const port = 3000;

const XLSXChart = require('xlsx-chart');
const xlsxChart = new XLSXChart();

let _fields = [];
let _kwh = {"chart": "column"};
let _radiation = {"chart": "line"};

for(let i=5; i < 22; i++) {
	_fields.push(i);
	_kwh[i] = Math.floor(Math.random() * 100);
	_radiation[i] = Math.floor(Math.random() * 50);
}

const opts = {
    file: "chart.xlsx",
	titles: [
		"kWh",
		"radiation"
	],
	fields:_fields,
	data: {
		"kWh": _kwh,
		"radiation": _radiation,
	},
	chartTitle: "Column and Line chart"
}

app.get("/", (req, res)=>{
    console.log("/");
    res.sendFile(__dirname + "/html/index.html");
});

app.get("/excel", (req, res) => {
    console.log("/excel");
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