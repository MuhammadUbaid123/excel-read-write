var express = require('express');
var router = express.Router();
const XLSX = require('xlsx');
/* GET home page. */
router.get('/', function(req, res, next) {
  
  const workbook = XLSX.readFile('./excel-file.xlsx');
  // console.log(workbook);
  // const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  // Select the worksheet named "master"
  const sheetName = 'Master Sheet';
  const worksheet = workbook.Sheets[sheetName];

  // Convert worksheet to JSON format
  const data = XLSX.utils.sheet_to_json(worksheet);
  console.log(data);
    res.render('index', { data: data });
  });

module.exports = router;
