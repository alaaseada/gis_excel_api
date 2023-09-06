const express = require('express');
const Router = express.Router();
const {
  readExcel,
  convertExcelToCSV,
  readCSV,
  openHome,
  writeCSV,
} = require('../Controllers/mainController');

Router.route('/').get(openHome);
Router.route('/excel').get(readExcel);
Router.get('/convert', convertExcelToCSV);
Router.get('/csv', readCSV);
Router.get('/write_csv', writeCSV);

module.exports = Router;
