const path = require('path');
const csv = require('fast-csv');
const fs = require('fs');
const Excel = require('exceljs');

const rootDir = path.dirname(require.main.filename);

const openHome = (req, res) => {
  res.sendFile(path.join(rootDir, 'views', 'index.html'));
};

const readExcel = async (req, res) => {
  return res.send({
    msg: 'Excel file has been read correctly',
  });
};

const convertExcelToCSV = async (req, res) => {
  const dataWorkBook = new Excel.Workbook();
  const workbook = await dataWorkBook.xlsx.readFile(
    path.join(rootDir, 'data.xlsx')
  );
  await workbook.csv.writeFile(path.join(rootDir, 'converted.csv'), {
    encoding: 'UTF-8',
    writeBOM: true,
  });
  return res.send({
    msg: 'Excel file has been successfully converted ino CSV',
  });
};

const readCSV = (req, res) => {
  const stream = fs.createReadStream(path.join(rootDir, 'data.csv'));
  const csvData = [];
  let dataRowCount = 0;
  csv
    .parseStream(stream, { headers: true, encoding: 'utf-8' })
    .on('error', (err) =>
      res.send({
        msg: err,
        data: csvData,
        rowCount: dataRowCount,
      })
    )
    .on('data', (row) => {
      csvData.push(row);
    })
    .on('end', (rowCount) => {
      dataRowCount = rowCount;
      return res.send({
        msg: 'CSV file has been read successfully.',
        data: csvData,
        rowCount: dataRowCount,
      });
    });
};

const writeCSV = async (req, res) => {
  const csvStream = csv.format({ headers: true, writeBOM: true });
  const writeStream = fs.createWriteStream(
    path.join(rootDir, 'written_from_excel.csv')
  );
  csvStream.pipe(writeStream);
  const dataWorkBook = new Excel.Workbook();
  const workbook = await dataWorkBook.xlsx.readFile(
    path.join(rootDir, 'data.xlsx'),
    { headers: 1 }
  );
  const worksheet = workbook.getWorksheet(1);
  const headerRow = worksheet.getRow(1);
  const columnHeaders = [];
  headerRow.eachCell((cell, colNum) => {
    columnHeaders.push({ header: cell.value, key: cell.value });
  });
  worksheet.spliceRows(0, 1);
  worksheet.eachRow((row, rowNum) => {
    const item = {};
    row.eachCell((cell, colNum) => {
      item[columnHeaders[colNum]?.header] = cell.value;
    });
    csvStream.write(item);
  });
  csvStream.end();
  return res.send({
    msg: 'check the project folder for written_from_excel.csv file',
  });
};

module.exports = {
  readExcel,
  convertExcelToCSV,
  readCSV,
  openHome,
  writeCSV,
};
