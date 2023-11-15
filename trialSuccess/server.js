const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const port = 3001;

// Set up the storage for multer
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Serve static files from the "public" directory
app.use(express.static('public'));
// const cors = require('cors');
// app.use(cors());
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});


app.post('/upload', upload.single('excelFile'), async (req, res) => {
  const arrayBuffer = req.file.buffer;

  try {
    function gg(table_name){
      const tableReference = worksheet.tables[table_name];
    const tableRange = tableReference.table.tableRef;
    
    // Extract starting and ending cell references from the tableRange
    const [startCell, endCell] = tableRange.split(':');
    
    // Convert cell references to row and column indices
    const startRowIndex = parseInt(startCell.match(/\d+/)[0], 10);
    const endRowIndex = parseInt(endCell.match(/\d+/)[0], 10);
    const startColumnIndex = startCell.match(/[A-Z]+/)[0];
    const endColumnIndex = endCell.match(/[A-Z]+/)[0];
    
    // Convert column letters to numerical indices
    const startColumnNum = columnToNumber(startColumnIndex);
    const endColumnNum = columnToNumber(endColumnIndex);
    
    // Function to convert column letter to numerical index
    function columnToNumber(column) {
      let result = 0;
      for (let i = 0; i < column.length; i++) {
        result = result * 26 + column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
      }
      return result;
    }
    
    // Convert numerical column index to letter
    function numberToColumn(number) {
      let result = '';
      while (number > 0) {
        const remainder = (number - 1) % 26;
        result = String.fromCharCode('A'.charCodeAt(0) + remainder) + result;
        number = Math.floor((number - 1) / 26);
      }
      return result;
    }
    
    // Initialize JSON object
    const resultJson = {
      [`${tableReference.name}_columns`]: [],
      [`${tableReference.name}_data`]: [],
    };
    
    // Populate column names from the first row
    for (let col = startColumnNum; col <= endColumnNum; col++) {
      const colLetter = numberToColumn(col);
      const cellReference = `${colLetter}${startRowIndex}`;
      const cellValue = worksheet.getCell(cellReference).value;
      resultJson[`${tableReference.name}_columns`].push(cellValue);
    }
    
    // Iterate through rows and columns to get table data
    for (let row = startRowIndex + 1; row <= endRowIndex; row++) {
      const rowData = {};
      for (let col = startColumnNum; col <= endColumnNum; col++) {
        const colLetter = numberToColumn(col);
        const cellReference = `${colLetter}${row}`;
        const cellValue = worksheet.getCell(cellReference).value;
        rowData[resultJson[`${tableReference.name}_columns`][col - startColumnNum]] = cellValue;
      }
      resultJson[`${tableReference.name}_data`].push(rowData);
    }
    
    // Log the resulting JSON object
    console.log(JSON.stringify(resultJson, null, 2));
    // res.json(JSON.stringify(resultJson, null, 2));
    return resultJson;
    }
    const workbook = new ExcelJS.Workbook();
    
    await workbook.xlsx.load(arrayBuffer);
    // console.log(workbook);
    // Assuming the table is on the first sheet
    const worksheet = workbook.getWorksheet(1);
    let TableArray = ['Table1','Table2','Table3','Table4','Table5'];
    let resp;
    let array=[];
for (let i = 0; i < TableArray.length; i++) {
  resp=gg(TableArray[i]);
  array.push(resp);
  }
  res.json(array);
  } catch (error) {
    console.error('Error:', error.message);
  }
});

app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});


// console.log(worksheet.tables['Table1'].table.columns)
    // const rows = worksheet.tables['Table1'].table.rows;
    // const columns = worksheet.tables['Table1'].table.columns;
    // console.log(rows)
    // console.log(worksheet.tables['Table1']);
//     const tableReference = worksheet.tables['Table2'];
// const tableRange = tableReference.table.tableRef;

// // Extract starting and ending cell references from the tableRange
// const [startCell, endCell] = tableRange.split(':');

// // Convert cell references to row and column indices
// const startRowIndex = parseInt(startCell.match(/\d+/)[0], 10);
// const endRowIndex = parseInt(endCell.match(/\d+/)[0], 10);
// const startColumnIndex = startCell.match(/[A-Z]+/)[0];
// const endColumnIndex = endCell.match(/[A-Z]+/)[0];

// // Convert column letters to numerical indices
// const startColumnNum = columnToNumber(startColumnIndex);
// const endColumnNum = columnToNumber(endColumnIndex);

// // Function to convert column letter to numerical index (e.g., A to 1, B to 2, ...)
// function columnToNumber(column) {
//   let result = 0;
//   for (let i = 0; i < column.length; i++) {
//     result = result * 26 + column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
//   }
//   return result;
// }

// // Iterate through rows and columns to get table data
// for (let row = startRowIndex; row <= endRowIndex; row++) {
//   for (let col = startColumnNum; col <= endColumnNum; col++) {
//     // Convert numerical column index back to letter
//     const colLetter = numberToColumn(col);
//     const cellReference = `${colLetter}${row}`;
//     const cellValue = worksheet.getCell(cellReference).value;
//     console.log(`Row: ${row}, Column: ${colLetter}, Value: ${cellValue}`);
//   }
// }

// // Function to convert numerical column index to letter (e.g., 1 to A, 2 to B, ...)
// function numberToColumn(number) {
//   let result = '';
//   while (number > 0) {
//     const remainder = (number - 1) % 26;
//     result = String.fromCharCode('A'.charCodeAt(0) + remainder) + result;
//     number = Math.floor((number - 1) / 26);
//   }
//   return result;
// }


// gg('Table2')
    // columns.forEach((column, index) => {
    //   console.log(`Column ${index + 1}:`);
    //   console.log(`Name: ${column.name}`);
    //   console.log(`Totals Row Label: ${column.totalsRowLabel}`);
    //   console.log(`Totals Row Function: ${column.totalsRowFunction}`);
    //   console.log(`Filter Button: ${column.filterButton}`);
    //   console.log(`Totals Row Shown: ${column.totalsRowShown}`);
    //   console.log('------------------------');
    // });
//     const tableReference = worksheet.tables['Table1'];
// const columns = tableReference.table.columns;
// const rows = tableReference.table.rows;

// // Example: Log the data in the console
// rows.forEach((row, rowIndex) => {
//   console.log(`Row ${rowIndex + 1}:`);
//   columns.forEach((column, columnIndex) => {
//     const columnName = column.name;
//     const cellValue = row.getCell(columnName).value;
//     console.log(`${columnName}: ${cellValue}`);
//   });
//   console.log('------------------------');
// });

    // Replace 'YourTableName' with the actual name of your table
//     const tableName = 'Release_details';

//     // Find the table by name
//     // Replace the line with .find with a loop
// let table;
// for (const t of worksheet.tables) {
//   if (t.name === tableName) {
//     table = t;
//     break;
//   }
// }

// if (table) {
//   // Continue with the rest of your code for processing the table
//   const tableRange = table.getSheet().getRange(table.address);

//   // Extract data from the table
//   const tableData = [];
//   tableRange.eachCell({ includeEmpty: false }, (cell, rowNumber, colNumber) => {
//     if (!tableData[rowNumber]) {
//       tableData[rowNumber] = [];
//     }
//     tableData[rowNumber][colNumber] = cell.value;
//   });

//   // Display or do something with the tableData
//   displayTableData(tableData);
// } else {
//   console.log(`Table "${tableName}" not found.`);
// }


    // if (table) {
    //   // Get the data body range of the table
    //   const tableRange = table.getSheet().getRange(table.address);

    //   // Extract data from the table
    //   const tableData = [];
    //   tableRange.eachCell({ includeEmpty: false }, (cell, rowNumber, colNumber) => {
    //     if (!tableData[rowNumber]) {
    //       tableData[rowNumber] = [];
    //     }
    //     tableData[rowNumber][colNumber] = cell.value;
    //   });

    //   // Send the table data as a JSON response
    //   res.json({ success: true, data: tableData });
    // } else {
    //   res.json({ success: false, message: `Table "${tableName}" not found.` });
    // }
    // res.json()