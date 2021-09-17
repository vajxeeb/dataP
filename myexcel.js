
// //let data = xlsx.utils.sheet_to_json(ws)


// //let workbook=xlsx.readFile('test.xlsx')

// const { Workbook } = require('exceljs');
const XLSX = require('xlsx');
const workbook = XLSX.readFile('test.xlsx');
const WSH = workbook.SheetNames[0];


const installment = 'D1';
const date = 'F1';
//// ........get value from cell.........
let worksheet = workbook.Sheets[WSH]; 

const desired_cell_installment = worksheet[installment];
const installment_value = (desired_cell_installment ? desired_cell_installment.v : undefined);

const desired_cell_date = worksheet[date];
const date_value = (desired_cell_date ? desired_cell_date.v : undefined);

// worksheet['M3'].v = 'ttt'

//console.log('Installment is: '+installment_value);
//console.log('Date is: '+date_value);

// XLSX.writeFile(Workbook, "test.xlsx");


////.....retrive data from excel................
// const columnA = [];
const columnB = [];
// const columnC = [];
// const columnD = [];
// const columnE = [];
// const columnF = [];

for (let z in worksheet) {
  // if(z.toString()[0] === 'A'){
  //   columnA.push(worksheet[z].v);
  //   columnA.splice(0,2)
  //   columnA.splice(columnA.length-2,2)
  // }
  if(z.toString()[0] === 'B'){
    columnB.push(worksheet[z].v);
    
  }
  // if(z.toString()[0] === 'C'){
  //   columnC.push(worksheet[z].v);
    
  // }
  // if(z.toString()[0] === 'D'){
  //   columnD.push(worksheet[z].v);
    
  // }
  // if(z.toString()[0] === 'E'){
  //   columnE.push(worksheet[z].v);
    
  // }
  // if(z.toString()[0] === 'F'){
  //   columnF.push(worksheet[z].v);
    
  // }
}

 
columnB.splice(0,3)
columnB.splice(columnB.length-2,2)
console.log("array: "+columnB)





////..........Update value in cell and then write the data back to exel.............................
const Excel = require('exceljs');
const WB = new Excel.Workbook();
WB.xlsx.readFile('test.xlsx')
    .then(function() {

        const worksheet = WB.getWorksheet(1);
        let row1 = worksheet.getRow(1);
        let column1 = worksheet.getColumn(1)
        //let row2 = worksheet.getRow(2);
        //let row3 = worksheet.getRow(3);

        

       
        row1.commit();

        return WB.xlsx.writeFile('test.xlsx');
    })














//var Excel = require('exceljs');
// var workbook = new Excel.Workbook();
// workbook.xlsx.readFile('test.xlsx')
//     .then(function (data) {
//         var worksheet = workbook.getWorksheet(1);

//         worksheet.eachRow(function (row, rowNumber) {

//             var cell = row.getCell(3).value;
//             if (!cell) {
//                 console.log(rowNumber);
//                 row.splice(rowNumber, 1);
//                 row.commit();
//             }


//         });

//         return workbook.xlsx.writeFile('new.xlsx');

//     });




