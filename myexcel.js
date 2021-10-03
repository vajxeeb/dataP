
let XLSX = require('xlsx');
let workbook = XLSX.readFile('test.xlsx');
let WSH = workbook.SheetNames[0];
let worksheet = workbook.Sheets[WSH];

const  getJsDateFromExcel  = require("excel-date-to-js") 

const installment = 'D1';
const date = 'F1';

//// ........get date and installment.........

const desired_cell_installment = worksheet[installment];
const installment_value = (desired_cell_installment ? desired_cell_installment.v : undefined);
const desired_cell_date = worksheet[date];
const date_value = (desired_cell_date ? desired_cell_date.v : undefined);
   

  var date1 = new Date(Math. round((date_value - (25567 + 1)) * 86400 * 1000));
  var converted_date = date1. toISOString(). split('T')[0];

  
//getJsDateFromExcel(date_value)
//var d = Date.parse(date_value)
console.log("my date is : "+converted_date.toString())
////.....valaible for conatain lottery number value................
const columnA_Id = [];
const columnB_Sn1 = [];
const columnC_Sn2 = [];
const columnD_Sn3 = [];
const columnE_Sn4 = [];
const columnF_Sn5 = [];
const columnG_Sn6 = [];

const columnH_Pn1 = [];
const columnI_Pn2 = [];
const columnJ_Pn3 = [];
const columnK_Pn4 = [];
const columnL_Pn5 = [];
const columnM_Pn6 = [];

const level = [];

//push lottery number to array
for (let z in worksheet) {

  //get level 
  if (z.toString()[0] === 'A') {
    level.push(worksheet[z].v);
  }
  //get id
  if (z.toString()[0] === 'B') {
    columnA_Id.push(worksheet[z].v);
  }
  //Total sells
  if (z.toString()[0] === 'C') {
    columnB_Sn1.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'D') {
    columnC_Sn2.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'E') {
    columnD_Sn3.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'F') {
    columnE_Sn4.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'G') {
    columnF_Sn5.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'H') {
    columnG_Sn6.push(worksheet[z].v);

  }
  //Total pay
  if (z.toString()[0] === 'J') {
    columnH_Pn1.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'K') {
    columnI_Pn2.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'L') {
    columnJ_Pn3.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'M') {
    columnK_Pn4.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'N') {
    columnL_Pn5.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'O') {
    columnM_Pn6.push(worksheet[z].v);
  }
}

//delete title and subtitle
columnA_Id.splice(0, 2)
columnA_Id.splice(columnA_Id.length - 2, 2)

columnB_Sn1.splice(0, 2)
columnB_Sn1.splice(columnB_Sn1.length - 2, 2)

columnC_Sn2.splice(0, 2)
columnC_Sn2.splice(columnC_Sn2.length - 1, 1)

columnD_Sn3.splice(0, 2)
columnD_Sn3.splice(columnD_Sn3.length - 1, 1)

columnE_Sn4.splice(0, 2)
columnE_Sn4.splice(columnE_Sn4.length - 1, 1)


columnF_Sn5.splice(0, 1)
columnF_Sn5.splice(columnF_Sn5.length - 1, 1)

columnG_Sn6.splice(0, 0)
columnG_Sn6.splice(columnF_Sn5.length - 1, 1)

columnH_Pn1.splice(0, 1)
columnH_Pn1.splice(columnH_Pn1.length - 1, 1)

columnI_Pn2.splice(0, 1)
columnI_Pn2.splice(columnI_Pn2.length - 1, 1)

columnJ_Pn3.splice(0, 1)
columnJ_Pn3.splice(columnJ_Pn3.length - 1, 1)

columnK_Pn4.splice(0, 1)
columnK_Pn4.splice(columnK_Pn4.length - 1, 1)

columnL_Pn5.splice(0, 1)
columnL_Pn5.splice(columnL_Pn5.length - 1, 1)

columnM_Pn6.splice(0, 1)
columnM_Pn6.splice(columnM_Pn6.length - 1, 1)


////..........Update value in cell and then write the data back to exel.............................
const Excel = require('exceljs');
const WB = new Excel.Workbook();

WB.xlsx.readFile('test.xlsx')
  .then(function () {

    const worksheet = WB.getWorksheet(1);
    let row1 = worksheet.getRow(1);
    let row2 = worksheet.getRow(2);
    let row3 = worksheet.getRow(3);
  
    //delete 4 column below
    let clcell
    for (let i = 1; i <= 30; i++) {
      clcell = worksheet.getRow(level.length)
      clcell.getCell(i).value = ""

      clcell = worksheet.getRow(level.length-1)
      clcell.getCell(i).value = ""

      clcell = worksheet.getRow(level.length-2)
      clcell.getCell(i).value = ""

      clcell = worksheet.getRow(level.length-3)
      clcell.getCell(i).value = ""
    }
    //.....Set title.......
    row1.getCell(1).value = "id"
    row1.getCell(2).value = "sn1"
    row1.getCell(3).value = "sn2"
    row1.getCell(4).value = "sn3"
    row1.getCell(5).value = "sn4"
    row1.getCell(6).value = "sn5"
    row1.getCell(7).value = "sn6"

    row1.getCell(8).value = "pn1"
    row1.getCell(9).value = "pn2"
    row1.getCell(10).value = "pn3"
    row1.getCell(11).value = "pn4"
    row1.getCell(12).value = "pn5"
    row1.getCell(13).value = "pn6"

    //.....Set colums values.......
    let rowindex
    for (let cell = 1; cell <= columnA_Id.length; cell++) {

      rowindex = worksheet.getRow(cell + 1)

      rowindex.getCell(1).value = columnA_Id[cell - 1].toString()
      rowindex.getCell(2).value = columnB_Sn1[cell - 1].toString()
      rowindex.getCell(3).value = columnC_Sn2[cell - 1].toString()
      rowindex.getCell(4).value = columnD_Sn3[cell - 1].toString()
      rowindex.getCell(5).value = columnE_Sn4[cell - 1].toString()
      rowindex.getCell(6).value = columnF_Sn5[cell - 1].toString()
      rowindex.getCell(7).value = columnG_Sn6[cell - 1].toString()

      rowindex.getCell(8).value = columnH_Pn1[cell - 1].toString()
      rowindex.getCell(9).value = columnI_Pn2[cell - 1].toString()
      rowindex.getCell(10).value = columnJ_Pn3[cell - 1].toString()
      rowindex.getCell(11).value = columnK_Pn4[cell - 1].toString()
      rowindex.getCell(12).value = columnL_Pn5[cell - 1].toString()
      rowindex.getCell(13).value = columnM_Pn6[cell - 1].toString()

      rowindex.getCell(14).value = ""
      rowindex.getCell(15).value = ""
      rowindex.getCell(16).value = ""
      rowindex.getCell(17).value = ""
    }

    
    row1.commit();
    row2.commit();
    row3.commit();
    rowindex.commit();
    clcell.commit();

    return WB.xlsx.writeFile('test3.xlsx');
  })



  