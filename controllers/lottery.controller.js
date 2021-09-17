
const express = require('express')
const cors = require('cors');
const mysql = require('mysql');
const mixin = require('datapumps/lib/mixin');
const excel = require("exceljs");
const xlsx = require('xlsx')


//........create connetion.............
let dbCon = mysql.createConnection({
  host: 'localhost',
  user: 'root',
  password: '12345',
  database: 'DB_Excel_Finace',
});
//......using datapumps..............
var
  datapumps = require('datapumps'),
  Pump = datapumps.Pump,
  ExcelReaderMinxin = datapumps.mixin.ExcelReaderMixin,
  pump = new Pump();

//........Check connection.....
dbCon.connect(function (err) {
  if (err) {
    console.log("connect faild");
  }
  if (!err) {
    console.log("connected");
  }
});
//.................load file excel .........................
const upload = async (req, res) => {
  try {
    if (req.file == undefined) {
      return res.status(400).send("Please upload an excel file!");
    }
    //....path of excel file.............
    let path =
      __basedir + "/" + req.file.filename;
    //resources/static/assets/uploads/
    let wb = xlsx.readFile(req.file.filename);
    let wsh = wb.SheetNames[0];
    let installment = '220'
    console.log(wsh)
    //..............read excel file to db
    pump
      .mixin(ExcelReaderMinxin({
        path: path,
        worksheet: 'Sheet1'
      }))
      //........query(INSER)........
      .process(async function (lottery) {
        //dbCon.query(`insert into user(name,email) values ('${user.Name}','${user.Email}')`);
        await dbCon.query('insert into TBTotal_Sell (lottery_id,total_sell_num1,total_sell_num2,total_sell_num3,total_sell_num4,total_sell_num5,total_sell_num6,installment)  VALUES (?,?,?,?,?,?,?,?)', [lottery.id, lottery.sn1, lottery.sn2, lottery.sn3, lottery.sn4, lottery.sn5, lottery.sn6, installment]
          , (error, result) => {
            if (error) throw error;
            else {
             // console.log('Completed')
            }
          })
      });
    pump
      .logErrorsToConsole()
      .run()
      .then(function () {
        console.log("Done writing contacts to file");
      });


  } catch (error) {
    console.log(error);
    res.status(500).send({
      message: "Could not upload the file: " + req.file.originalname,

    });
  }

};
module.exports = {
  upload
};





