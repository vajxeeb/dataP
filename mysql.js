
const express = require('express')
const cors = require('cors');
const mysql = require('mysql');
const mixin = require('datapumps/lib/mixin');

let dbCon = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: '12345',
    database: 'dbaccout',
});


//Check connection.....
dbCon.connect(function (err) {
    if (err) {
        console.log("connect faild");
    }
    if (!err) {
        console.log("connected");
    }
});


//....using datapumps ..........
var
    datapumps = require('datapumps'),
    Pump = datapumps.Pump,
    ExcelReaderMinxin = datapumps.mixin.ExcelReaderMixin,
    pump = new Pump();


//......write  form mysql to excel............
// pump.from(dbCon.query('SELECT * FROM user'));
//   pump
//   .mixin(datapumps.mixin.CsvWriterMixin({
//     path: 'aaaa.csv',
//     headers: [ 'Id', 'name', 'email' ]
//   }))
//pump.from(dbCon.query('insert into user(name,email) values (?,?)'));
//  pump.process(function(user) {
//     this.writeRow([ user.id, user.name, user.email ]);
//   });
// pump
//     .logErrorsToConsole()
//     .run()
//     .then(function () {
//         console.log("Done writing contacts to file");
//     });



//.................reader data to mysql .........................


pump
    .mixin(ExcelReaderMinxin({
        path: 'source/lottery.xlsx',
        worksheet: 'lottery'
    }))
    .process(function (lottery){ 
       //dbCon.query(`insert into user(name,email) values ('${user.Name}','${user.Email}')`);
       //dbCon.query('insert into tblottery_sell values (?,?,?,?,?,?,?)',[lottery.id,lottery.n1,lottery.n2,lottery.n3,lottery.n4,lottery.n5,lottery.n6])
    });
pump
    .logErrorsToConsole()
    .run()
    .then(function () {
        console.log("Done writing contacts to file");
    });

