// const express = require('express')
// const mongoose = require('mongoose')
// const cors = require('cors');


// //using 
// var
//   datapumps = require('datapumps'),
//   Pump = datapumps.Pump,
//   RestMixin = datapumps.mixin.RestMixin,
//   MongodbMixin = datapumps.mixin.MongodbMixin,
//   ExcelWriterMixin = datapumps.mixin.ExcelWriterMixin,
//   ExcelReaderMinxin = datapumps.mixin.ExcelReaderMixin,
//   pump = new Pump();

//   pump
//     .mixin(MongodbMixin('mongodb://localhost:27017/maketing'))
//     .useCollection('contacts')
//     .mixin(ExcelReaderMinxin({
//       path: 'lottery.xlsx',
//       worksheet: 'lottery'
//     }))
//     .process(function(user) {
//       return pump.insert({name:user.Name,email:user.Email});
//     })
//   //   .mixin(ExcelWriterMixin())
//   // .createWorkbook('ContactsInUs1.xlsx')
//   // .createWorksheet('Contacts21')
//   // .writeHeaders(['Name', 'Email', 'Hua'])
 
//   // .process(function(contacts) {
//   //   return pump.writeRow([contacts.Name, contacts.Email, contacts.Hua]);
//   // })
//   .logErrorsToConsole()
//   .run()
//     .then(function() {
//       console.log("Done writing contacts to file");
//     });
  
// //..............Mongodb........................
//   ////...............MongoToExcel......................
// //  pump
// //   .mixin(MongodbMixin('mongodb://localhost:27017/maketing'))
// //   .useCollection('contacts')
// //   .from(pump.find({}))


//   //>>>>.....This write data from API.....
//   // .mixin(RestMixin)
//   // .fromRest({
//   //   query: function () {return pump.get('https://jsonplaceholder.typicode.com/posts');},
//   //   resultMapping: function (message) {
//   //     return message.result;
//   //   }
//   // })

//   //>>>.......this is write to excel....
//   // .mixin(ExcelWriterMixin())
//   // .createWorkbook('ContactsInUs1.xlsx')
//   // .createWorksheet('Contacts22')
//   // .writeHeaders(['Name', 'Email'])
 
//   // .process(function(contacts) {
//   //   return pump.writeRow([contacts.Name, contacts.Email]);
//   // })
//   // .logErrorsToConsole()
//   // .run()
//   //   .then(function() {
//   //     console.log("Done writing contacts to file");
//   //   });


//   //..........Data to mongodb..........................


// pump
//     .mixin(MongodbMixin('mongodb://localhost:27017/dataTomongo'))
//     .useCollection('user')
//     //.from(pump.find({}))


  
//  // >>>>..... data from API.....
//   .mixin(RestMixin)
//   .fromRest({
//     query: function () {return pump.get('https://jsonplaceholder.typicode.com/posts');},
//     resultMapping: function (message) {
//       return message.result;
//     }
//   })
// //.......data from excel file

// // .mixin(ExcelReaderMinxin())
// // .from()
  
//   .process(function(user) {
//     return pump.insert({id:user.id, title:user.title});
//   })
//   .logErrorsToConsole()
//   .run()
//     .then(function() {
//       console.log("Done writing contacts to file");
//     });

const express = require("express");
const app = express();
const initRoutes = require("./routes/lottery.route");
global.__basedir = __dirname + "/";

app.use(express.urlencoded({ extended: true }));
initRoutes(app);
//homepage rout
app.get("/", (req, res) => {
  return res.send('pppppp')
});

const date = new Date()

const mydate = date.getDate() +'-'+date.getMonth()+'-'+date.getFullYear();

console.log(mydate)



let port = 8080;
app.listen(port, () => {
  console.log(`Running at localhost:${port}`);
});
