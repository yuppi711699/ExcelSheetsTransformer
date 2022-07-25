// /*Do do
// 1. fix data calculating

// */
// let XLSX = require("xlsx");
// let moment = require('moment');
// const colors = require("colors");

// let fileInputName ="export.xlsx";
// let fileOutputName ="timesheet.xlsx";
// let workbook = XLSX.readFile(fileInputName);
// let endLineNumber = workbook.Sheets.Export['!ref'];
// let border = endLineNumber.indexOf(':')+2;

// if (isNaN(endLineNumber[border]) === false){
//     endLineNumber = endLineNumber.slice(border, endLineNumber.length);
// } else {
//     border ++;
//     endLineNumber = endLineNumber.slice(border, endLineNumber.length);
// }

// const cellsWorkDay =['N',  'O', 'P', 'Q', 'R', 'S', 'T'];
// var day;
// const content =[ ];
// let obj = {};
// function clearObjOfTask(){
//     obj = {
//         'Created By': 0,
//         'Date and time of work': '01.01.1970',
//         'Time (hh.mm)': 3,
//         'Details': '',
//         'Client or partner name': ''
//     }
// }
// for(let i = 2; i < 8; i++) {//endLineNumber; i++) {
//     clearObjOfTask();
//     // console.log(workbook.Sheets.Export[`B${i}`].v)
//     obj['Created By'] = workbook.Sheets.Export[`B${i}`].v;
//     var day = moment(`${workbook.Sheets.Export[`F${i}`].v}`, "DD-MM-YYYY");
//     // console.log('day:', day)
//     obj['Time (hh.mm)'] = workbook.Sheets.Export[`U${i}`].v;
//     obj['Details'] = workbook.Sheets.Export[`L${i}`].v;
//     obj['Client or partner name'] = workbook.Sheets.Export[`K${i}`].v;
//     if (workbook.Sheets.Export[`M${i}`]) obj['Details'] += ". " + workbook.Sheets.Export[`M${i}`].v;
//     for (let j = 0; j < cellsWorkDay.length; j++) {
//         if (workbook.Sheets.Export[`${cellsWorkDay[j]}${i}`].v) {
//             let countWorkDayAtWeek = Number(workbook.Sheets.Export[`${cellsWorkDay[j]}${1}`].v.charAt(3));
//             // console.log('countWorkDayAtWeek: ', (countWorkDayAtWeek))
//             // console.log(day._i)
//             // console.log(countWorkDayAtWeek)
//             // console.log('50: ', day._i)
//             day.add(countWorkDayAtWeek, 'day').format("DD-MM-YYYY");
//             console.dir('51: ', day._i);
//             console.log('52: ', day._d)//
//             let workDate = (JSON.stringify(day._d)).slice(1, 11).replace( /-/g, "." )
//             console.log('workDate: ', workDate.length)//

//              workDate = String(workDate[8]) + String(workDate[9]) + '.' + workDate[5] + workDate[6] + '.' + workDate[0] + workDate[1] + workDate[2] + workDate[3];
//             // console.log('56: ', date)//

//             obj['Date and time of work'] = workDate//t//day._d;
//             console.log('59: ', obj['Date and time of work'])

//             // console.log(workbook.Sheets.Export[`${cellsWorkDay[j]}${1}`].v)
//         } 
//     }
//     content.push(obj);
// }

// const worksheet = XLSX.utils.json_to_sheet(content);
// const workbookk = XLSX.utils.book_new();
// XLSX.utils.book_append_sheet(workbookk, worksheet, `Deca4 timesheet `);
// XLSX.utils.sheet_add_aoa(worksheet, [['Created By','Date and time of work','Time (hh.mm)','Details','Client or partner name']], { origin: "A1" });

// worksheet["!cols"] = [ { wch: 20} ];
// XLSX.writeFile(workbookk, "timesheet.xlsx");
// console.log(`Data in file`, `${fileInputName}`.green, `transformed and written in`, `${fileOutputName}`.green);
