let XLSX = require("xlsx");
let moment = require('moment');
let fs = require('fs');


// import { writeFile, set_fs } from "xlsx/xlsx.mjs";
let workbook = XLSX.readFile("export.xlsx");
let endLineNumber = workbook.Sheets.Export['!ref']
let border = endLineNumber.indexOf(':')+2;
// console.log(endLineNumber, ': ', isNaN(endLineNumber[border]), ': ', (endLineNumber[border]))
if (isNaN(endLineNumber[border]) === false){
    endLineNumber = endLineNumber.slice(border, endLineNumber.length);
} else {
    border ++;
    endLineNumber = endLineNumber.slice(border, endLineNumber.length);
}
const dayConst = {
    N: 1,
    O: 2,
    P: 3,
    Q: 4,
    R: 5,
    S: 6,
    T: 7
}

const cells =[
    'B', 'F', 'I', 'K', 'L', 'M'
]
const cellsWorkDay =[
    'N',  'O', 'P', 'Q', 'R', 'S', 'T'
]
var day;
let test = {
    'B': 1,
    'D': 2,
    'F': 3,
    'G': 4, 'H': 5, 'K': 6 , 'L': 7
}
const content =[ ];
let outputData = [];
let counterWorkHours
let obj = {
    'Created By': 0,
    'Date and time of work': '01.01.1970',
    'Time (hh.mm)': 3,
    'Details': '',
    'Client or partner name': ''
}
for(let i = 2; i < endLineNumber; i++) {
    // console.log(workbook.Sheets.Export[`B${i}`].v)
    obj['Created By'] = workbook.Sheets.Export[`B${i}`].v;
    // obj['Date and time of work'] = workbook.Sheets.Export[`F${i}`].v;
    var day = moment(`${workbook.Sheets.Export[`F${i}`].v}`, "DD-MM-YYYY");
    console.log('day:', day)
    obj['Time (hh.mm)'] = workbook.Sheets.Export[`U${i}`].v;
    obj['Details'] = workbook.Sheets.Export[`L${i}`].v;
    obj['Client or partner name'] = workbook.Sheets.Export[`K${i}`].v;
    if (workbook.Sheets.Export[`M${i}`]) obj['Details'] += ". " + workbook.Sheets.Export[`M${i}`].v
    // counterWorkHours = 0;
    for (let j = 0; j < cellsWorkDay.length; j++) {
        // const element = array[i];
        if (workbook.Sheets.Export[`${cellsWorkDay[j]}${i}`].v) {
            let countWorkDayAtWeek = Number(workbook.Sheets.Export[`${cellsWorkDay[j]}${1}`].v.charAt(3));
            // console.log(countWorkDayAtWeek)
            day.add(countWorkDayAtWeek, 'days').calendar();
            // console.dir(day);
            // console.dir(day._i);
            // day = day.slice(6, 10);
            // console.log(day);
            obj['Date and time of work'] = day._i;
            // console.log(workbook.Sheets.Export[`${cellsWorkDay[j]}${1}`].v)
        } 
    }

    content.push(obj);
    // counter = 0;
    obj ={
        'Created By': 0,
        'Date and time of work': '01.01.1970',
        'Time (hh.mm)': 3,
        'Details': '',
        'Client or partner name': ''
    }
    // for(let j=0; j<cells.length; j++) {

    //     const cell = cells[j];
        // console.log(`${cells[j]}${i}`,workbook.Sheets.Export[`${cells[j]}${i}`])
    // }
}
// for (const item of cells) {
//     // console.log(item,test[item])
//     let cellAdress = workbook.Sheets.Export[`${item}${test[item]}`]
//     // if(cellAdress) console.dir(cellAdress.v)
// }
console.log('/////////////');
// console.dir(workbook.Sheets.Export)
// let outputData = [];
// let obj = {
//     'Created By': 0,
//     'Date and time of work': '01.01.1970',
//     'Time (hh.mm)': 3,
//     'Details': '',
//     'Client or partner name': ''
// }

// let counter = 0;
// for(let i in workbook.Sheets.Export){
//     if(counter == 21){
//         content.push(obj);
//         counter = 0;
//         obj ={
//             'Created By': 0,
//             'Date and time of work': '01.01.1970',
//             'Time (hh.mm)': 3,
//             'Details': '',
//             'Client or partner name': ''
//         }
//     }  
//     // console.log(counter + workbook.Sheets.Export[i].v)
//     if(counter === 2){
//         obj['Created By'] = workbook.Sheets.Export[i].v
//     } else if (counter === 6){
//         obj['Date and time of work'] = workbook.Sheets.Export[i].v
//     } else if (counter === 11){
//         obj['Client or partner name'] = workbook.Sheets.Export[i].v
//     } else if (counter === 12){
//         obj['Details'] += workbook.Sheets.Export[i].v
//     } else if (counter === 13){
//         obj['Details'] += workbook.Sheets.Export[i].v
//     }
//     // console.log('/////////////');
//     // let tt = workbook.Sheets.Export[i].v
//     // if (workbook.Sheets.Export[i].v !==null || workbook.Sheets.Export[i].v !== undefined || workbook.Sheets.Export[i].v.length !== 0){
//     //     // console.log('/////+++++++++++++++////////', tt, '  ::  ', typeof(tt),  tt.length);
//     //     // console.log(workbook.Sheets.Export[i].v, " : ", (workbook.Sheets.Export[i]));
//     // }
//     // counter ++;
// }
// let outputData;
// console.dir(workbook)
// console.dir(workbook.Sheets.Export['!margins'])
// console.dir(workbook.Sheets.Export['!ref'])
// let endLineNumber = workbook.Sheets.Export['!ref']
// let border = endLineNumber.indexOf(':')+2;
// console.log(endLineNumber, ': ', isNaN(endLineNumber[border]), ': ', (endLineNumber[border]))
// if (isNaN(endLineNumber[border]) === false){
//     endLineNumber = endLineNumber.slice(border, endLineNumber.length)
// }
// console.log('endLineNumber:', endLineNumber)
// console.dir(border)
// for(let i in workbook.Sheets.Export){
    // console.log('/////////////');
    // console.dir(i)//, ':  ', workbook.Sheets.Export.A2[i]);
    // outputFile = workbook.Sheets.Export.A2.v
// }
// console.dir(workbook.Sheets.Export.A2)

// console.dir(outputData)
//outputFile //= workbook.Sheets.Export.B2
// console.dir(outputFile)

// const content =
//  [
//     { name: 'George Washington', birthday: '1732-02-22' },
//     { name: 'John Adams', birthday: '1735-10-19' }
// ];
let a = workbook.Sheets.Export;
// console.dir(a)
const workbook2 = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(content);
const workbookk = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbookk, worksheet, "Deca4 timesheet");
XLSX.utils.sheet_add_aoa(worksheet, [['Created By','Date and time of work','Time (hh.mm)','Details','Client or partner name']], { origin: "A1" });

worksheet["!cols"] = [ { wch: 20} ];
XLSX.writeFile(workbookk, "timesheet.xlsx");
// async() => {
//     /* fetch JSON data and parse */
//     const url = "https://sheetjs.com/executive.json";
//     const raw_data = await (await fetch(url)).json();
  
//     /* filter for the Presidents */
//     const prez = raw_data.filter(row => row.terms.some(term => term.type === "prez"));
  
//     /* flatten objects */
//     const rows = prez.map(row => ({
//       name: row.name.first + " " + row.name.last,
//       birthday: row.bio.birthday
//     }));
// //   console.dir(rows);
//     /* generate worksheet and workbook */
//     const worksheet = XLSX.utils.json_to_sheet(rows);
//     const workbook = XLSX.utils.book_new();
//     XLSX.utils.book_append_sheet(workbook, worksheet, "Dates");
  
//     /* fix headers */
//     XLSX.utils.sheet_add_aoa(worksheet, [["Name", "Birthday"]], { origin: "A1" });
  
//     /* calculate column width */
//     // const max_width = rows.reduce((w, r) => Math.max(w, r.name.length), 10);
//     // worksheet["!cols"] = [ { wch: max_width } ];
  
//     /* create an XLSX file and try to save to Presidents.xlsx */
//     // XLSX.writeFile(workbook, "Presidents.xlsx");
//   }
// try {
//     XLSX.writeFile('/Users/dmy/Downloads/programming/deca4/dataTimeSheetTransformer/qqqq.xlsx', workbook2, { type: "file", bookType: "xlsx" });
//   // file written successfully
// } catch (err) {
//   console.error(err);
// }
// outputFile = 
// XLSX.writeFile('qqqq.xlsx', outputFile, { type: "array", bookType: "xlsx" });
// fs.writeFile("/Users/dmy/Downloads/programming/deca4/dataTimeSheetTransformer/out.txt", 'dsfv');
// fs.writeFile("/out.xlsx", outputFile);