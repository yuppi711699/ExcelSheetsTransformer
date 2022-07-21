let XLSX = require("xlsx");
let moment = require('moment');

let workbook = XLSX.readFile("export.xlsx");
let endLineNumber = workbook.Sheets.Export['!ref']
let border = endLineNumber.indexOf(':')+2;

if (isNaN(endLineNumber[border]) === false){
    endLineNumber = endLineNumber.slice(border, endLineNumber.length);
} else {
    border ++;
    endLineNumber = endLineNumber.slice(border, endLineNumber.length);
}

const cellsWorkDay =[
    'N',  'O', 'P', 'Q', 'R', 'S', 'T'
]
var day;
const content =[ ];
let obj = {}
function clearObjOfTask(){
    obj = {
        'Created By': 0,
        'Date and time of work': '01.01.1970',
        'Time (hh.mm)': 3,
        'Details': '',
        'Client or partner name': ''
    }
}
for(let i = 2; i < endLineNumber; i++) {
    clearObjOfTask();
    // console.log(workbook.Sheets.Export[`B${i}`].v)
    obj['Created By'] = workbook.Sheets.Export[`B${i}`].v;
    var day = moment(`${workbook.Sheets.Export[`F${i}`].v}`, "DD-MM-YYYY");
    console.log('day:', day)
    obj['Time (hh.mm)'] = workbook.Sheets.Export[`U${i}`].v;
    obj['Details'] = workbook.Sheets.Export[`L${i}`].v;
    obj['Client or partner name'] = workbook.Sheets.Export[`K${i}`].v;
    if (workbook.Sheets.Export[`M${i}`]) obj['Details'] += ". " + workbook.Sheets.Export[`M${i}`].v
    for (let j = 0; j < cellsWorkDay.length; j++) {
        if (workbook.Sheets.Export[`${cellsWorkDay[j]}${i}`].v) {
            let countWorkDayAtWeek = Number(workbook.Sheets.Export[`${cellsWorkDay[j]}${1}`].v.charAt(3));
            // console.log(countWorkDayAtWeek)
            day.add(countWorkDayAtWeek, 'days').calendar();
            obj['Date and time of work'] = day._i;
            // console.log(workbook.Sheets.Export[`${cellsWorkDay[j]}${1}`].v)
        } 
    }
    content.push(obj);
}

let a = workbook.Sheets.Export;
const worksheet = XLSX.utils.json_to_sheet(content);
const workbookk = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbookk, worksheet, `Deca4 timesheet `);
XLSX.utils.sheet_add_aoa(worksheet, [['Created By','Date and time of work','Time (hh.mm)','Details','Client or partner name']], { origin: "A1" });

worksheet["!cols"] = [ { wch: 20} ];
XLSX.writeFile(workbookk, "timesheet.xlsx");
