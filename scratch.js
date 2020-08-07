XLSX = require('xlsx');
/* output format determined by filename */
const workbook = XLSX.utils.book_new();
var ws_name = "SheetJS";

/* make worksheet */
var ws_data = [
  [ "S", "h", "e", "e", "t", "J", "S" ],
  [  1 ,  2 ,  3 ,  4 ,  5 ]
];
var ws = XLSX.utils.aoa_to_sheet(ws_data);

/* Add the worksheet to the workbook */
XLSX.utils.book_append_sheet(workbook, ws, ws_name);
console.log(ws['A1'])
ws['F2'] = { v: 'S', t: 's' }
console.log(ws)
XLSX.writeFile(workbook, 'out.xlsb');
/* at this point, out.xlsb is a file that you can distribute */