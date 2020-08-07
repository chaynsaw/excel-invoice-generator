XLSX = require('xlsx');
/* output format determined by filename */
const workbook = XLSX.utils.book_new();
var ws_name = "SheetJS";

/* make worksheet */

const name = 'Chaynor Hsiao'
const address = '95 Rae Ave, San Francisco, CA'
const phone = 5105858082

const invoiceNumber = ''
const invoiceDate = '8/5/20'
const taxID = ''

var ws_data = [
  ["Name", name, "", "", "", "Invoice" ],
	["Address" , address, '', '', '', "Invoice No.:", invoiceNumber],
	["Tel:", phone, '', '', '', "Date:", invoiceDate],
	['', '', '', '', '', "Tax ID:", taxID],
	[],
	['DATES']
];
var ws = XLSX.utils.aoa_to_sheet(ws_data);

/* Add the worksheet to the workbook */
XLSX.utils.book_append_sheet(workbook, ws, ws_name);
ws['!ref'] = 'A1:G41'

let totalDaysWorked = 30
offset = 6
let date = new Date(2020, 8, 1)
for (i = 1; i <= totalDaysWorked; i++) {
	cell = `A${offset + i}`
	date = new Date(date.getTime() + 86400000)
	ws[cell] = { v: date, t: 'd'}
}

XLSX.writeFile(workbook, 'out.xlsb');
/* at this point, out.xlsb is a file that you can distribute */