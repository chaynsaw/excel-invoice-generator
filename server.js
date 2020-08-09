XLSX = require('xlsx');
/* output format determined by filename */
const workbook = XLSX.utils.book_new();
const path = require('path')

const express = require('express')
const app = express()
const port = 3000

app.get('/', (req, res) => {
  res.send('Hello World!')
})
app.get('/write', (req, res) => {
	// XLSX.writeFile(workbook, 'outbound.xlsx');
	res.setHeader('Content-Type', 'application/vnd.ms-excel');
	// res.setHeader("Content-Disposition", "attachment; filename=" + "out.xlsx");
	res.sendFile('/Users/chaynorhsiao/code/excel-invoice-generator/outbound.xlsx')
	// res.end()
	// res.send('Writing file?')
})

app.get('/excel', function (req, res, next) {
	const filename = "outbound.xlsx"
  const options = {
    root: path.join(__dirname),
    headers: {
      'Content-Type': 'application/vnd.ms-excel',
      "Content-Disposition": "attachment; filename=" + filename
    }
  }

  res.sendFile(filename, options, function (err) {
    if (err) {
      next(err)
    } else {
			console.log('Sent:', filename)
    }
  })
})


app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`)
})

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
];
var ws = XLSX.utils.aoa_to_sheet(ws_data);

/* Add the worksheet to the workbook */
XLSX.utils.book_append_sheet(workbook, ws, ws_name);
ws['!ref'] = 'A1:G41'

let totalDaysWorked = 22
let dateInArrayForm = [2020, 8, 2]
let hoursPerDay = 8
let workDescription = 'Software Engineering'
let hourlyRate = 40

offset = 8
ws['A7'] = { t: 's', v: 'DATES'}
ws['B7'] = { t: 's', v: 'HOURS'}
ws['C7'] = { t: 's', v: 'WORK DESCRIPTION'}
ws['D7'] = { t: 's', v: 'HOURLY RATE'}
ws['E7'] = { t: 's', v: 'TOTAL'}

ws['A4'] = { t: 's', v: 'Start Date'}
ws['B4'] = { t: 'd', v: new Date([dateInArrayForm])}

let dayNum = 0
while (dayNum < totalDaysWorked) {
	let currentCellNum = offset

	let dateCell = `A${currentCellNum}`
	ws[dateCell] = { t: 'd', f: `WORKDAY(B4, ${dayNum})`}

	let hoursCell = `B${currentCellNum}`
	ws[hoursCell] = { t: 'n', v: hoursPerDay}

	let workDescCell = `C${currentCellNum}`
	ws[workDescCell] = { t: 's', v: workDescription }

	let hourlyRateCell = `D${currentCellNum}`
	ws[hourlyRateCell] = { t: 's', v: `$${hourlyRate}`}

	let totalCell = `E${currentCellNum}`
	ws[totalCell] = { t: 'n', v: hoursPerDay * hourlyRate}

	dayNum += 1
	offset += 1
}
ws[`E${offset}`] = { t: 'n', f: `SUM(E8:E${offset - 1})`}
ws[`D${offset}`] = { t: 's', v: 'Total:'}
