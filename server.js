XLSX = require('xlsx');
/* output format determined by filename */

const path = require('path')

const express = require('express')
const app = express()
const port = 3000

var bodyParser = require('body-parser');
var multer = require('multer');
var upload = multer();

app.set('view engine', 'pug');
app.set('views', './views');

// for parsing application/json
app.use(bodyParser.json()); 

// for parsing application/xwww-
app.use(bodyParser.urlencoded({ extended: true })); 
//form-urlencoded

// for parsing multipart/form-data
app.use(upload.array()); 
app.use(express.static('public'));

app.post('/', function(req, res){
   console.log(req.body);
	 res.send("recieved your request!");
})

app.get('/', (req, res) => {
  res.send('Hello World!')
})


app.post('/promises', async function(req, res, next){
	const data = await createWorkbook(req.body);
	const sendExcelFile = await sendFile(req, res, next);
})

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`)
})

const sendFile = (req, res, next) => {
	const filename = "output.xlsx"
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
}

const createWorkbook = (params) => {
	console.log(params)
	var ws_name = "Invoice";
	const workbook = XLSX.utils.book_new();
/* make worksheet */

	const name = params.fullname
	const address = params.address
	const phone = params.phone

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

	let totalDaysWorked = params.days_worked
	let firstDate = new Date(params.start_date)
	let hoursPerDay = 8
	let workDescription = 'Software Engineering'
	let hourlyRate = params.hourly_rate

	offset = 8
	ws['A7'] = { t: 's', v: 'DATES'}
	ws['B7'] = { t: 's', v: 'HOURS'}
	ws['C7'] = { t: 's', v: 'WORK DESCRIPTION'}
	ws['D7'] = { t: 's', v: 'HOURLY RATE'}
	ws['E7'] = { t: 's', v: 'TOTAL'}

	ws['A4'] = { t: 's', v: 'Start Date'}
	ws['B4'] = { t: 'd', v: firstDate}
	console.log(ws['B4'])

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
	XLSX.writeFile(workbook, 'output.xlsx');
	
}
