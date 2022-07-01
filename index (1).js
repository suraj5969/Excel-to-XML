/*changes done 
1. notify user that error are logged in log folder
2. remove space from XML file name
3. validate date
4. Match row count of excel file and xml file (demo.js)
5. 2 input and 2 output
*/

// const cron = require('node-cron');
const fs = require('fs')
const jsonxml = require('jsontoxml');
const vkbeautify = require('vkbeautify');
const XLSX = require('xlsx');
const { getJsDateFromExcel } = require("excel-date-to-js");
const { format, isValid, parse } = require('date-fns');
const path = require('path');
const express = require('express');
const chokidar = require('chokidar');
const sql = require('mssql');

const app = express()
const port = 6000;
app.listen(port);

let sqlConfig = {
	user: 'pratik',
	password: 'Pratik@123',
	server: 'DESKTOP-K1028K1',
	database: 'files',
	connectionTimeout: 30000,
	requestTimeout: 30000,
	dialect: 'mssql',
	options: {
		encrypt: true, // for azure
		trustServerCertificate: true, // change to true for local dev / self-signed certs
	}
};

//for ITK files
const inputDirectoryITK = './InputITK';
const outputDirectoryITK = './OutputITK';
const logsDirectoryITK = './LogsITK';

//watcher for ITK
const watcherITK = chokidar.watch(inputDirectoryITK, {
	persistent: false,
	awaitWriteFinish: {
		stabilityThreshold: 5000,
		pollInterval: 500
	},
})


//for D&D files
const inputDirectoryDnD = './InputDnD';
const outputDirectoryDnD = './OutputDnD';
const logsDirectoryDnD = './LogsDnD';

//watcher for D&D
const watcherDnD = chokidar.watch(inputDirectoryDnD, {
	persistent: false,
	awaitWriteFinish: {
		stabilityThreshold: 5000,
		pollInterval: 500
	},
})

const ext = ['.xlsx', '.xlsm', '.xls', '.csv'];
const headers = ["USER_NAME", "CONTACT", "CUSTOMER_REFERENCE", "ACCOUNT_EMAIL", "PRODUCT_ID", "PRODUCT_GROUP", "PRODUCT_DESCRIPTION", "TRANSACTIONID", "DATE", "REFERENCE", "MATTER", "REQUEST_ID", "AMT_BEFORE_TAX", "AMT_AFTER_TAX", "GST_AMT", "BILLING_FREQUENCY", "RETAILER_REFERENCE", "PERIOD_END_DT"];


try {
	const directoryPathITK = path.join(__dirname, inputDirectoryITK);
	// make files for ITK
	if (!fs.existsSync(outputDirectoryITK)) {
		fs.mkdirSync(outputDirectoryITK);
	}
	if (!fs.existsSync(logsDirectoryITK)) {
		fs.mkdirSync(logsDirectoryITK);
	}
	if (!fs.existsSync(`${logsDirectoryITK}/files`)) {
		fs.mkdirSync(`${logsDirectoryITK}/files`);
	}
	if (!fs.existsSync(`${logsDirectoryITK}/files/Success`)) {
		fs.mkdirSync(`${logsDirectoryITK}/files/Success`);
	}
	if (!fs.existsSync(`${logsDirectoryITK}/files/Fail`)) {
		fs.mkdirSync(`${logsDirectoryITK}/files/Fail`);
	}
	if (!fs.existsSync(`${logsDirectoryITK}/files/XML`)) {
		fs.mkdirSync(`${logsDirectoryITK}/files/XML`);
	}

	const directoryPathDnD = path.join(__dirname, inputDirectoryDnD);
	// make files for D&D
	if (!fs.existsSync(outputDirectoryDnD)) {
		fs.mkdirSync(outputDirectoryDnD);
	}
	if (!fs.existsSync(logsDirectoryDnD)) {
		fs.mkdirSync(logsDirectoryDnD);
	}
	if (!fs.existsSync(`${logsDirectoryDnD}/files`)) {
		fs.mkdirSync(`${logsDirectoryDnD}/files`);
	}
	if (!fs.existsSync(`${logsDirectoryDnD}/files/Success`)) {
		fs.mkdirSync(`${logsDirectoryDnD}/files/Success`);
	}
	if (!fs.existsSync(`${logsDirectoryDnD}/files/Fail`)) {
		fs.mkdirSync(`${logsDirectoryDnD}/files/Fail`);
	}
	if (!fs.existsSync(`${logsDirectoryDnD}/files/XML`)) {
		fs.mkdirSync(`${logsDirectoryDnD}/files/XML`);
	}


	function removeSpecialChars(str) {
		// always trim the string its important that values dont have any leading or trailing spaces
		let val = String(str).replace(/&/g, '&amp;');
		val = val.replace(/>/g, '&gt;');
		val = val.replace(/</g, '&lt;').trim();

		//to remove invisible(zero width) characters
		const zeroWidthSpace = '\u200B';
		const zeroWidthNoBreakSpace = '\uFEFF';
		const zeroWidthSpaceRegEx = new RegExp(`${zeroWidthSpace}`, 'g');
		const zeroWidthNoBreakSpaceRegEx = new RegExp(`${zeroWidthNoBreakSpace}`, 'g');
		val = val.replace(zeroWidthSpaceRegEx, '');
		val = val.replace(zeroWidthNoBreakSpaceRegEx, '');

		//ask which special characters to replace as it will work with keeping all special characters
		// val = String(val).replace(/[^\w\s]/g, '').trim()
		// we don't need to replace single and double quotes as we are not giving attributes to elements
		// atributes are strings and are represented in single or double quotes so we need to escape ' or ", depending which one was used as the attribute delimiter.
		// val = val.replace(/'/g, '&apos;');
		// val = val.replace(/"/g, '&quot;');
		return val;
	}

	function formatDate(date, isTime, toUTC) {
		let d = new Date(date);
		if (isValid(d)) {
			if (toUTC) {
				//getJsDateFromExcel return Date Object and it is in our local timezone
				//converting the timezone to UTC beacsue I think excel stores date in UTC
				d = new Date(d.valueOf() + d.getTimezoneOffset() * 60 * 1000);
			}

			if (isTime) {
				//to round timeto nearest second as processing is producing mismatch in miliseconds
				d.setSeconds(d.getSeconds() + Math.round(d.getMilliseconds() / 1000))
				d.setMilliseconds(0);
				d = format(d, "yyyy-MM-dd'T'HH:mm:ss.SSS");
			}
			else {
				d = format(d, "yyyy-MM-dd")
			}
			return d;
		}
		else {
			return 'Invalid Date';
		}
	}

	function parseExcelDate(date, isTime) {
		try {
			let d = getJsDateFromExcel(date)
			d = formatDate(d, isTime, true);
			return d;
		}
		catch (e) {
			// if getJsDateFromExcel throws a error when this field is a string but valid date
			const d = formatDate(date, isTime, false);
			if (d == 'Invalid Date') {
				let result = 'Invalid Date';
				if (date.includes('-')) {
					const [day, month, year] = date.split('-');
					result = [year, month, day].join('-');
				}
				if (date.includes('/')) {
					const [day, month, year] = date.split('/');
					result = [year, month, day].join('-');
				}
				return formatDate(result, false, true);
				// return result;
			}
			else {
				return d;
			}
		}
	}

	function ContinueToNextFile(fileName) {
		//check if file is an excel file and not error.txt and not a hidden file if the same file is opened in excel app
		if (fileName.startsWith('~$') || !(ext.includes(path.extname(fileName))))
			return true;
		else return false;
	}

	function readExcelFile(fileName, fileDirectory, logDirectory) {
		const workbook = XLSX.readFile(`${fileDirectory}/${fileName}`);
		const sheet_name_list = workbook.SheetNames;

		if (sheet_name_list.length > 1) {
			fs.writeFile(`${logDirectory}/warning.txt`,
				`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - ${fileName} | File has more than one sheets. Continuing to process only first sheet in file.`,
				{ flag: 'a' },
				(error) => { })
		}
		const result = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]], { defval: "" });
		return result;
	}

	function insertDataInDB(fileName, success, data) {
		sql.connect(sqlConfig).then(pool => {
			return pool.request()
				.query(`INSERT INTO files VALUES ('${fileName.replace(/'/g, "''")}','${format(new Date(), 'yyyy-MM-dd-hh:mm:ss')}',${success ? '1' : '0'},'${data.replace(/'/g, "''")}')`)
		}).then(result => {
			console.log(result)
		}).catch(err => {
			// ... error checks
			console.log(err);
		});
	}

	function moveLogFile(file, outputDirectory, logsDirectory) {
		const { birthtime } = fs.statSync(file)
		const nowDate = new Date();
		// console.log(file, nowDate);
		// console.log(file, birthtime);
		let Difference_In_Time = nowDate.getTime() - birthtime.getTime();
		let Difference_In_Days = Difference_In_Time / (1000 * 3600 * 24);
		// console.log(Math.trunc(Difference_In_Days));
		// console.log("Total number of days between dates  :" + (Difference_In_Days));
		const nameInfo = path.parse(fileName);
		if (Math.trunc(Difference_In_Days) == 0) {
			fs.renameSync(`${outputDirectory}/${file}`, `${logsDirectory}/files/${nameInfo.name}_${count}.${nameInfo.ext}`)
		}
	}

	function isColCountCorrect(fileRowJson, fileName, outputDirectory, inputDirectory, logsDirectory, count) {
		// we considering fixed structure of excel file
		// means columns ordering will be same for all input excel files
		const len = Object.keys(fileRowJson).length;
		if (len === 18)
			return true;
		else {
			console.log(`Files moved to ${logsDirectory}/files/Fail`);
			const nameInfo = path.parse(fileName);
			fs.renameSync(`${inputDirectory}/${fileName}`, `${logsDirectory}/files/Fail/${nameInfo.name}_${count}.${nameInfo.ext}`)
			// fs.unlinkSync(`${inputDirectory}/${fileName}`);

			fs.writeFileSync(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`,
				`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - ${fileName} | The number of columns in the Excel file is not correct. There should be exactly 18 columns in file.`,
				{ flag: 'a' },
				(error) => { })

			fs.readFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`, 'utf8', function (err, data) {
				//upload data to database
				// insertDataInDB(fileName, true, data)
			});
			return false;
		}
	}


	function readDirectory(filePath, inputDirectory, logsDirectory, outputDirectory, count) {
		// console.log(onPath, 'File added');
		console.log("Process started");
		try {
			const fileName = path.basename(filePath);
			let parseXML = true;
			if (ContinueToNextFile(fileName))
				return;
			const result = readExcelFile(fileName, inputDirectory, outputDirectory);

			if (!isColCountCorrect(result[0], fileName, outputDirectory, inputDirectory, logsDirectory, count)) {
				// parseXML = false;
				return;
			}

			const emailRowToDel = [];
			const emptyCustRef = [];
			const billFreq = [];
			const date = [];
			const billFreqRegex = /weekly|monthly/i;
			const emailRegex = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
			const email = [];
			const transIDDuplicate = [];
			const transIDSet = new Set();
			result.forEach(function (row, rowno) {
				Object.keys(row).forEach(function (key, i) {
					if (i === 2 && row[key] === '') {
						emptyCustRef.push(row);
					}
					// if (i === 3) {
					// 	if (emailRegex.test(row[key])) {

					// 	}
					// 	else {
					// 		email.push(row)
					// 	}
					// }
					if (i === 3 && row[key].includes('@lexisnexis.com')) {
						console.log(row[key]);
						emailRowToDel.push(row);

					}
					if (i === 12) {
						row[key] = Number(row[key]);
					}
					if (i === 13 && row[key] === '') {
						row[key] = Number(row[key]);
					}
					if (i === 14 && row[key] === '') {
						row[key] = Number(row[key]);
					}
					if (i === 15) {
						//because they want to only include Weekly in output not monthly
						if (billFreqRegex.test(row[key])) {
							row[key] = 'Weekly';
						}
						else {
							billFreq.push(row);
						}
					}
					if (i === 7) {
						if (transIDSet.has(row[key])) {
							transIDDuplicate.push(row);
						} else {
							transIDSet.add(row[key]);
						}
					}
					if (i === 8) {
						row[key] = parseExcelDate(row[key], true);
						if (row[key].includes('Invalid Date')) {
							date.push(row);
						}
					}
					if (i === 17) {
						row[key] = parseExcelDate(row[key], false);
					}

					let val = row[key];
					if (typeof val === 'number') {
						//round the number to 2 decimal places as sometimes excel may give a number with more than 2 decimal places
						// for example: excel may show a number 1.80 in excel app but here it sometimes gives 1.7999999999999998
						val = Math.round((val + Number.EPSILON) * 100) / 100;
					}
					else {
						val = removeSpecialChars(val);
					}

					// Remove key-value from object
					delete row[key];
					// Add value with new key
					// insted of creating a new object we are modifying the original object got from excel file
					row[headers[i]] = val;
				});
				result[rowno] = { "LN_ITK_GBX_TBL": row }
			});

			if (emptyCustRef.length > 0) {
				parseXML = false;
				emptyCustRef.forEach(function (row) {
					const nameInfo = path.parse(fileName);
					// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
					fs.writeFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`,
						`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | File has empty customer reference for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}. `,
						{ flag: 'a' },
						(error) => { })
				});
			}

			if (date.length > 0) {
				parseXML = false;
				const nameInfo = path.parse(fileName);
				fs.writeFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`,
					`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | File has invalid date for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}. `,
					{ flag: 'a' },
					(error) => { })
			}

			if (emailRowToDel.length > 0) {
				parseXML = false;
				emailRowToDel.forEach(function (row) {
					const nameInfo = path.parse(fileName);
					// writing email row as (index+2) because array index starts from 0 and excel file has first row of headers
					fs.writeFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`,
						`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | File has a email = ${row['ACCOUNT_EMAIL']} that includes lexisnexis.com domain for username = ${row['USER_NAME']}. .`,
						{ flag: 'a' },
						(error) => { });
				});
			}
			// if (email.length > 0) {
			// 	parseXML = false;
			// 	email.forEach(function (row) {
			// 		// writing email row as (index+2) because array index starts from 0 and excel file has first row of headers
			// 		fs.writeFile(`${outputDirectory}/${count}_${fileName}_ERROR.txt`,
			// 			`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | File has a email = ${row['ACCOUNT_EMAIL']} that is not in valid email format for username = ${row['USER_NAME']}. .`,
			// 			{ flag: 'a' },
			// 			(error) => { });
			// 	});
			// }
			if (billFreq.length > 0) {
				parseXML = false;
				billFreq.forEach(function (row) {
					const nameInfo = path.parse(fileName);
					// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
					fs.writeFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`,
						`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | File has Billing_Frequency = ${row['BILLING_FREQUENCY']} which is not weekly or monthly for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`,
						{ flag: 'a' },
						(error) => { })
				});
			}

			if (transIDDuplicate.length > 0) {
				parseXML = false;
				transIDDuplicate.forEach(function (row) {
					const nameInfo = path.parse(fileName);
					// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
					fs.writeFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`,
						`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | File has Duplicate Transaction ID = ${row['TRANSACTIONID']} for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`,
						{ flag: 'a' },
						(error) => { })
				});
			}
			if (parseXML) {
				// count++;
				let xml = jsonxml({ 'ORDER_DATA': result }, { xmlHeader: { standalone: true } })
				xml = vkbeautify.xml(xml, 4);
				fs.writeFileSync(`${outputDirectory}/${path.parse(fileName).name.replace(/\s/g, '')}.xml`, xml, { flag: 'w' }) //remove spaces from output file name
				console.log(`Output file created`);
				//delete a file from a folder
				//copy XML file in logs folder
				// const delFile = `${inputDirectory}/${fileName}`;
				const nameInfo = path.parse(fileName);
				fs.renameSync(`${inputDirectory}/${fileName}`, `${logsDirectory}/files/Success/${nameInfo.name}_${count}.${nameInfo.ext}`)
				fs.copyFileSync(`${outputDirectory}/${path.parse(fileName).name.replace(/\s/g, '')}.xml`, `${logsDirectory}/files/XML/${path.parse(fileName).name.replace(/\s/g, '')}_${count}.xml`);
				// fs.unlinkSync(delFile);

				// console.log(`file_name : ${fileName}`);
				// console.log(`date_time : ${format(new Date(), 'yyyy-MM-dd-hh:mm:ss')}`);

				//upload data to database
				// insertDataInDB(fileName, false, '')

			}
			else {
				// count++;
				console.log(`Files moved to ${logsDirectory}/files/Fail`);
				const nameInfo = path.parse(fileName);
				// console.log(`${outputDirectoryITK}/${format(new Date(), 'yyyy-MM-dd')}_${fileName}_ERROR.txt`);
				fs.renameSync(`${inputDirectory}/${fileName}`, `${logsDirectory}/files/Fail/${nameInfo.name}_${count}.${nameInfo.ext}`)
				// fs.unlinkSync(`${inputDirectory}/${fileName}`);
				fs.readFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`, 'utf8', function (err, data) {
					//upload data to database
					// insertDataInDB(fileName, true, data)
				});
				// fs.readdirSync(outputDirectory).forEach(fileName => {
				// 	console.log(fileName);
				// 	if (fileName.includes('_ERROR.txt')) {
				// 		moveLogFile(`${fileName}`, outputDirectory, logsDirectory);
				// 	}

				// });
			}
		}
		catch (err) {
			// console.log(err);
			const nameInfo = path.parse(fileName);
			fs.writeFileSync(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`,
				`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - ${fileName} | ${err}`,
				{ flag: 'a' },
				(error) => { })
			fs.readFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`, 'utf8', function (err, data) {
				//upload data to database
				// insertDataInDB(fileName, true, data)
			});
			// fs.readdirSync(outputDirectory).forEach(fileName => {
			// 	console.log(fileName);
			// 	if (fileName.includes('_ERROR.txt')) {
			// 		moveLogFile(`${fileName}`, outputDirectory, logsDirectory);
			// 	}

			// });
		}



	}

	// start watcherITK when there is file in inputITK folder
	let FileCount = 0;
	watcherITK.on('add', onPath => {
		++FileCount;
		readDirectory(onPath, inputDirectoryITK, logsDirectoryITK, outputDirectoryITK, FileCount);
	})

	// start watcherDnD when there is file in inputITK folder
	watcherDnD.on('add', onPath => {
		++FileCount;
		readDirectory(onPath, inputDirectoryDnD, logsDirectoryDnD, outputDirectoryDnD, FileCount);
	})

}
catch (err) {
	const nameInfo = path.parse(fileName);
	fs.writeFile(`${outputDirectoryDnD}/${nameInfo.name}_ERROR_${count}.txt`,
		`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - ${err}`,
		{ flag: 'a' },
		(error) => { })
	fs.readFile(`${outputDirectoryDnD}/${nameInfo.name}_ERROR_${count}.txt`, 'utf8', function (err, data) {
		//upload data to database
		// sql.connect(sqlConfig).then(pool => {
		// 	// Query
		// 	return pool.request()
		// 		.query(`INSERT INTO dbo.xml_generator VALUES ('${fileName}','${format(new Date(), 'yyyy-MM-dd-hh:mm:ss')}','1','${data}')`)
		// }).then(result => {
		// 	console.dir(result)
		// }).catch(err => {
		// 	// ... error checks
		// 	console.log(err);
		// });
	});
}