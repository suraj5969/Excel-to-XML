/*changes done 
1. notify user that error are logged in log folder
2. remove space from XML file name
3. validate date
4. Match row count of excel file and xml file (demo.js)
5. 2 input and 2 output
*/

const cron = require('node-cron');
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
const config = require('./dbconfig.js');

const app = express()
const port = 6000;
app.listen(port);

let sqlConfig = config;
const providers = {
	itk: 'InfoTrack',
	dnd: 'Dye & Durham',
}

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

// const dndeaders = [/username/i,]
// const dndheaders = ["USER_NAME", "CONTACT", "CUSTOMER_REFERENCE", "ACCOUNT_EMAIL", "PRODUCT_ID", "PRODUCT_GROUP", "PRODUCT
// dndeaders.forEach((regex, i) => {
// 	if (regex.test(dndheaders[i])){
// 		headers.push(regex);
// 	}
// })

try {

	function createDirectories(outputDirectory, logsDirectory) {
		if (!fs.existsSync(outputDirectory)) {
			fs.mkdirSync(outputDirectory);
		}
		if (!fs.existsSync(logsDirectory)) {
			fs.mkdirSync(logsDirectory);
		}
		// if (!fs.existsSync(`${logsDirectory}/files`)) {
		// 	fs.mkdirSync(`${logsDirectory}/files`);
		// }
		if (!fs.existsSync(`${logsDirectory}/Success`)) {
			fs.mkdirSync(`${logsDirectory}/Success`);
		}
		if (!fs.existsSync(`${logsDirectory}/Fail`)) {
			fs.mkdirSync(`${logsDirectory}/Fail`);
		}
		if (!fs.existsSync(`${logsDirectory}/XML`)) {
			fs.mkdirSync(`${logsDirectory}/XML`);
		}
	}

	function removeHiddenChars(str) {
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
		//if date would be string then convert it to date
		// but strngs of dd-mm-yy foramt are not supported by Date
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

	// function moveFile() {
	// 	for (let i = 0; true; i++) {
	// 		if (fs.existsSync(`${inputDirectoryITK}/file path`)) {
	// 			console.log('ITK folder exists');
	// 			// continue;
	// 		}
	// 		else {
	// 			//move file with to output with count in it
	// 			if (i === 0) {
	// 				// move file without renaming
	// 			}
	// 			else {
	// 				// move file with renaming (adding i to end of file name)
	// 			}
	// 			break;
	// 		}
	// 	}
	// }

	function writeGeneralError(fileName, outputDirectory, fileCount, errorStr) {
		const nameInfo = path.parse(fileName);
		fs.writeFileSync(`${outputDirectory}/${nameInfo.name}_ERROR_${fileCount}.txt`,
			`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | ${errorStr}`,
			{ flag: 'a' })
	}

	function wirteErrorWithRowNo(fileName, outputDirectory, fileCount, result, row, errorStr) {
		// let a = JSON.stringify(result)
		// let b = JSON.stringify(row)
		// console.log(row);
		const rowno = result.indexOf(row);
		// console.log('row no', rowno);
		// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
		writeGeneralError(fileName, outputDirectory, fileCount, `${errorStr} at row no. ${rowno + 2}`);
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
			//if string would be dd-mm-yy format then convert it to date
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
				return formatDate(result, false, false);
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

	function insertDataInDB(fileName, success, data, provider, exportFileName, xmlFileName, exportXmlFileName) {
		return
		function replaceQuote(str) {
			return str.replace(/'/g, "''");
		}
		sql.connect(sqlConfig).then(pool => {
			return pool.request()
				.query(`INSERT INTO xml_generator ([file_name],[date_time] ,[is_success] ,[error_logs] ,[provider] ,[changed_excel_file_name], [xml_file_name], [changed_xml_file_name])
				 VALUES ('${replaceQuote(fileName)}','${format(new Date(), 'yyyy-MM-dd-hh:mm:ss')}',${success ? '1' : '0'},'${replaceQuote(data)}','${replaceQuote(provider)}', '${replaceQuote(exportFileName)}', '${replaceQuote(xmlFileName)}', '${replaceQuote(exportXmlFileName)}')`)
		}).then(result => {
			console.log(result)
		}).catch(err => {
			// ... error checks
			console.log(err);
		});
	}

	function readExcelFile(fileName, fileDirectory, outputDirectory) {
		const workbook = XLSX.readFile(`${fileDirectory}/${fileName}`);
		const sheet_name_list = workbook.SheetNames;

		if (sheet_name_list.length > 1) {
			// fs.renameSync(`${fileDirectory}/${fileName}`, `${outputDirectory}/${fileName}`)
			fs.writeFile(`${outputDirectory}/${fileName}_warning.txt`,
				`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - ${fileName} | File has more than one sheets. Continuing to process only first sheet in file.`,
				{ flag: 'a' },
				(error) => { })
		}
		const result = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]], { defval: "" });
		return result;
	}

	function moveLogFile(file, outputDirectory, logsDirectory) {
		const { birthtime } = fs.statSync(`${outputDirectory}/${file}`)
		const nowDate = new Date();
		let Difference_In_Time = nowDate.getTime() - birthtime.getTime();
		let Difference_In_Days = Difference_In_Time / (1000 * 3600 * 24);
		if (Math.trunc(Difference_In_Days) >= 5) {
			fs.renameSync(`${outputDirectory}/${file}`, `${logsDirectory}/${file}`)
		}
	}

	function checkColPosition(fileRowJson, fileName, outputDirectory, inputDirectory, logsDirectory, count, provider) {
		const headers = Object.keys(fileRowJson)
		//&& /contact/i.test(headers[1]) && /customer ref/i.test(headers[2]) && /account email/i.test(headers[3]) && /product id/i.test(headers[4]) && /product group/i.test(headers[5]) && /product description/i.test(headers[6]) && /transaction id/i.test(headers[7]) && /date/i.test(headers[8]) && /reference/i.test(headers[9]) && /matter/i.test(headers[10]) && /request id/i.test(headers[11]) && /price(ex gst)/i.test(headers[12]) && /price(inc gst)/i.test(headers[13]) && /gst/i.test(headers[14]) && /billing frequency/i.test(headers[15]) && /retailer reference/i.test(headers[16]) && /period ending/i.test(headers[17])
		if (provider === providers.dnd && (/user id/i).test(headers[0]) && /contact/i.test(headers[1]) && /customer ref/i.test(headers[2]) && /account email/i.test(headers[3]) && /product id/i.test(headers[4]) && /product group/i.test(headers[5]) && /product description/i.test(headers[6]) && /transaction id/i.test(headers[7]) && /date/i.test(headers[8]) && /reference/i.test(headers[9]) && /matter/i.test(headers[10]) && /request id/i.test(headers[11]) && /price\(ex gst\)/i.test(headers[12]) && /price\(inc gst\)/i.test(headers[13]) && /gst/i.test(headers[14]) && /billing frequency/i.test(headers[15]) && /retailer reference/i.test(headers[16]) && /period ending/i.test(headers[17])) {
			return true;
		}
		if (provider === providers.itk && (/username/i).test(headers[0]) && /contact/i.test(headers[1]) && /customer reference/i.test(headers[2]) && /account email/i.test(headers[3]) && /product id/i.test(headers[4]) && /product group/i.test(headers[5]) && /product description/i.test(headers[6]) && /order id/i.test(headers[7]) && /date/i.test(headers[8]) && /reference/i.test(headers[9]) && /matter/i.test(headers[10]) && /request id/i.test(headers[11]) && /price ex gst/i.test(headers[12]) && /price inc gst/i.test(headers[13]) && /gst/i.test(headers[14]) && /billing frequency/i.test(headers[15]) && /retailer reference/i.test(headers[16]) && /period ending/i.test(headers[17])) {
			return true;
		}
		else {
			console.log(`Files moved to ${logsDirectory}/Fail`);
			const nameInfo = path.parse(fileName);
			fs.renameSync(`${inputDirectory}/${fileName}`, `${logsDirectory}/Fail/${nameInfo.name}_${count}${nameInfo.ext}`)
			// fs.unlinkSync(`${inputDirectory}/${fileName}`);
			const err = `The Columns in the excel file are not in proper arrangement`
			writeGeneralError(fileName, outputDirectory, count, err);
			fs.readFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`, 'utf8', function (err, data) {
				//upload data to database
				insertDataInDB(fileName, false, data, provider, `${nameInfo.name}_${count}${nameInfo.ext}`, '', '')
			});
			return false;
		}
	}

	function isColCountCorrect(fileRowJson, fileName, outputDirectory, inputDirectory, logsDirectory, count, provider) {
		// we considering fixed structure of excel file
		// means columns ordering will be same for all input excel files

		const len = Object.keys(fileRowJson).length;
		if (len === 18)
			return true;
		else {
			console.log(`Files moved to ${logsDirectory}/Fail`);
			const nameInfo = path.parse(fileName);
			fs.renameSync(`${inputDirectory}/${fileName}`, `${logsDirectory}/Fail/${nameInfo.name}_${count}${nameInfo.ext}`)
			// fs.unlinkSync(`${inputDirectory}/${fileName}`);
			const err = `The number of columns in the Excel file is not correct. There should be exactly 18 columns in file.`
			writeGeneralError(fileName, outputDirectory, count, err);
			// fs.writeFileSync(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`,
			// 	`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - ${fileName} | The number of columns in the Excel file is not correct. There should be exactly 18 columns in file.`,
			// 	{ flag: 'a' },
			// 	(error) => { })

			fs.readFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`, 'utf8', function (err, data) {
				//upload data to database
				insertDataInDB(fileName, false, data, provider, `${nameInfo.name}_${count}${nameInfo.ext}`, '', '')
			});
			return false;
		}
	}

	// to move files from output directory to logs directory
	cron.schedule('0 */6 * * * ', () => {
		fs.readdirSync(outputDirectoryDnD).forEach(fileName => {
			if (fileName.includes('_ERROR') || fileName.includes('_warning.txt')) {
				moveLogFile(`${fileName}`, outputDirectoryDnD, logsDirectoryDnD);
			}
		});

		fs.readdirSync(outputDirectoryITK).forEach(fileName => {
			if (fileName.includes('_ERROR') || fileName.includes('_warning.txt')) {
				moveLogFile(`${fileName}`, outputDirectoryITK, logsDirectoryITK);
			}
		});
	})

	function processFile(filePath, inputDirectory, logsDirectory, outputDirectory, count, provider) {
		try {
			console.log("Process started");
			createDirectories(outputDirectory, logsDirectory);
			const fileName = path.basename(filePath);
			let parseXML = true;
			if (ContinueToNextFile(fileName))
				return;
			const result = readExcelFile(fileName, inputDirectory, outputDirectory, provider);


			if (!isColCountCorrect(result[0], fileName, outputDirectory, inputDirectory, logsDirectory, count, provider)) {
				// parseXML = false;
				return;
			}
			if (!checkColPosition(result[0], fileName, outputDirectory, inputDirectory, logsDirectory, count, provider)) {
				// parseXML = false;
				return;
			}

			const emailRowToDel = [];
			const emptyCustRef = [];
			const billFreq = [];
			const date = [];
			const billFreqRegex = /weekly|monthly/i;
			const urnRegex = /^urn:ecn:/i;  // ^ is used to check for start of string and i is used to ignore case
			const emailRegex = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
			const email = [];
			const beforeGst = [];
			const afterGst = [];
			const gst = [];
			const transIDDuplicateDnD = [];
			const transIDDuplicateITK = [];
			const twoDuplicateTransITK = [];
			const transIDSet = new Set();
			const transIDMap = new Map();
			result.forEach(function (row, rowno) {
				Object.keys(row).forEach(function (key, i) {
					if (i === 2 && (row[key] === '' || urnRegex.test(row[key]))) {
						emptyCustRef.push(row);
					}
					if (i === 3) {
						if (!emailRegex.test(row[key])) {
							let multiEmail = row[key].split(';');
							multiEmail.forEach(function (email) {
								if (!emailRegex.test(email)) {
									emailRowToDel.push(row);
								}
							})
						}
						if (row[key].includes('@lexisnexis.com')) {
							// console.log(row[key]);
							emailRowToDel.push(row);
	
						}
					}
					if (i === 12) {
						row[key] = Number(row[key]);
						if (isNaN(row[key])) {
							// console.log(row[key]);
							beforeGst.push(row);
						}
					}
					if (i === 13) {
						row[key] = Number(row[key]);
						if (isNaN(row[key])) {
							// console.log(row[key]);
							afterGst.push(row);
						}
					}
					if (i === 14) {
						row[key] = Number(row[key]);
						if (isNaN(row[key])) {
							// console.log(row[key]);
							gst.push(row);
						}
					}
					if (i === 15) {
						//because they want to only include Weekly in output not monthly
						if (billFreqRegex.test(String(row[key]).trim())) {
							row[key] = 'Weekly';
						}
						else {
							billFreq.push(row);
						}
					}
					if (i === 7) {
						if (provider === providers.dnd) {
							if (transIDSet.has(row[key])) {
								transIDDuplicateDnD.push(row);
							} else {
								transIDSet.add(row[key]);
							}
						}
						if (provider === providers.itk) {
							let value = transIDMap.get(row[key]);
							if (value) {
								// console.log(value);
								if (value === 2) {
									transIDDuplicateITK.push(row);
								}
								if(value === 1){
									twoDuplicateTransITK.push(row);
								}
								transIDMap.set(row[key], 2)
							}
							else {
								transIDMap.set(row[key], 1)
							}
							
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
						if (row[key].includes('Invalid Date')) {
							date.push(row);
						}
					}

					let val = row[key];
					if (typeof val === 'number') {
						//round the number to 2 decimal places as sometimes excel may give a number with more than 2 decimal places
						// for example: excel may show a number 1.80 in excel app but here it sometimes gives 1.7999999999999998
						val = Math.round((val + Number.EPSILON) * 100) / 100;
					}
					else {
						val = removeHiddenChars(val);
					}

					// Remove key-value from object
					delete row[key];
					// Add value with new key
					// insted of creating a new object we are modifying the original object got from excel file
					row[headers[i]] = val;
				});
			});

			if (emptyCustRef.length > 0) {
				parseXML = false;
				emptyCustRef.forEach(function (row) {
					const err = `File has empty customer reference or does not starts with "urn:ecm:" for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}. `;
					wirteErrorWithRowNo(fileName, outputDirectory, count, result, row, err);
				});
			}
			if (beforeGst.length > 0) {
				parseXML = false;
				beforeGst.forEach(function (row) {
					const err = `File does not have a valid number in (Price ex GST) column for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}. Expected number.`
					wirteErrorWithRowNo(fileName, outputDirectory, count, result, row, err);
				});
			}
			if (afterGst.length > 0) {
				parseXML = false;
				afterGst.forEach(function (row) {
					const err = `File does not have a valid number in (Price inc GST) column for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}. Expected number.`
					wirteErrorWithRowNo(fileName, outputDirectory, count, result, row, err);
				});
			}
			if (gst.length > 0) {
				parseXML = false;
				gst.forEach(function (row) {
					const err = `File does not have a valid number in (GST) column for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}. Expected number.`
					wirteErrorWithRowNo(fileName, outputDirectory, count, result, row, err);
				});
			}

			if (date.length > 0) {
				parseXML = false;
				date.forEach(function (row) {
					const err = `File has invalid date in "Date" or "Period End Date" column for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`;
					wirteErrorWithRowNo(fileName, outputDirectory, count, result, row, err);
				});
			}

			if (emailRowToDel.length > 0) {
				parseXML = false;
				emailRowToDel.forEach(function (row) {
					const nameInfo = path.parse(fileName);
					// writing email row as (index+2) because array index starts from 0 and excel file has first row of headers
					const err = `File has a email = ${row['ACCOUNT_EMAIL']} that incudes lexisnexis.com domain for username = ${row['USER_NAME']}`;
					wirteErrorWithRowNo(fileName, outputDirectory, count, result, row, err);
					// fs.writeFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`,
					// 	`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | File has a email = ${row['ACCOUNT_EMAIL']} that includes lexisnexis.com domain for username = ${row['USER_NAME']}. .`,
					// 	{ flag: 'a' },
					// 	(error) => { });
				});
			}
			if (email.length > 0) {
				parseXML = false;
				email.forEach(function (row) {
					// writing email row as (index+2) because array index starts from 0 and excel file has first row of headers
					const err = `File has a email = ${row['ACCOUNT_EMAIL']} that is not in valid email format for username = ${row['USER_NAME']}`;
					wirteErrorWithRowNo(fileName, outputDirectory, count, result, row, err);
					// fs.writeFile(`${outputDirectory}/${count}_${fileName}_ERROR.txt`,
					// 	`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | File has a email = ${row['ACCOUNT_EMAIL']} that is not in valid email format for username = ${row['USER_NAME']}. .`,
					// 	{ flag: 'a' },
					// 	(error) => { });
				});
			}
			if (billFreq.length > 0) {
				parseXML = false;
				billFreq.forEach(function (row) {
					const nameInfo = path.parse(fileName);
					// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
					const err = `File has Billing_Frequency = ${row['BILLING_FREQUENCY']} which is not weekly or monthly for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`;
					wirteErrorWithRowNo(fileName, outputDirectory, count, result, row, err);
					// fs.writeFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`,
					// 	`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | File has Billing_Frequency = ${row['BILLING_FREQUENCY']} which is not weekly or monthly for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`,
					// 	{ flag: 'a' },
					// 	(error) => { })
				});
			}

			if (transIDDuplicateDnD.length > 0) {
				parseXML = false;
				transIDDuplicateDnD.forEach(function (row) {
					// console.log(row);

					const nameInfo = path.parse(fileName);
					// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
					const err = `(DND) File has duplicate Transaction_ID = ${row['TRANSACTIONID']} for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`;
					wirteErrorWithRowNo(fileName, outputDirectory, count, result, row, err);
					// fs.writeFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`,
					// 	`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | File has Duplicate Transaction ID = ${row['TRANSACTIONID']} for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`,
					// 	{ flag: 'a' },
					// 	(error) => { })
				});
			}

			if (transIDDuplicateITK.length > 0) {
				parseXML = false;
				transIDDuplicateITK.forEach(function (row) {
					// console.log(row);

					const nameInfo = path.parse(fileName);
					// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
					const err = `(ITK) File has duplicate Transaction_ID = ${row['TRANSACTIONID']} for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`;
					wirteErrorWithRowNo(fileName, outputDirectory, count, result, row, err);
					// fs.writeFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`,
					// 	`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | File has Duplicate Transaction ID = ${row['TRANSACTIONID']} for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`,
					// 	{ flag: 'a' },
					// 	(error) => { })
				});
			}

			if (twoDuplicateTransITK.length > 0) {
				// parseXML = true;
				twoDuplicateTransITK.forEach(function (row) {
					// console.log(row);

					const nameInfo = path.parse(fileName);
					// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
					// const err = `(ITK) File has duplicate Transaction_ID = ${row['TRANSACTIONID']} for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`;
					// wirteErrorWithRowNo(fileName, outputDirectory, count, result, row, err);
					fs.writeFile(`${outputDirectory}/${nameInfo.name}_${count}_warning.txt`,
						`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${fileName} | File has Duplicate Transaction ID = ${row['TRANSACTIONID']} for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`,
						{ flag: 'a' },
						(error) => { })
				});
			}

			if (parseXML) {
				// count++;
				const res = result.map(row => ({ "LN_ITK_GBX_TBL": row }))
				let xml = jsonxml({ 'ORDER_DATA': res }, { xmlHeader: { standalone: true } })
				xml = vkbeautify.xml(xml, 4);
				for (let i = 0; true; i++) {
					let endFileCount = '';
					if (i === 0) {
						endFileCount = '';
					}
					else {
						endFileCount = `(${i})`;
					}
					if (fs.existsSync(`${outputDirectory}/${path.parse(fileName).name.replace(/\s/g, '')}${endFileCount}.xml`)) {
						console.log('File already exist in output folder');
						// continue;
					}
					else {
						//move file to output with count in it
						fs.writeFileSync(`${outputDirectory}/${path.parse(fileName).name.replace(/\s/g, '')}${endFileCount}.xml`, xml, { flag: 'w' }) //remove spaces from output file name
						break;
					}
				}
				// fs.writeFileSync(`${outputDirectory}/${path.parse(fileName).name.replace(/\s/g, '')}.xml`, xml, { flag: 'w' }) //remove spaces from output file name
				console.log(`Output file created`);
				//delete a file from a folder
				//copy XML file in logs folder
				// const delFile = `${inputDirectory}/${fileName}`;
				const nameInfo = path.parse(fileName);
				fs.renameSync(`${inputDirectory}/${fileName}`, `${logsDirectory}/Success/${nameInfo.name}_${count}${nameInfo.ext}`)
				fs.copyFileSync(`${outputDirectory}/${path.parse(fileName).name.replace(/\s/g, '')}.xml`, `${logsDirectory}/XML/${path.parse(fileName).name.replace(/\s/g, '')}_${count}.xml`);

				//upload data to database
				insertDataInDB(fileName, true, '', provider, `${nameInfo.name}_${count}${nameInfo.ext}`, `${path.parse(fileName).name.replace(/\s/g, '')}.xml`, `${path.parse(fileName).name.replace(/\s/g, '')}_${count}.xml`)
			}
			else {
				// count++;
				console.log(`Files moved to ${logsDirectory}/Fail`);
				const nameInfo = path.parse(fileName);
				// console.log(`${outputDirectoryITK}/${format(new Date(), 'yyyy-MM-dd')}_${fileName}_ERROR.txt`);
				fs.renameSync(`${inputDirectory}/${fileName}`, `${logsDirectory}/Fail/${nameInfo.name}_${count}${nameInfo.ext}`)
				// fs.unlinkSync(`${inputDirectory}/${fileName}`);
				fs.readFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`, 'utf8', function (err, data) {
					//upload data to database
					insertDataInDB(fileName, false, data, provider, `${nameInfo.name}_${count}${nameInfo.ext}`, '', '')
				});
			}
		}
		catch (err) {
			// console.log(err);
			const nameInfo = path.parse(filePath);
			writeGeneralError(filePath, outputDirectory, count, err);
			fs.readFile(`${outputDirectory}/${nameInfo.name}_ERROR_${count}.txt`, 'utf8', function (err, data) {
				//upload data to database
				insertDataInDB(filePath, false, data, provider, `${nameInfo.name}_${count}${nameInfo.ext}`, '', '')
			});
		}

	}

	// start watcherITK when there is file in inputITK folder
	let FileCount = 0;
	// if(FileCount === 0){
	// 	console.log('please see if Filecount value is correct it is 0. Increment it if needed.')
	// 	process.exit(1)
	// }
	watcherITK.on('add', onPath => {
		++FileCount;
		processFile(onPath, inputDirectoryITK, logsDirectoryITK, outputDirectoryITK, FileCount, providers.itk);
	})

	// start watcherDnD when there is file in inputITK folder
	watcherDnD.on('add', onPath => {
		++FileCount;
		processFile(onPath, inputDirectoryDnD, logsDirectoryDnD, outputDirectoryDnD, FileCount, providers.dnd);
	})

}
catch (err) {
	// the error thrown here will not be for file but will be on process level
	// like not listning to folder or not having permission to access folder 
	// or error with chokidar or functions not defined properly
	insertDataInDB('Error in XML Generator process', false, String(err), '', '', '', '');
}