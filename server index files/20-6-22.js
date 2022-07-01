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

const app = express()
const port = 6000;
app.listen(port);

let sqlConfig = {
	user: 'SVC-OBIEE',
	password: 'Lexis!Te1st',
	server: 'LNGSYDAPPD026',
	database: 'ProposalGenerator',
	connectionTimeout: 30000,
	requestTimeout: 30000,
	dialect: 'mssql',
	port: 1433,
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
		stabilityThreshold: 9000,
		pollInterval: 3000
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
		stabilityThreshold: 9000,
		pollInterval: 3000
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
			let result = '';
			if (d == 'Invalid Date') {
				// const setdate = d;
				// console.log(date);
				if (date.includes('-')) {
					const [day, month, year] = date.split('-');

					result = [year, month, day].join('-');

					// const ans = [result, time].join(' ');
					// console.log(result);
				}
				else if (date.includes('/')) {
					const [day, month, year] = date.split('/');

					result = [year, month, day].join('-');

					// const ans = [result, time].join(' ');
					// console.log(result);
				}
				return result;
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
				.query(`INSERT INTO xml_generator VALUES ('${fileName.replace(/'/g, "''")}','${format(new Date(), 'yyyy-MM-dd-hh:mm:ss')}',${success ? '1' : '0'},'${data.replace(/'/g, "''")}')`)
		}).then(result => {
			console.log(result)
		}).catch(err => {
			// ... error checks
			console.log(err);
		});
	}

	function isColCountCorrect(fileRowJson, fileName, outputDirectory, inputDirectory, logsDirectory) {
		// we considering fixed structure of excel file
		// means columns ordering will be same for all input excel files
		const len = Object.keys(fileRowJson).length;
		if (len === 18)
			return true;
		else {

			console.log("File moved to log => files");
			fs.copyFileSync(`${inputDirectory}/${fileName}`, `${logsDirectory}/files/${fileName}`)
			fs.unlinkSync(`${inputDirectory}/${fileName}`);

			fs.writeFileSync(`${outputDirectory}/${format(new Date(), 'yyyy-MM-dd')}_${fileName}_ERROR.txt`,
				`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - ${fileName} | The number of columns in the Excel file is not correct. There should be exactly 18 columns in file.`,
				{ flag: 'a' },
				(error) => { })

			fs.readFile(`${outputDirectory}/${format(new Date(), 'yyyy-MM-dd')}_${fileName}_ERROR.txt`, 'utf8', function (err, data) {
				//upload data to database
				insertDataInDB(fileName, true, data)
			});
			return false;
		}
	}


	function readDirectory(directoryPath, inputDirectory, logsDirectory, outputDirectory) {
		// console.log(onPath, 'File added');
		console.log("Process started");
		fs.readdir(directoryPath, function (err, files) {
			if (err) {
				fs.writeFile(`${outputDirectory}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`,
					`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - Unable to scan or read Input Directory`,
					{ flag: 'a' },
					(error) => {
						// In case of a error throw err.
					})
			}

			for (let i in files) {
				try {
					let parseXML = true;
					if (ContinueToNextFile(files[i]))
						continue;
					const result = readExcelFile(files[i], inputDirectory, outputDirectory);

					if (!isColCountCorrect(result[0], files[i], outputDirectory, inputDirectory, logsDirectory)) {
						parseXML = false;
						continue;
					}

					const emailRowToDel = [];
					const emptyCustRef = [];
					const billFreq = [];
					const date = [];
					const billFreqRegex = /weekly|monthly/i;

					const transIDDuplicate = [];
					const transIDSet = new Set();
					result.forEach(function (row, rowno) {
						Object.keys(row).forEach(function (key, i) {
							if (i === 2 && row[key] === '') {
								emptyCustRef.push(row);
							}
							if (i === 3 && row[key].includes('@lexisnexis.com')) {
								console.log(row[key]);
								emailRowToDel.push(row);
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
							// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
							fs.writeFile(`${outputDirectory}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`,
								`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${files[i]} | File has empty customer reference for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}. `,
								{ flag: 'a' },
								(error) => { })
						});
					}

					// if (date.length > 0) {
					// 	parseXML = false;
					// 	// notify user in output folder that error is logged in logs folder 
					// 	fs.writeFile(`${outputDirectoryITK}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`,
					// 		`\n${files[i]} - errors are logged in folder path : ${__dirname}\\logsITK`,
					// 		{ flag: 'a' },
					// 		(error) => { })
					// 	date.forEach(function (row) {
					// 		const indexToDel = result.indexOf(row);
					// 		if (indexToDel !== -1) {
					// 			// result.splice(indexToDel, 1);
					// 			// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
					// 			fs.writeFile(`${outputDirectoryITK}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`,
					// 				`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${files[i]} | File has invalid date for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}. `,
					// 				{ flag: 'a' },
					// 				(error) => { })
					// 		}
					// 	});
					// }

					if (emailRowToDel.length > 0) {
						parseXML = false;
						emailRowToDel.forEach(function (row) {
							// writing email row as (index+2) because array index starts from 0 and excel file has first row of headers
							fs.writeFile(`${outputDirectory}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`,
								`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${files[i]} | File has a email = ${row['ACCOUNT_EMAIL']} that includes lexisnexis.com domain for username = ${row['USER_NAME']}. .`,
								{ flag: 'a' },
								(error) => { });
						});
					}
					if (billFreq.length > 0) {
						parseXML = false;
						billFreq.forEach(function (row) {
							// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
							fs.writeFile(`${outputDirectory}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`,
								`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${files[i]} | File has Billing_Frequency = ${row['BILLING_FREQUENCY']} which is not weekly or monthly for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`,
								{ flag: 'a' },
								(error) => { })
						});
					}
					if (transIDDuplicate.length > 0) {
						parseXML = false;
						transIDDuplicate.forEach(function (row) {
							// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
							fs.writeFile(`${outputDirectory}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`,
								`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${files[i]} | File has Duplicate Transaction ID = ${row['TRANSACTIONID']} for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}.`,
								{ flag: 'a' },
								(error) => { })
						});
					}
					if (parseXML) {
						let xml = jsonxml({ 'ORDER_DATA': result }, { xmlHeader: { standalone: true } })
						xml = vkbeautify.xml(xml, 4);
						fs.writeFileSync(`${outputDirectory}/${path.parse(files[i]).name.replace(/\s/g, '')}.xml`, xml, { flag: 'w' }) //remove spaces from output file name
						console.log(`Output file created`);
						//delete a file from a folder
						const delFile = `${inputDirectory}/${files[i]}`;
						fs.unlinkSync(delFile);

						// console.log(`file_name : ${files[i]}`);
						// console.log(`date_time : ${format(new Date(), 'yyyy-MM-dd-hh:mm:ss')}`);

						//upload data to database
						insertDataInDB(files[i], false, '')

					}
					else {
						console.log("File moved to log => files");
						// console.log(`${outputDirectoryITK}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`);
						fs.copyFileSync(`${inputDirectory}/${files[i]}`, `${logsDirectory}/files/${files[i]}`)
						fs.unlinkSync(`${inputDirectory}/${files[i]}`);
						fs.readFile(`${outputDirectory}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`, 'utf8', function (err, data) {
							//upload data to database
							insertDataInDB(files[i], true, data)
						});
					}
				}
				catch (err) {
					fs.writeFileSync(`${outputDirectory}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`,
						`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - ${files[i]} | ${err}`,
						{ flag: 'a' },
						(error) => { })
					fs.readFile(`${outputDirectory}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`, 'utf8', function (err, data) {
						//upload data to database
						insertDataInDB(files[i], true, data)
					});
				}
			}

		});
	}

	// start watcherITK when there is file in inputITK folder
	watcherITK.on('add', onPath => {
		readDirectory(directoryPathITK, inputDirectoryITK, logsDirectoryITK, outputDirectoryITK);
	})

	// start watcherDnD when there is file in inputITK folder
	watcherDnD.on('add', onPath => {
		readDirectory(directoryPathDnD, inputDirectoryDnD, logsDirectoryDnD, outputDirectoryDnD);
	})

}
catch (err) {
	fs.writeFile(`${outputDirectoryDnD}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`,
		`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - ${err}`,
		{ flag: 'a' },
		(error) => { })
	fs.readFile(`${outputDirectoryDnD}/${format(new Date(), 'yyyy-MM-dd')}_${files[i]}_ERROR.txt`, 'utf8', function (err, data) {
		//upload data to database
		// sql.connect(sqlConfig).then(pool => {
		// 	// Query
		// 	return pool.request()
		// 		.query(`INSERT INTO dbo.xml_generator VALUES ('${files[i]}','${format(new Date(), 'yyyy-MM-dd-hh:mm:ss')}','1','${data}')`)
		// }).then(result => {
		// 	console.dir(result)
		// }).catch(err => {
		// 	// ... error checks
		// 	console.log(err);
		// });
	});
}