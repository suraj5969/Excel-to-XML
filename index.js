
const cron = require('node-cron');
const fs = require('fs')
const jsonxml = require('jsontoxml');
const vkbeautify = require('vkbeautify');
const XLSX = require('xlsx');
const { getJsDateFromExcel } = require("excel-date-to-js");
const { format, isValid } = require('date-fns');
const path = require('path');

const inputDirectory = './input';
const outputDirectory = './output';
const logsDirectory = './logs';
try {
	const directoryPath = path.join(__dirname, inputDirectory);

	if (!fs.existsSync('./output')) {
		fs.mkdirSync('output');
	}
	if (!fs.existsSync('./logs')) {
		fs.mkdirSync('logs');
	}

	//cron job for every 30 minutes
	// cron.schedule('*/30 * * * * *', () => {
	console.log("Process started");

	fs.readdir(directoryPath, function (err, files) {
		if (err) {
			fs.writeFile(`${logsDirectory}/error.txt`,
				`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - Unable to scan or read Input Directory`,
				{ flag: 'a' },
				(error) => {
					// In case of a error throw err.
				})
		}

		const removeSpecialChars = (str) => {
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

		const formatDate = (date, isTime, toUTC) => {
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

		const parseExcelDate = (date, isTime) => {
			try {
				let d = getJsDateFromExcel(date)
				d = formatDate(d, isTime, true);
				return d;
			}
			catch (e) {
				// if getJsDateFromExcel throws a error when this field is a string but valid date
				const d = formatDate(date, isTime, false);
				return d;
			}
		}

		// check how many extentions XLSX support to read
		const ext = ['.xlsx', '.xlsm', '.xls'];
		for (let i in files) {
			try {
				//check if file is an excel file and not error.txt and not a hidden file if the same file is opened in excel app
				if (files[i] === 'error.txt' || files[i].startsWith('~$') || !(ext.includes(path.extname(files[i]))))
					continue;
				const workbook = XLSX.readFile(`${inputDirectory}/${files[i]}`);
				const sheet_name_list = workbook.SheetNames;

				if (sheet_name_list.length > 1) {
					fs.writeFile(`${logsDirectory}/warning.txt`,
						`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - ${files[i]} | File has more than one sheets. Continuing to process only first sheet in file.`,
						{ flag: 'a' },
						(error) => { })
				}
				const result = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]], { defval: "" });

				const headers = ["USER_NAME", "CONTACT", "CUSTOMER_REFERENCE", "ACCOUNT_EMAIL", "PRODUCT_ID", "PRODUCT_GROUP", "PRODUCT_DESCRIPTION", "TRANSACTIONID", "DATE", "REFERENCE", "MATTER", "REQUEST_ID", "AMT_BEFORE_TAX", "AMT_AFTER_TAX", "GST_AMT", "BILLING_FREQUENCY", "RETAILER_REFERENCE", "PERIOD_END_DT"];
				const emailRowToDel = [];
				const emptyCustRef = [];
				const billFreq = [];
				const billFreqRegex = /weekly|monthly/i;

				const transIDDuplicate = [];
				const transIDSet = new Set();
				// we considering fixed structure of excel file
				// means columns ordering will be same for all input excel files
				const colCount = Object.keys(result[0]).length;
				result.forEach(function (row, rowno) {
					Object.keys(row).forEach(function (key, i) {
						if (i === 2 && row[key] === '') {
							emptyCustRef.push(row);
						}
						if (i === 3 && row[key].includes('@lexisnexis.com')) {
							emailRowToDel.push(row);
						}
						if (i === 15 && !billFreqRegex.test(row[key])) {
							billFreq.push(row);
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
					if (colCount < headers.length) {
						for (let i = colCount; i < headers.length; i++) {
							row[headers[i]] = '';
						}
					}
					// result[rowno] = { "LN_ITK_GBX_TBL": row }
				});

				//remove rows which are having lexisnexis.com email address
				emailRowToDel.forEach(function (row) {
					const indexToDel = result.indexOf(row);
					if (indexToDel !== -1) {
						result.splice(indexToDel, 1);
						// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
						fs.writeFile(`${logsDirectory}/warning.txt`,
							`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${files[i]} | File has a email = ${row['ACCOUNT_EMAIL']} that has lexisnexis.com domain for username = ${row['USER_NAME']}. Removing the row from file.`,
							{ flag: 'a' },
							(error) => { })
					}
				});
				// empty customer reference
				emptyCustRef.forEach(function (row) {
					const indexToDel = result.indexOf(row);
					if (indexToDel !== -1) {
						result.splice(indexToDel, 1);
						// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
						fs.writeFile(`${logsDirectory}/warning.txt`,
							`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${files[i]} | File has empty customer reference for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}. Removing the row from file.`,
							{ flag: 'a' },
							(error) => { })
					}
				});
				// billing frequency
				billFreq.forEach(function (row) {
					const indexToDel = result.indexOf(row);
					if (indexToDel !== -1) {
						result.splice(indexToDel, 1);
						// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
						fs.writeFile(`${logsDirectory}/warning.txt`,
							`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${files[i]} | File has Billing_Frequency = ${row['BILLING_FREQUENCY']} which is not weekly or monthly for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}. Removing the row from file.`,
							{ flag: 'a' },
							(error) => { })
					}
				});
				// Dup[licate Transcation ID
				transIDDuplicate.forEach(function (row) {
					const indexToDel = result.indexOf(row);
					if (indexToDel !== -1) {
						result.splice(indexToDel, 1);
						// writing email row as (index+2) because array index starts from 0 and excel file has first how of headers
						fs.writeFile(`${logsDirectory}/warning.txt`,
							`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - in ${files[i]} | File has Duplicate Transaction ID = ${row['TRANSACTIONID']} for username = ${row['USER_NAME']} and email = ${row['ACCOUNT_EMAIL']}. Removing the row from file.`,
							{ flag: 'a' },
							(error) => { })
					}
				});

				const finalResult = result.map(row => ({ "LN_ITK_GBX_TBL": row }));
				let xml = jsonxml({ 'ORDER_DATA': finalResult }, { xmlHeader: { standalone: true } })
				xml = vkbeautify.xml(xml, 4);
				fs.writeFileSync(`${outputDirectory}/${path.parse(files[i]).name}.xml`, xml, { flag: 'w' })
				console.log("Output file created");

				//delete a file from a folder
				const delFile = `${inputDirectory}/${files[i]}`;
				// fs.unlinkSync(delFile);
			}
			catch (err) {
				fs.writeFile(`${logsDirectory}/error.txt`,
					`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - ${files[i]} | ${err}`,
					{ flag: 'a' },
					(error) => { })
			}
		}

	});
}
catch (err) {
	fs.writeFile(`${logsDirectory}/error.txt`,
		`\n${format(new Date(), 'yyyy-MM-dd:hh:mm:ss')} - ${err}`,
		{ flag: 'a' },
		(error) => { })
}