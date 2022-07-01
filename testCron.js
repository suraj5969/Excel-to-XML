const cron = require('node-cron');
const fs = require('fs');

function moveLogFile(file, outputDirectory, logsDirectory) {
    const { birthtime } = fs.statSync(`${outputDirectory}/${file}`)

    const nowDate = new Date();
    let Difference_In_Time = nowDate.getTime() - birthtime.getTime();
    let Difference_In_Days = Difference_In_Time / (1000 * 3600 * 24);
    if (Math.trunc(Difference_In_Days) >= 5) {
        fs.renameSync(`${outputDirectory}/${file}`, `${logsDirectory}/${file}`)
    }
}

const outputDirectoryDnD = './BPS Files for XML testing';
const logsDirectoryDnD = './BPS Files for XML testing/Test';
cron.schedule('*/10 * * * * * ', () => {
    console.log('Started Cron');

    fs.readdirSync(outputDirectoryDnD).forEach(fileName => {
        console.log(fileName);
        if (fileName.includes('.xlsx')) {
            moveLogFile(`${fileName}`, outputDirectoryDnD, logsDirectoryDnD);
        }
    });
})