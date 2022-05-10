const { getJsDateFromExcel } = require("excel-date-to-js")
const { format, isValid } = require("date-fns")

try {
    let a = getJsDateFromExcel('04-05-2021  1:22:00 PM')
    a = new Date(a.valueOf() + a.getTimezoneOffset() * 60 * 1000);
    a = format(a, "yyyy-MM-dd'T'hh:mm:ss.SSS")
    console.log(a)
}
catch (e) {
    // console.log('got errror', e)
    let a = new Date('04-05-2021 1:22:00 PM');
    if(isValid(a)) {
        // a = new Date(a.valueOf() + a.getTimezoneOffset() * 60 * 1000);
        a = format(a, "yyyy-MM-dd'T'hh:mm:ss.SSS")
    }
    console.log(a)
}