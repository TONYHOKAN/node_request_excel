// go http://nodejs.org/dist/v10.15.1/ to download portable nodejs

const request = require('request');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

callUrl()

async function callUrl () {
    // test html
    // const url = 'https://www.google.com/'
    // test json
    // const url = 'https://demo.ckan.org/api/3/action/group_list'
    // test download xlsx
    const url = 'http://oss.sheetjs.com/js-xlsx/test_files/formula_stress_test_ajax.xlsx'
    const option =
    {
      method: 'GET',
      uri: url,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36',
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      }, // mock browser to call else not work
      encoding: null
    }

    try {
        await request(option, function (error, response, body) {
            const excelName = "some_excel.xlsx"
            console.log('[LOG] start request')
            if (!error && response.statusCode == 200) {
                // console.log(`[LOG] response body: ${body}`)
                const saveFilePath = __dirname + '/' + excelName
                console.log(`[LOG] save file to path: ${saveFilePath}`)
                fs.writeFileSync(saveFilePath, body, 'binary')

                // ref: https://stackoverflow.com/a/43614922/5824101
                var arraybuffer = body;
                /* convert data to binary string */
                var data = arraybuffer;            
                //var data = new Uint8Array(arraybuffer);                
                var arr = new Array();
                for (var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
                var bstr = arr.join("");


                /* Call XLSX */
                var sheetName = 'Database';
                var workbook = XLSX.read(bstr, { type: "binary" });
                var worksheet = workbook.Sheets[sheetName];

                // XLSX ref: https://github.com/SheetJS/js-xlsx
                // tutorial https://aotu.io/notes/2016/04/07/node-excel/index.html

                let a1 = worksheet['A1'];
                console.log("A1 cell value: " + a1.v);

                var csv = XLSX.utils.sheet_to_csv(worksheet);
                console.log("csv" + csv);
            }
            else
            {
                console.log(`[LOG] failed, url: ${url} return status code: ${response.statusCode}, body: ${body}`)
            }
        })
    } catch (error) {
        console.log(`[LOG] promise reject: ${error}`)
    }
   
}