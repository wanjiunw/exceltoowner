const XLSX = require('xlsx');
var moment = require('moment');
const workbook = XLSX.readFile('C:\\Users\\v-wawang\\searchGold\\deploy\\builds\\data\\IndexServe\\Stampdata\\CodeSense\\ReleaseOwner.xlsx');
const sheet_name_list = workbook.SheetNames;
const json_output = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
const nextWednesday = moment().day(3-7).format('YYYY-MM-DD').toString() + 'T12:00:00';
console.log(moment().day(3+7));

var owner;
for (i=0; i<json_output.length; i++)
{
    if (json_output[i].QualificationStartTime === nextWednesday)
    {
        owner = json_output[i].Alias;
        console.log(json_output[i].Alias);
        break;
    }
}

var fs = require('fs')
fs.readFile('C:\\Users\\v-wawang\\Documents\\Jane\\node.js\\excelJS\\email_template.txt', 'utf8', function (err,data) {
  if (err) {
    return console.log(err);
  }
  var result = data.replace(/RELEASE_OWNER/g, owner).replace(/RELEASE_DATE/g, nextWednesday);

  fs.writeFile('C:\\Users\\v-wawang\\Documents\\Jane\\node.js\\excelJS\\email_output.txt', result, 'utf8', function (err) {
     if (err) return console.log(err);
  });
});