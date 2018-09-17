const XLSX = require('xlsx');
const fs = require('fs');
const program = require('commander');
let fileName = "tasks-in-progress-data-dump-dnb.xlsx";
const buf = fs.readFileSync(fileName);
const wb = XLSX.read(buf, {type:'buffer'});
/*
// if input is taken from terminal
program
    .version('0.1.0')
    .option('-f, --file <file>', 'use specified workbook')
    .option('-u, --user <usr>', 'cmp user name')
    .option('-p, --password <pw>', 'user password to login to cmp')
    .parse(process.argv);

let filename = program.file;
let username =  program.user;
let password = program.password;

if(!filename) {
  console.error("Error: must specify a filename after -f flag");
  process.exit(1);
}

if(!fs.existsSync(filename)) {
	console.error(`${filename} : No such file or directory`);
	process.exit(2);
}
*/

var tasks = wb.Sheets.tasks;
var tasksJson = XLSX.utils.sheet_to_json(tasks);

var inProgressTasks = tasksJson.filter( task => {
  return task['Task status'] === 'In Progress'
});

const newWs = XLSX.utils.json_to_sheet(inProgressTasks, {header:["A","B","C","D","E","F"], skipHeader:true});
XLSX.utils.book_append_sheet(wb, newWs, "In progress");
XLSX.writeFile(wb, fileName)