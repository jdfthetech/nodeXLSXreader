
// requires node-xlsx installed

let xlsx = require('node-xlsx').default;

const workSheetsFromFile = xlsx.parse('/dir/filename');

// build header
// i will be the sheets so this is designed with 1 sheet files
// would want to specify the sheet manually if working with numerous sheets

let preheader = []

for (let i in workSheetsFromFile){
    for (let x in workSheetsFromFile[i]){
        for (let l in workSheetsFromFile[i][x][0]){
            preheader.push(workSheetsFromFile[i][x][0][l])
            
        }
    }
}
// need to remove 0 index
let header = preheader.slice(1)


// parse data to a row array
let rows = []

for (let i in workSheetsFromFile){
    for (let x in workSheetsFromFile[i]){
        for (let l in workSheetsFromFile[i][x]){
            rows.push(workSheetsFromFile[i][x][l])
            
        }
    }
}


let obj = {}
let finalArr = []

for(let i in rows){
    let keys = header
    let values = rows[i]     
    obj = Object.assign(...keys.map((k, n) => ({[k]: values[n]})))
    finalArr.push(obj)
 }


console.log(finalArr)


