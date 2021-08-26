
// requires node-xlsx installed

var xlsx = require('node-xlsx').default;

let preheader = []

const workSheetsFromFile = xlsx.parse('/dir/filename');

// build header
for (var i in workSheetsFromFile){
    for (var x in workSheetsFromFile[i]){
        for (var l in workSheetsFromFile[i][x][0]){
            preheader.push(workSheetsFromFile[i][x][0][l])
            
        }
    }
}
// need to remove 0 index
let header = preheader.slice(1)


// parse data to a row array
let rows = []

for (var i in workSheetsFromFile){
    for (var x in workSheetsFromFile[i]){
        for (var l in workSheetsFromFile[i][x]){
            rows.push(workSheetsFromFile[i][x][l])
            
        }
    }
}


let obj = {}
let finalArr = []


   for(var i = 0; i < rows.length; i++){
        let keys = header
        let values = rows[i]     
        obj = Object.assign(...keys.map((k, n) => ({[k]: values[n]})))
        finalArr.push(obj)
     }
    

console.log(finalArr)


