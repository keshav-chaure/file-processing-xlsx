console.log("testing");
const xlsx = require("xlsx");
const fs = require("fs");


const isPropValuesEqual = (subject, target, propNames) =>
  propNames.every(propName => subject[propName] === target[propName]);

const getUniqueItemsByProperties = (items, propNames) => {
    const propNamesArray = Array.from(propNames);
  
    return items.filter((item, index, array) =>
      index === array.findIndex(foundItem => isPropValuesEqual(foundItem, item, propNamesArray))
    );
  };


/* load 'fs' for readFile and writeFile support */
 


const workbook = xlsx.readFile("./files/Sample500.xls");
const sheetNames = workbook.SheetNames;

// Get the data of "Sheet1"
const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetNames[0]])

/// Do what you need with the received data
// data.map(person => {
//  console.log(person)
//   console.log(`${person['Serial Number']} => ${person['Company Name']}`);
// })


// console.log("before removing dup "+data.length)
// let modifiedData=getUniqueItemsByProperties(data, ['Serial Number']);
// console.log("after removing dup "+modifiedData.length)

//var  sorceDataToRemove- is the collection from whrere we get items to remove
 
const sourceWorkbook = xlsx.readFile("./files/source/source_to_remove.xls");
const sourceSheetNames = sourceWorkbook.SheetNames;

// Get the data of "Sheet1"
const sorceData= xlsx.utils.sheet_to_json(sourceWorkbook.Sheets[sourceSheetNames[0]])
let sorceDataToRemove = sorceData.map(a => a['ID']);
console.log("filed array"+sorceDataToRemove);
console.log("data :"+data.length+", sorceDataToRemove :"+sorceDataToRemove.length);

// filter data from 
let modifiedData = data.filter( ( el ) =>{    
    return ((sorceDataToRemove.indexOf( el['Serial Number'] )  < 0))
});
console.log("data :"+data.length);
console.log("modifiedData :"+modifiedData.length);
const ws = xlsx.utils.json_to_sheet(modifiedData)
const wb = xlsx.utils.book_new()
xlsx.utils.book_append_sheet(wb, ws, 'Responses')
xlsx.writeFile(wb, './files/out/sampleData.export.xls');