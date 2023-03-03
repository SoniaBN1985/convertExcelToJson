const xlsx = require('xlsx');
var fs = require('fs');

// function convertExcelFileToJsonUsingXlsx() {
//     const file = xlsx.readFile('./Files/data.xlsx');
//     const sheetNames = file.SheetNames;
//     let parsedData = [];
//     xlsx.utils.sheet_to_json()
//     const tempData = xlsx.utils.sheet_to_json(file.Sheets[sheetNames[0]]);
    
//     let objFinal = [];
//     tempData.forEach(obj =>{
//         const objectOne =Object.keys(obj);
//         const objectTwo = Object.values(obj);
//         const key0 = objectOne[0].replace(/['\s]/g, '');
//         const key1 = objectTwo[0].replace(/['\s]/g, '');
//         const value0 = objectOne[1].toString();
//         const value1 = objectTwo[1].toString();
//         objFinal.push({[key0]: value0});
//         objFinal.push({[key1]: value1});
//     });
//     const objConvertedJSON = objFinal.reduce((acc, cur) =>{
//         const key = Object.keys(cur)[0];
//         const value = cur[key];
//         acc[key] = value;
//         return acc;
//     },{});
//     parsedData.push(objConvertedJSON);
//     generateJSONFile(parsedData);
// }

function convertExcelFileToJsonUsingXlsx(fileName) {
    // Read the file Excel
    let file;
    try {
      file = xlsx.readFile(`./Files/${fileName}`);
    } catch (err) {
      console.error(err);
      return;
    }
  
    // Convert the sheet of file
    const sheetNames = file.SheetNames;
    const tempData = xlsx.utils.sheet_to_json(file.Sheets[sheetNames[0]]);
  
    // Create the final object
    const objFinal = {};
    tempData.forEach(obj => {
      const key = obj[Object.keys(obj)[0]].toString().replace(/['\s]/g, '');
      const value = obj[Object.keys(obj)[1]].toString();
      objFinal[key] = value;
    });
  
    // Genera el archivo JSON
    const parsedData = [objFinal];
    generateJSONFile(parsedData);
  }
  

function generateJSONFile(data) {
    try {
    fs.writeFileSync('data.json', JSON.stringify(data))
    } catch (err) {
    console.error(err)
    }
}

convertExcelFileToJsonUsingXlsx('data.xlsx');