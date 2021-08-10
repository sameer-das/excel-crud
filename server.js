const express = require('express');
const reader = require('xlsx');
const Excel = require('exceljs');

const app = express();

app.use(express.json());

app.get('/', (req,res) => {
    res.send('test route');
})

app.get('/read', async (req,res) => {
    //res.send('App works ok');
    const workbook = reader.readFile('data.xlsx');
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const firstSheetData = reader.utils.sheet_to_json(worksheet);

    
    res.status(200).json(firstSheetData);

})


app.post('/append', async (req,res) => {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile('./data.xlsx');
    const worksheet = workbook.getWorksheet('Sheet1');
    console.log(req.body)
    worksheet.insertRow(++(worksheet.lastRow.number), Object.values(req.body)).commit();
    workbook.xlsx.writeFile('./data.xlsx');

    res.status(200).json(req.body);
})

app.listen('3000', () => {
    console.log(`app listening`);
})