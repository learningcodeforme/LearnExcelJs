
// to use excel.js we need to import it
import ExcelJS from 'exceljs'
//define a function 
async function excelTest() {
    /// create  workbook object 
    // define object called output to handle the row and colum number
    const output = { row: -1, column: -1 }
    const workbook = new ExcelJS.Workbook()
    //read the excel using readFile ( pass the path of excel)
    await workbook.xlsx.readFile('download.xlsx')
    // .getWorksheet('Sheetname')
    const worksheet = workbook.getWorksheet('Sheet1')
    // iterate to each row
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, cellNumber) => {
            // console.log(cell.value)
            // if (cell.value === 'Apple') {
            if (cell.value === 'Banana') {
                console.log(rowNumber)
                console.log(cellNumber)
                //assign row and column of Apple in output.row and output.column
                output.row = rowNumber;
                output.column = cellNumber;

            }

        })
    })
    // we get cell coodinates
    //const cell = worksheet.getCell(3, 2)
    // uisng row and coulumn of banana to replace with Rolex
    const cell = worksheet.getCell(output.row, output.column)
    //update the cell value
    cell.value = 'ROLEX-TAM'
    await workbook.xlsx.writeFile('download.xlsx')



}


excelTest()