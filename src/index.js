import xlsx from 'exceljs'
import axios from 'axios'

const workbook = new xlsx.Workbook();

workbook.xlsx.readFile('image-list.xlsx').then(async () => {
    const worksheet = workbook.getWorksheet('Planilha2');
    const imageCol = worksheet.getColumn('G');
    const cells = imageCol.values;

    let statusCodes = [];

    await Promise.all(cells.map(async (cell, i, arr) => {
        if (i < 800) {
            return;
        }
        const link = cell.result;

        const responseStatus = (await axios.get(link).catch((e) => void(0)))?.status;

        if (responseStatus == undefined) statusCodes.push(404);
        else statusCodes.push(responseStatus);
        
    }));

    const lacking = [];

    imageCol.eachCell((cell, row) => {

        if (row < 800) {
            return;
        } else {
    
            if (statusCodes[row - 800] == 200) {
                console.log("Colorindo com a cor certa.") 
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF98FB98' }
                }
            } else {
                lacking.push()
            }

        }
    })

    await workbook.xlsx.writeFile('test.xlsx')
})