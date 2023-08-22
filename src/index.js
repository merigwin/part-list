import xlsx from 'xlsx'
import axios from 'axios'
import path from 'path'
import fs from 'fs'

const filePath = path.resolve("image-list.xlsx");
(async () => {
    const wb = xlsx.readFile(filePath);
    const ws = wb.Sheets['Planilha2'];

    const rows = xlsx.utils.sheet_to_json(ws);

    let newRows = [];

    for (let i = 0; i < rows.length; i++) {
        newRows[i] = rows[i];

        const row = rows[i];
        const link = row.Coluna2;

        try {
            const status = (await axios.get(link)).status;
            newRows[i].Existe = "SIM"
        } catch (err) {
            const sku = row.Sku;
            const imagePath = path.resolve("..", "images", `${sku}.png`);
            const newImagePath = path.resolve("..", "lacking", `${sku}.png`);

            if (fs.existsSync(imagePath)) {
                console.log("\x1B[32mImagem existe e está no diretório de imagens => " + `${imagePath}\x1b[0m`);
                fs.copyFileSync(imagePath, newImagePath);
                newRows[i].Existe = "SIM"
            }
            else {
                console.log("\x1b[31mA imagem faltando não está no diretório de imagens => " + imagePath + "\x1b[0m");
                newRows[i].Existe = "NÃO"
            }
        }
    }
    const newSheet = xlsx.utils.json_to_sheet(newRows);
    const workBoork = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workBoork, newSheet);

    xlsx.writeFile(workBoork, "updated-list.xlsx");
})()