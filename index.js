const Excel = require('exceljs');

const ExportExcel = async () => {
    let month = [
        "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"
    ];
    
    month = month.reverse();

    const workbook = new Excel.Workbook();

    let plants = ['R1', 'R2', 'R3', 'R4'];

    for (let sheetId = 0; sheetId < plants.length; sheetId++) {

        const newSheet = workbook.addWorksheet(plants[sheetId]);
        // set width C column
        newSheet.getColumn(3).width = 40;

        // set width A column
        newSheet.getColumn(1).width = 4;

        newSheet.getCell('B2').value = 'CPO Refinery 1 Production & Cost Summary';
        newSheet.getCell('G2').value = 'Plant:';

        newSheet.getCell('B4').value = 'No';
        newSheet.mergeCells('B4:B5');

        newSheet.getCell('C4').value = 'Description';
        newSheet.mergeCells('C4:C5');

        [
            'B4',
            'C4',
        ].forEach(cell => {
            newSheet.getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FEBF00' },
                bgColor: { argb: 'FEBF00' },
            };
            newSheet.getCell(cell).alignment = {
                vertical: 'middle',
                horizontal: 'center'
            };
            newSheet.getCell(cell).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        let D = 4;
        let E = 5;
        let F = 6;
        let G = 7;

        for (let index = 0; index < 12; index++) {
            newSheet.getCell(4, D).value = month[index] + '-22';
            newSheet.getCell(5, D).value = 'Actual';
            newSheet.getCell(5, E).value = 'Budget';
            newSheet.getCell(5, F).value = 'Variance';
            newSheet.getCell(5, G).value = 'Normalize';

            newSheet.mergeCells(4, D, 4, G);

            newSheet.getCell(4, D).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FEBF00' },
                bgColor: { argb: 'FEBF00' },
            };

            [
                { r: 5, c: D },
                { r: 5, c: E },
                { r: 5, c: F },
                { r: 5, c: G },
            ].forEach(cell => {
                newSheet.getCell(cell.r, cell.c).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FEE598' },
                    bgColor: { argb: 'FEE598' },
                };
            });

            [
                { r: 4, c: D },
                { r: 5, c: D },
                { r: 5, c: E },
                { r: 5, c: F },
                { r: 5, c: G },
            ].forEach(cell => {
                newSheet.getCell(cell.r, cell.c).alignment = {
                    vertical: 'middle',
                    horizontal: 'center'
                };
                newSheet.getCell(cell.r, cell.c).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            newSheet.addRow().commit();
            D += 4; E += 4; F += 4; G += 4;
        }
    }

    // write to a file
    await workbook.xlsx.writeFile('exported.xlsx');
}

// const charToCode = (str) => str.charCodeAt();

// const codeToChar = (code) => String.fromCharCode(code)

ExportExcel();