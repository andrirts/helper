const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const moment = require('moment');
const archiver = require('archiver');

const app = express();
const upload = multer({ dest: 'uploads/' });

function createExcelFile(columns) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');

    worksheet.columns = columns;
    worksheet.getRow(1).eachCell((cell) => {
        cell.font = { bold: true };
    });
    return [workbook, worksheet];
}

// Upload endpoint
app.post('/upload', upload.single('file'), async (req, res) => {
    const file = req.file;
    console.log(file.originalname.split('.')[0].replace(/\s+/g, '-'));
    // Create a directory to store the Excel files
    const dir = path.join(__dirname, 'zipped');
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir);
    }
    const columns = [
        { header: 'Transaction Date', key: 'Transaction Date', width: 25 },
        { header: 'Transaction Time', key: 'Transaction Time', width: 25 },
        { header: 'Reff ID', key: 'Reff ID', width: 50 },
        { header: 'Partner Reff', key: 'Partner Reff', width: 25 },
        { header: 'Product Name', key: 'Product Name', width: 50 },
        { header: 'Billing Number', key: 'Billing Number', width: 25 },
        { header: 'Biller Product Code', key: 'Biller Product Code', width: 25 },
        { header: 'Sell Price', key: 'Sell Price', width: 25 },
        { header: 'Status', key: 'Status', width: 25 },
        { header: 'Serial Number', key: 'Serial Number', width: 50 },
    ];

    if (!file) {
        return res.status(400).send('No file uploaded.');
    }
    console.log('starting...')
    // Read the uploaded Excel file
    const workbook = XLSX.readFile(file.path);
    // Get the first sheet
    const ctiRiu = workbook.SheetNames.find(sheetName => sheetName === 'CTIRIU');
    const worksheetCtiriu = workbook.Sheets[ctiRiu];
    // Get the second sheet
    const via = workbook.SheetNames.find(sheetName => sheetName === 'VIA');
    const worksheetVia = workbook.Sheets[via];
    // Get the third sheet
    const core = workbook.SheetNames.find(sheetName => sheetName === 'CORE');
    const worksheetCore = workbook.Sheets[core];

    // const alto = workbook.SheetNames.find(sheetName => sheetName === 'Alto');
    // const worksheetAlto = workbook.Sheets[alto];

    // const tokopedia = workbook.SheetNames.find(sheetName => sheetName === 'Tokped');
    // const worksheetTokopedia = workbook.Sheets[tokopedia];

    // Convert sheet to JSON
    const jsonDataCtiriu = XLSX.utils.sheet_to_json(worksheetCtiriu);
    const jsonDataVia = XLSX.utils.sheet_to_json(worksheetVia);
    const jsonDataCore = XLSX.utils.sheet_to_json(worksheetCore);
    // const jsonDataAlto = XLSX.utils.sheet_to_json(worksheetAlto);
    // const jsonDataTokopedia = XLSX.utils.sheet_to_json(worksheetTokopedia);

    const [matchedVIAWorkbook, matchedVIAWorksheet] = createExcelFile(columns);
    const [unmatchedVIAWorkbook, unmatchedVIAWorksheet] = createExcelFile(columns);
    const [unmatchedCtiriuWorkbook, unmatchedCtiriuWorksheet] = createExcelFile(columns);
    const [unmatchedWorkbook, unmatchedWorksheet] = createExcelFile(columns);
    // const [matchedAltoWorkbook, matchedAltoWorksheet] = createExcelFile(columns);
    // const [unmatchedAltoWorkbook, unmatchedAltoWorksheet] = createExcelFile(columns);
    // const [matchedTokopediaWorkbook, matchedTokopediaWorksheet] = createExcelFile(columns);
    // const [unmatchedTokopediaWorkbook, unmatchedTokopediaWorksheet] = createExcelFile(columns);

    for (let i = 0; i < jsonDataCore.length; i++) {
        console.log(i);
        const isExistsOnCtiRiu = jsonDataCtiriu.find(item => item['TRANSACTION_ID'] === jsonDataCore[i]['Reff ID']);
        // const isExistsOnAlto = jsonDataAlto.find(item => item['Trx Reff ID'] === jsonDataCore[i]['Reff ID']);
        // const isExistsOnTokopedia = jsonDataTokopedia.find(item => item['Trx Reff ID'] === jsonDataCore[i]['Reff ID']);
        const isExistsOnVia = jsonDataVia.find(item => item['Partner Reff'] === jsonDataCore[i]['Partner Reff']);

        const inputtedData = {
            'Transaction Date': jsonDataCore[i]['Transaction Date'],
            'Transaction Time': jsonDataCore[i]['Transaction Time'],
            'Reff ID': jsonDataCore[i]['Reff ID'],
            'Partner Reff': jsonDataCore[i]['Partner Reff'],
            'Product Name': jsonDataCore[i]['Product Name'],
            'Billing Number': jsonDataCore[i]['Billing Number'],
            'Biller Product Code': jsonDataCore[i]['Biller Product Code'],
            'Sell Price': jsonDataCore[i]['Sell Price'],
            'Status': jsonDataCore[i]['Status'],
            // 'Serial Number': serialNumber,
        };

        // let statusRiu = '';
        // let statusVia = '';

        // if (isExistsOnCtiRiu) {
        //     statusRiu = 'Match';
        // } else {
        //     if (jsonDataCore[i]['Status'] === 'SUCCESS') {
        //         statusRiu = 'Match';
        //     } else {
        //         statusRiu = 'Unmatch CTI RIU';
        //     }
        // }

        // if (isExistsOnVia) {
        //     statusVia = isExistsOnVia['Status'] === jsonDataCore[i]['Status'] ? 'Match' : 'Unmatch VIA';
        // } else {
        //     statusVia = 'Unmatch VIA';
        // }

        let serialNumber = jsonDataCore[i]['Serial Number'];
        if (!serialNumber) {
            serialNumber = isExistsOnCtiRiu ? isExistsOnCtiRiu['SERIAL_NUMBER'] : '';
        }

        // if (statusRiu === 'Match' && statusVia === 'Match') {
        //     matchedVIAWorksheet.addRow({
        //         ...inputtedData,
        //         'Serial Number': serialNumber,
        //     });
        // }
        // if (statusRiu === 'Match' && statusVia === 'Unmatch VIA') {
        //     unmatchedVIAWorksheet.addRow({
        //         ...inputtedData,
        //         'Serial Number': serialNumber,
        //     })
        // }
        // if (statusRiu === 'Unmatch CTI RIU') {
        //     unmatchedCtiriuWorksheet.addRow({
        //         ...inputtedData,
        //         'Serial Number': serialNumber,
        //     })
        // }

        if (jsonDataCore[i]['Status'] === 'SUCCESS') {
            if (isExistsOnCtiRiu) {
                if (jsonDataCore[i]['Partner App Name'] === 'VIA' || jsonDataCore[i]['Partner App Name'] === 'VIA_SOLUTAIRE') {
                    if (isExistsOnVia) {
                        await matchedVIAWorksheet.addRow({
                            ...inputtedData,
                            'Serial Number': serialNumber,
                        });
                    } else {
                        await unmatchedVIAWorksheet.addRow({
                            ...inputtedData,
                            'Serial Number': serialNumber,
                        })
                    }
                }
            } else {
                if (jsonDataCore[i]['Partner App Name'] === 'VIA' || jsonDataCore[i]['Partner App Name'] === 'VIA_SOLUTAIRE') {
                    if (isExistsOnVia) {
                        await unmatchedCtiriuWorksheet.addRow({
                            ...inputtedData,
                            'Serial Number': serialNumber,
                        })
                        await matchedVIAWorksheet.addRow({
                            ...inputtedData,
                            'Serial Number': serialNumber,
                        })
                    } else {
                        await unmatchedWorksheet.addRow({
                            ...inputtedData,
                            'Serial Number': serialNumber,
                        })
                    }
                }
            }
        } else {
            if (isExistsOnCtiRiu) {
                if (jsonDataCore[i]['Partner App Name'] === 'VIA' || jsonDataCore[i]['Partner App Name'] === 'VIA_SOLUTAIRE') {
                    if (isExistsOnVia) {
                        await matchedVIAWorksheet.addRow({
                            ...inputtedData,
                            'Serial Number': serialNumber,
                        })
                    } else {
                        await unmatchedVIAWorksheet.addRow({
                            ...inputtedData,
                            'Serial Number': serialNumber,
                        })
                    }
                }
            } else {
                if (jsonDataCore[i]['Partner App Name'] === 'VIA' || jsonDataCore[i]['Partner App Name'] === 'VIA_SOLUTAIRE') {
                    if (isExistsOnVia) {
                        await unmatchedVIAWorksheet.addRow({
                            ...inputtedData,
                            'Serial Number': serialNumber,
                        })
                    } else {
                        await unmatchedWorksheet.addRow({
                            ...inputtedData,
                            'Serial Number': serialNumber,
                        })
                    }
                }
            }
        }
    }

    const unmatchedExcelFilePath = path.join(__dirname, 'zipped', moment().format('YYYY-MM-DD') + ' Unmatched' + '.xlsx');
    await unmatchedWorkbook.xlsx.writeFile(unmatchedExcelFilePath);

    const matchedVIAExcelFilePath = path.join(__dirname, 'zipped', moment().format('YYYY-MM-DD') + ' Matched VIA' + '.xlsx');
    await matchedVIAWorkbook.xlsx.writeFile(matchedVIAExcelFilePath);

    const unmatchedVIAExcelFilePath = path.join(__dirname, 'zipped', moment().format('YYYY-MM-DD') + ' Unmatched VIA' + '.xlsx');
    await unmatchedVIAWorkbook.xlsx.writeFile(unmatchedVIAExcelFilePath);

    const unmatchedCtiriuExcelFilePath = path.join(__dirname, 'zipped', moment().format('YYYY-MM-DD') + ' Unmatched CTI RIU' + '.xlsx');
    await unmatchedCtiriuWorkbook.xlsx.writeFile(unmatchedCtiriuExcelFilePath);

    // const matchedAltoExcelFilePath = path.join(__dirname, 'zipped', moment().format('YYYY-MM-DD') + ' Matched ALTO' + '.xlsx');
    // await matchedAltoWorkbook.xlsx.writeFile(matchedAltoExcelFilePath);

    // const unmatchedAltoExcelFilePath = path.join(__dirname, 'zipped', moment().format('YYYY-MM-DD') + ' Unmatched ALTO' + '.xlsx');
    // await unmatchedAltoWorkbook.xlsx.writeFile(unmatchedAltoExcelFilePath);

    // const matchedTokopediaExcelFilePath = path.join(__dirname, 'zipped', moment().format('YYYY-MM-DD') + ' Matched TOKOPEDIA' + '.xlsx');
    // await matchedTokopediaWorkbook.xlsx.writeFile(matchedTokopediaExcelFilePath);

    // const unmatchedTokopediaExcelFilePath = path.join(__dirname, 'zipped', moment().format('YYYY-MM-DD') + ' Unmatched TOKOPEDIA' + '.xlsx');
    // await unmatchedTokopediaWorkbook.xlsx.writeFile(unmatchedTokopediaExcelFilePath);

    console.log('Excel file successfully written');
    const folderPath = path.join(__dirname, 'zipped');
    const zilFileName = `Reconciliation-${file.originalname.split('.')[0].replace(/\s+/g, '-')}.zip`;
    console.log(zilFileName);
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename=' + zilFileName);

    const archive = archiver('zip', {
        zlib: { level: 9 } // Sets the compression level.
    });

    archive.pipe(res);

    archive.directory(folderPath, false);

    archive.finalize();

    res.on('finish', async () => {
        try {
            await fs.promises.rm(folderPath, { recursive: true, force: true });
            const uploadsDir = path.join(__dirname, 'uploads');
            await fs.promises.rm(uploadsDir, { recursive: true, force: true });
            console.log(`Folder ${folderPath} successfully deleted`);
            console.log(`Folder ${uploadsDir} successfully deleted`);

            //create again the folder
            // fs.mkdirSync(folderPath);
            fs.mkdirSync(uploadsDir);
        } catch (err) {
            console.error(`Error deleting folder ${folderPath}:`, err);
        }
    });

});


app.listen(3000, () => {
    console.log('Server is running on port 3000');
});
