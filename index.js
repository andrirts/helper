const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const moment = require('moment');
const archiver = require('archiver');
const { queryDatabase } = require('./database');
const createTemplate = require('./template-struk-listrik');

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

function excelDateToJsDate(serial) {
    // Excel date serial number is the number of days since January 1, 1900
    const excelEpoch = new Date(Date.UTC(1900, 0, 1));
    // Adjust for Excel leap year bug
    const excelDate = new Date(excelEpoch.getTime() + (serial - 1) * 24 * 60 * 60 * 1000);

    // Extract fractional part to calculate time
    const timePortion = serial % 1;
    const hours = Math.floor(timePortion * 24);
    const minutes = Math.floor((timePortion * 24 - hours) * 60);
    const seconds = Math.floor(((timePortion * 24 - hours) * 60 - minutes) * 60);

    // Add time to the date
    excelDate.setUTCHours(hours, minutes, seconds);

    const offset = 7 * 60;
    excelDate.setUTCMinutes(excelDate.getUTCMinutes() + offset);

    return excelDate;
}

app.get("/", (req, res, next) => {
    return res.send("Hello World");
})

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
            // await fs.promises.rm(uploadsDir, { recursive: true, force: true });
            fs.unlinkSync(file.path);
            console.log(`Folder ${folderPath} successfully deleted`);
            console.log(`Folder ${uploadsDir} successfully deleted`);

            //create again the folder
            // fs.mkdirSync(folderPath);
            // fs.mkdirSync(uploadsDir);
        } catch (err) {
            console.error(`Error deleting folder ${folderPath}:`, err);
        }
    });

});

app.post('/upload-trx', upload.single('file'), async (req, res) => {
    const file = req.file;
    const workbook = XLSX.readFile(file.path);
    const worksheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[worksheetName];
    const datas = XLSX.utils.sheet_to_json(worksheet);
    console.log("Processing");
    try {
        await queryDatabase('BEGIN');

        for (const data of datas) {
            if (!data['Reference ID']) {
                await queryDatabase('ROLLBACK');
                return res.json({
                    message: 'Reference ID not found'
                })
            }
            const findData = await queryDatabase('select * from ppob_transaction where reference_id = $1', [data['Reference ID']]);
            if (findData.length === 0) {
                await queryDatabase('ROLLBACK');
                return res.json({
                    message: `Trx Id ${data['Reference ID']} not found`
                })
            }

            const updateQuery = 'update ppob_transaction set status = $1, response_data = $2, response_biller = $3 where reference_id = $4';
            const status = data['Status'] === 'SUCCESS' ? 'SUCCESS' : 'FAILED';
            const response_biller = status === 'SUCCESS' ? 'Approve' : findData[0].response_biller;
            let response_data = findData[0].response_data ? JSON.parse(findData[0].response_data) : null;
            if (status === 'SUCCESS') {
                if (response_data) {
                    response_data['serial_number'] = data['Serial Number'];
                } else {
                    response_data = {
                        customer_number: '',
                        customer_name: '',
                        serial_number: data['Serial Number']
                    }
                }
            }
            const params = [status, response_data, response_biller, data['Reference ID']];
            await queryDatabase(updateQuery, params);
        }
        await queryDatabase('COMMIT');
        console.log("Deleting the file");
        fs.unlinkSync(file.path);
        return res.json({
            message: 'Update trx successfully'
        })
    } catch (err) {
        await queryDatabase('ROLLBACK');
        console.log(err)
        return res.json({
            message: 'Internal server error'
        })
    }
})

app.post('/upload-struk-pln', upload.single('file'), async (req, res) => {
    const file = req.file;
    if (!file) {
        return res.status(400).send('No file uploaded');
    }
    const workbook = new ExcelJS.Workbook();
    const datas = XLSX.readFile(file.path);
    const dataWorksheet = datas.SheetNames[0];
    const jsonData = XLSX.utils.sheet_to_json(datas.Sheets[dataWorksheet]);
    const sheets = {};
    console.log(jsonData);
    for (let i = 0; i < jsonData.length; i++) {
        sheets[`Sheet ${i + 1}`] = await createTemplate(workbook, i + 1);
        const cellC6 = sheets[`Sheet ${i + 1}`].getCell('C6');
        cellC6.value = `${jsonData[i]['Customer ID']}`;

        const cellC7 = sheets[`Sheet ${i + 1}`].getCell('C7');
        cellC7.value = jsonData[i]['Customer Name'];

        const cellC8 = sheets[`Sheet ${i + 1}`].getCell('C8');
        cellC8.value = jsonData[i]['Tarif/Daya'];

        const cellC9 = sheets[`Sheet ${i + 1}`].getCell('C9');
        cellC9.value = `${jsonData[i]['Periode']}`

        const cellC10 = sheets[`Sheet ${i + 1}`].getCell('C10');
        cellC10.value = `${jsonData[i]['Stan Meter']}`;

        const cellC11 = sheets[`Sheet ${i + 1}`].getCell('C11');
        cellC11.value = `${jsonData[i]['SN']}`;
        cellC11.alignment = {
            wrapText: true
        }

        const cellC12 = sheets[`Sheet ${i + 1}`].getCell('C12');
        cellC12.value = `Rp ${jsonData[i]['Base Bill'].toLocaleString('id-ID')}`;

        const cellC13 = sheets[`Sheet ${i + 1}`].getCell('C13');
        cellC13.value = `Rp ${jsonData[i]['Admin Fee'].toLocaleString('id-ID')}`;

        const cellC14 = sheets[`Sheet ${i + 1}`].getCell('C14');
        cellC14.value = `Rp ${jsonData[i]['Price'].toLocaleString('id-ID')}`;

        const cellG6 = sheets[`Sheet ${i + 1}`].getCell('G6');
        cellG6.value = `${jsonData[i]['Customer ID']}`;

        const cellG7 = sheets[`Sheet ${i + 1}`].getCell('G7');
        cellG7.value = jsonData[i]['Customer Name'];

        const cellG8 = sheets[`Sheet ${i + 1}`].getCell('G8');
        cellG8.value = jsonData[i]['Tarif/Daya'];

        const cellG9 = sheets[`Sheet ${i + 1}`].getCell('G9');
        cellG9.value = `Rp ${jsonData[i]['Base Bill'].toLocaleString('id-ID')}`;

        const cellG10 = sheets[`Sheet ${i + 1}`].getCell('G10');
        cellG10.value = `${jsonData[i]['SN']}`;

        const cellG12 = sheets[`Sheet ${i + 1}`].getCell('G12');
        cellG12.value = `Rp ${jsonData[i]['Admin Fee'].toLocaleString('id-ID')}`;

        const cellG13 = sheets[`Sheet ${i + 1}`].getCell('G13');
        cellG13.value = `Rp ${jsonData[i]['Price'].toLocaleString('id-ID')}`;

        const cellJ1 = sheets[`Sheet ${i + 1}`].getCell('J1');
        cellJ1.value = `${moment(excelDateToJsDate(jsonData[i]['Created Date'])).format('YYYY-MM-DD HH:mm:ss')}`;

        const cellK6 = sheets[`Sheet ${i + 1}`].getCell('K6');
        cellK6.value = `${jsonData[i]['Periode']}`;

        const cellK7 = sheets[`Sheet ${i + 1}`].getCell('K7');
        cellK7.value = `${jsonData[i]['Stan Meter']}`;

        const listCells = [cellC6, cellC7, cellC8, cellC9, cellC10, cellC11, cellC12, cellC13, cellC14, cellG6, cellG7, cellG8, cellG9, cellG10, cellG12, cellG13, cellJ1, cellK6, cellK7];
        const defaultFont = {
            name: 'Arial',
            size: 8,
            bold: false,       // Optional: Set bold font
            italic: false     // Optional: Set italic font
        }
        listCells.forEach((cell) => {
            cell.font = defaultFont;
        });

        sheets[`Sheet ${i + 1}`].pageSetup.margins = {
            left: 0.25,
            right: 0.25,
            top: 0.25,
            bottom: 0.25,
            header: 0.3,
            footer: 0.3
        }
    }
    // Save the workbook to a file
    const filePath = `${file.originalname}`;
    await workbook.xlsx.writeFile(filePath)
        .then(() => {
            console.log(`Workbook saved to ${filePath}`);
        })
        .catch((error) => {
            console.error('Error writing workbook:', error);
        });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=' + `${file.originalname}`);
    res.download(filePath, `${file.originalname}`, (err) => {
        if (err) {
            console.log(err);
        }
        fs.unlinkSync(filePath);
        fs.unlinkSync(file.path);
        // const uploadsDir = path.join(__dirname, 'uploads');
        // fs.rm(uploadsDir, { recursive: true }, (err) => {
        //     if (err) {
        //         console.error('Error deleting uploads folder:', err);
        //     } else {
        //         console.log('Uploads folder deleted successfully');
        //         // Recreate the uploads folder
        //         fs.mkdir(uploadsDir, (err) => {
        //             if (err) {
        //                 console.error('Error creating uploads folder:', err);
        //             } else {
        //                 console.log('Uploads folder recreated successfully');
        //             }
        //         });
        //     }
        // });
    });
});


app.listen(3000, () => {
    console.log('Server is running on port 3000');
});
