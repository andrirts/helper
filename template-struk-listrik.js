const ExcelJS = require('exceljs');

function applyDefaultStyle(sheet) {
    const defaultFont = {
        name: 'Arial',
        size: 8,
        bold: false,       // Optional: Set bold font
        italic: false     // Optional: Set italic font
    }

    sheet.eachRow({ includeEmpty: true }, (row) => {
        row.eachCell({ includeEmpty: true }, (cell) => {
            cell.font = defaultFont;
        });
    });
}

// Create a new workbook and add a worksheet
function createTemplate(workbook, sheetPage) {
    const worksheet = workbook.addWorksheet(`Sheet ${sheetPage}`);

    worksheet.getColumn('A').width = 12;
    worksheet.getColumn('B').width = 1.91;
    worksheet.getColumn('C').width = 18.09;
    worksheet.getColumn('D').width = 5.09;
    worksheet.getColumn('E').width = 11.64;
    worksheet.getColumn('F').width = 1.91;
    worksheet.getColumn('G').width = 16.82;
    worksheet.getColumn('H').width = 4.19;
    worksheet.getColumn('I').width = 12.09;
    worksheet.getColumn('J').width = 1.91;
    worksheet.getColumn('K').width = 15.64;
    worksheet.getColumn('L').width = 18.73;

    //Merging Cells
    worksheet.mergeCells('A4:C4');
    worksheet.mergeCells('E4:K4');
    worksheet.mergeCells('E11:K11');
    worksheet.mergeCells('E15:K15');
    worksheet.mergeCells('E16:K16');

    const fontTitle = {
        name: 'Arial',
        size: 9,
        bold: true,       // Optional: Set bold font
        italic: false     // Optional: Set italic font
    }

    // Cell A
    // Write data to cell A1
    const cellA1 = worksheet.getCell('A1');
    cellA1.value = 'BANK MANDIRI';

    const cellA2 = worksheet.getCell('A2');
    cellA2.value = 'RTS 01 (54RTSOP1 )';

    const cellA4 = worksheet.getCell('A4');
    cellA4.value = 'STRUK TAGIHAN LISTRIK';

    const cellA6 = worksheet.getCell('A6');
    cellA6.value = 'IDPEL';

    const cellA7 = worksheet.getCell('A7');
    cellA7.value = 'NAMA';

    const cellA8 = worksheet.getCell('A8');
    cellA8.value = 'TARIF/DAYA';

    const cellA9 = worksheet.getCell('A9');
    cellA9.value = 'BL/TH';

    const cellA10 = worksheet.getCell('A10');
    cellA10.value = 'STAND METER';

    const cellA11 = worksheet.getCell('A11');
    cellA11.value = 'NO REF';

    const cellA12 = worksheet.getCell('A12');
    cellA12.value = 'RP TAG PLN';

    const cellA13 = worksheet.getCell('A13');
    cellA13.value = 'ADMIN BANK';

    const cellA14 = worksheet.getCell('A14');
    cellA14.value = 'TOTAL  BAYAR';

    for (let row = 6; row <= 14; row++) {
        worksheet.getCell(`B${row}`).value = ':';
        worksheet.getCell(`B${row}`).alignment = {
            vertical: 'bottom',
            horizontal: 'center'
        };
    }

    //Cell E

    const cellE1 = worksheet.getCell('E1');
    cellE1.value = 'BANK MANDIRI';

    const cellE2 = worksheet.getCell('E2');
    cellE2.value = 'RTS 01 (54RTSOP1 )';

    const cellE4 = worksheet.getCell('E4');
    cellE4.value = 'STRUK TAGIHAN LISTRIK';

    const cellE6 = worksheet.getCell('E6');
    cellE6.value = 'IDPEL';

    const cellE7 = worksheet.getCell('E7');
    cellE7.value = 'NAMA';

    const cellE8 = worksheet.getCell('E8');
    cellE8.value = 'TARIF/DAYA';

    const cellE9 = worksheet.getCell('E9');
    cellE9.value = 'RP TAG PLN';

    const cellE10 = worksheet.getCell('E10');
    cellE10.value = 'NO REF';

    const cellE11 = worksheet.getCell('E11');
    cellE11.value = 'PLN menyatakan struk ini sebagai bukti pembayaran yang sah';

    const cellE12 = worksheet.getCell('E12');
    cellE12.value = 'ADMIN BANK';

    const cellE13 = worksheet.getCell('E13');
    cellE13.value = 'TOTAL  BAYAR';

    const cellE15 = worksheet.getCell('E15');
    cellE15.value = 'TERIMA KASIH';

    const cellE16 = worksheet.getCell('E16');
    cellE16.value = `"Informasi Hubungi Call Center 123 Atau Hub PLN Terdekat"`;


    for (let row = 6; row <= 13; row++) {
        if (row == 11) {
            continue;
        }
        worksheet.getCell(`F${row}`).value = ':';
        worksheet.getCell(`F${row}`).alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
    }

    //Cell I
    const cellI1 = worksheet.getCell('I1');
    cellI1.value = 'TGL BAYAR :';

    const cellI6 = worksheet.getCell('I6');
    cellI6.value = 'BL/TH';

    const cellI7 = worksheet.getCell('I7');
    cellI7.value = 'STAND METER';

    //Cell J

    // const cellJ1 = worksheet.getCell('J1');
    // cellJ1.value = `2024-07-15 16:08:21`;

    for (let row = 6; row <= 7; row++) {
        worksheet.getCell(`J${row}`).value = ':';
        worksheet.getCell(`J${row}`).alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
    }

    applyDefaultStyle(worksheet);
    cellA1.font = fontTitle;
    cellA2.font = fontTitle;
    cellA4.font = fontTitle;
    cellA4.alignment = {
        horizontal: 'center',
        vertical: 'middle'
    }

    cellE1.font = fontTitle;
    cellE2.font = fontTitle;
    cellE4.font = fontTitle;
    cellE4.alignment = {
        horizontal: 'center',
        vertical: 'top'
    }
    cellE11.font = {
        name: 'Arial',
        size: 8,
        bold: true,
    };
    cellE11.alignment = {
        horizontal: 'center',
        vertical: 'middle'
    }

    cellE15.alignment = {
        horizontal: 'center',
        vertical: 'top'
    }
    cellE16.alignment = {
        horizontal: 'center',
        vertical: 'top'
    }

    cellI1.alignment = {
        horizontal: 'right',
        vertical: 'bottom'
    }
    return worksheet;
}

module.exports = createTemplate;
