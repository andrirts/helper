const ExcelJS = require('exceljs');

function applyDefaultStyle(sheet) {
  const defaultFont = {
    name: "Arial",
    size: 8,
    bold: false, // Optional: Set bold font
    italic: false, // Optional: Set italic font
  };

  sheet.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.font = defaultFont;
    });
  });
}

function createTemplatePrepaid(workbook, sheetPage) {
  const worksheet = workbook.addWorksheet(`Sheet ${sheetPage}`);

  // worksheet.getColumn("A").width = 12.9;
  // worksheet.getColumn("B").width = 2.99;
  // worksheet.getColumn("C").width = 19.81;
  // worksheet.getColumn("D").width = 3.08;
  // worksheet.getColumn("E").width = 14.99;
  // worksheet.getColumn("F").width = 2.9;
  // worksheet.getColumn("G").width = 17.81;
  // worksheet.getColumn("H").width = 3.9;
  // worksheet.getColumn("I").width = 15.99;
  // worksheet.getColumn("J").width = 2.81;
  // worksheet.getColumn("K").width = 17.36;
  // // worksheet.getColumn("L").width = 18.73;
  // worksheet.getRow(13).height = 27.75;

  worksheet.getColumn('A').width = 9;
  worksheet.getColumn('B').width = 1.91;
  worksheet.getColumn('C').width = 18.09;
  worksheet.getColumn('D').width = 5.09;
  worksheet.getColumn('E').width = 12.64;
  worksheet.getColumn('F').width = 1.91;
  worksheet.getColumn('G').width = 16.82;
  worksheet.getColumn('H').width = 4.19;
  worksheet.getColumn('I').width = 14.72;
  worksheet.getColumn('J').width = 1.91;
  worksheet.getColumn('K').width = 15.64;
  worksheet.getColumn('L').width = 18.73;

  //Merging Cells
  worksheet.mergeCells("A4:D4");
  worksheet.mergeCells("E4:K4");
  worksheet.mergeCells("E15:I15");
  worksheet.mergeCells("E16:I16");
  worksheet.mergeCells("E17:I17");
  worksheet.mergeCells("A19:C19");
  worksheet.mergeCells("E19:G19");
  worksheet.mergeCells("I19:K19");

  const fontTitle = {
    name: "Arial",
    size: 8,
    bold: true, // Optional: Set bold font
    italic: false, // Optional: Set italic font
  };

  // Cell A
  // Write data to cell A1
  const cellA1 = worksheet.getCell("A1");
  cellA1.value = "BANK MANDIRI";

  const cellA2 = worksheet.getCell("A2");
  cellA2.value = "RTS 01 (54RTSOP1 )";

  const cellA4 = worksheet.getCell("A4");
  cellA4.value = "STRUK LISTRIK PRABAYAR";

  const cellA6 = worksheet.getCell("A6");
  cellA6.value = "NO METER";

  const cellA7 = worksheet.getCell("A7");
  cellA7.value = "IDPEL";

  const cellA8 = worksheet.getCell("A8");
  cellA8.value = "NAMA";

  const cellA9 = worksheet.getCell("A9");
  cellA9.value = "TARIF/DAYA";

  const cellA10 = worksheet.getCell("A10");
  cellA10.value = "NO REF";

  const cellA12 = worksheet.getCell("A12");
  cellA12.value = "RP BAYAR";

  for (let row = 6; row <= 12; row++) {
    if (row === 11) {
      continue;
    }

    worksheet.getCell(`B${row}`).value = ":";
    worksheet.getCell(`B${row}`).alignment = {
      vertical: "bottom",
      horizontal: "center",
    };
  }

  const cellE1 = worksheet.getCell("E1");
  cellE1.value = "BANK MANDIRI";

  const cellE2 = worksheet.getCell("E2");
  cellE2.value = "RTS 01 (54RTSOP1 )";

  const cellE4 = worksheet.getCell("E4");
  cellE4.value = "STRUK PEMBELIAN LISTRIK PRABAYAR";

  const cellE6 = worksheet.getCell("E6");
  cellE6.value = "NO METER";

  const cellE7 = worksheet.getCell("E7");
  cellE7.value = "IDPEL";

  const cellE8 = worksheet.getCell("E8");
  cellE8.value = "NAMA";

  const cellE9 = worksheet.getCell("E9");
  cellE9.value = "TARIF/DAYA";

  const cellE10 = worksheet.getCell("E10");
  cellE10.value = "NO REF";

  const cellE11 = worksheet.getCell("E11");
  cellE11.value = "RP BAYAR";

  const cellE13 = worksheet.getCell("E13");
  cellE13.value = "STROOM/TOKEN";

  const cellE15 = worksheet.getCell("E15");
  cellE15.value = "MUP - MANDIRI";

  const cellE16 = worksheet.getCell("E16");
  cellE16.value = "Informasi Hubungi Call Center 123 atau hubungi PLN Terdekat";

  const cellE17 = worksheet.getCell("E17");
  cellE17.value = "Download PLN Mobile";

  for (let row = 6; row <= 13; row++) {
    if (row == 12) {
      continue
    }
    worksheet.getCell(`F${row}`).value = ":";
    worksheet.getCell(`F${row}`).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
  }

  //Cell I
  const cellI1 = worksheet.getCell("I1");
  cellI1.value = "TGL BAYAR :";

  const cellI6 = worksheet.getCell("I6");
  cellI6.value = "MATERAI";

  const cellI7 = worksheet.getCell("I7");
  cellI7.value = "PPN";

  const cellI8 = worksheet.getCell("I8");
  cellI8.value = "PPJ";

  const cellI9 = worksheet.getCell("I9");
  cellI9.value = "ANGSURAN";

  const cellI10 = worksheet.getCell("I10");
  cellI10.value = "RP STROOM/TOKEN";

  const cellI11 = worksheet.getCell("I11");
  cellI11.value = "JML KWH";

  const cellI12 = worksheet.getCell("I12");
  cellI12.value = "ADMIN BANK";

  for (let row = 1; row <= 12; row++) {
    if (row == 2 || row == 3 || row == 4 || row == 5) {
      continue;
    }
    worksheet.getCell(`J${row}`).value = ":";
    worksheet.getCell(`J${row}`).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
  }
  applyDefaultStyle(worksheet);
  cellA1.font = fontTitle;
  cellA2.font = fontTitle;
  cellA4.font = fontTitle;
  cellA4.alignment = {
    horizontal: "center",
    vertical: "middle",
  };

  cellE1.font = fontTitle;
  cellE2.font = fontTitle;
  cellE4.font = fontTitle;
  cellE4.alignment = {
    horizontal: "center",
    vertical: "top",
  };

  cellE15.alignment = {
    horizontal: 'center',
    vertical: 'bottom'
  }

  cellE16.alignment = {
    horizontal: 'center',
    vertical: 'bottom'
  }

  cellE17.alignment = {
    horizontal: 'center',
    vertical: 'bottom'
  }

  return worksheet;
}

module.exports = createTemplatePrepaid;
