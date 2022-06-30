const Excel = require('exceljs');


function test() {

    const workbook = new Excel.Workbook();
    workbook.views = [
        {
            x: 0,
            y: 0,
            width: 10000,
            height: 20000,
            firstSheet: 0,
            activeTab: 1,
            visibility: "visible",
        },
    ];
    const worksheet = workbook.addWorksheet("Daftar Perhitungan");

    const calibri = "Calibri"
    const fontStyle = {
        name: calibri,
        size: 8,
    };

    const fontStyleBold = {
        name: calibri,
        size: 8,
        bold: true
    }

    const alignmentStyle = {
        vertical: "middle",
        horizontal: "center",
        wrapText: "true"
    };

    const borderStyle = {
        top: {style: "thin"},
        bottom: {style: "thin"},
        left: {style: "thin"},
        right: {style: "thin"}
    }

    // title 1
    worksheet.mergeCells("A1:Q1");
    worksheet.getCell("A1").value = "LAPORAN PRODUKSI GERAKAN DAN PENDAPATAN PENUNDAAN KAPAL";
    worksheet.getCell("A1").font = fontStyleBold;
    worksheet.getCell("A1").alignment = alignmentStyle;

    // title 2
    worksheet.mergeCells("A2:Q2");
    worksheet.getCell("A2").value = "PT PELABUHAN INDONESIA (PERSERO) REGIONAL 2 BENGKULU";
    worksheet.getCell("A2").font = fontStyleBold;
    worksheet.getCell("A2").alignment = alignmentStyle;


    worksheet.mergeCells("A4:C4");
    worksheet.mergeCells("A5:C5");

    worksheet.getCell("A4").value = "LOKASI PELABUHAN";
    worksheet.getCell("D4").value = ":";
    worksheet.getCell("E4").value = "BENGKULU";

    worksheet.getCell("A5").value = "BULAN / TAHUN";
    worksheet.getCell("D5").value = ":";
    worksheet.getCell("E5").value = "APPRIL 2022";

    // styling
    worksheet.mergeCells("A6:A8"); // no
    worksheet.mergeCells("B6:B8"); // nama kapal tunda
    worksheet.mergeCells("C6:C8"); // hp

    worksheet.mergeCells("D6:F7"); // gerakan
    worksheet.mergeCells("G6:I7"); // total waktu
    worksheet.mergeCells("J6:K6"); // PENDAPATAN
    worksheet.mergeCells("L6:M6"); // PNBP


    const arrayHeaderBorder = [
        "A6",
        "B6",
        "C6",
        "D6",
        "G6",
        "J6",
        "L6",

        "J7",
        "K7",
        "L7",
        "M7",

        "D8",
        "E8",
        "F8",
        "G8",
        "H8",
        "I8",
        "J8",
        "K8",
        "L8",
        "M8"
    ]
    for (let i = 0; i < arrayHeaderBorder.length; i++) {
        worksheet.getCell(arrayHeaderBorder[i]).border = borderStyle
        worksheet.getCell(arrayHeaderBorder[i]).alignment = alignmentStyle
        worksheet.getCell(arrayHeaderBorder[i]).font = fontStyleBold
    }

    worksheet.getCell("A6").value = "NO";
    worksheet.getCell("B6").value = "NAMA KAPAL TUNDA";
    worksheet.getCell("C6").value = "HP";
    worksheet.getCell("D6").value = "GERAKAN";
    worksheet.getCell("G6").value = "TOTAL WAKTU PENUNDAAN (MENIT)";
    worksheet.getCell("J6").value = "PENDAPATAN";
    worksheet.getCell("L6").value = "PNBP";

    worksheet.getCell("J7").value = "PER HP";
    worksheet.getCell("K7").value = "TOTAL";
    worksheet.getCell("L7").value = "PER HP (5%)";
    worksheet.getCell("M7").value = "TOTAL (5%)";

    worksheet.getCell("D8").value = "DALAM NEGERI";
    worksheet.getCell("E8").value = "LUAR NEGERI";
    worksheet.getCell("F8").value = "TOTAL";
    worksheet.getCell("G8").value = "EFEKTIF TIME(MENIT)";
    worksheet.getCell("H8").value = "MOB-DEMOB(MENIT)";
    worksheet.getCell("I8").value = "TOTAL (MENIT)";
    worksheet.getCell("J8").value = "Rp";
    worksheet.getCell("K8").value = "Rp";
    worksheet.getCell("L8").value = "Rp";
    worksheet.getCell("M8").value = "Rp";


    let startCell = 9;

    const jumlahKapal = 10
    const gerakanKapal = 50

    for (let i = 0; i < (jumlahKapal * gerakanKapal); i++) {
        worksheet.getCell(`A${startCell}`).value = `i + 1`;
        worksheet.getCell(`B${startCell}`).value = `i + 1`;
        worksheet.getCell(`C${startCell}`).value = `i + 1`;
        worksheet.getCell(`D${startCell}`).value = `i + 1`;
        worksheet.getCell(`E${startCell}`).value = `i + 1`;
        worksheet.getCell(`F${startCell}`).value = `i + 1`;
        worksheet.getCell(`G${startCell}`).value = `i + 1`;
        worksheet.getCell(`H${startCell}`).value = `i + 1`;
        worksheet.getCell(`I${startCell}`).value = `i + 1`;
        worksheet.getCell(`J${startCell}`).value = `i + 1`;
        worksheet.getCell(`K${startCell}`).value = `i + 1`;
        worksheet.getCell(`L${startCell}`).value = `i + 1`;
        worksheet.getCell(`M${startCell}`).value = `i + 1`;

        worksheet.getCell(`A${startCell}`).border = borderStyle;
        worksheet.getCell(`B${startCell}`).border = borderStyle;
        worksheet.getCell(`C${startCell}`).border = borderStyle;
        worksheet.getCell(`D${startCell}`).border = borderStyle;
        worksheet.getCell(`E${startCell}`).border = borderStyle;
        worksheet.getCell(`F${startCell}`).border = borderStyle;
        worksheet.getCell(`G${startCell}`).border = borderStyle;
        worksheet.getCell(`H${startCell}`).border = borderStyle;
        worksheet.getCell(`I${startCell}`).border = borderStyle;
        worksheet.getCell(`J${startCell}`).border = borderStyle;
        worksheet.getCell(`K${startCell}`).border = borderStyle;
        worksheet.getCell(`L${startCell}`).border = borderStyle;
        worksheet.getCell(`M${startCell}`).border = borderStyle;

        startCell++;
    }

    workbook.xlsx.writeFile(
      "Laporan Tunda.xlsx"
    );
}

test()