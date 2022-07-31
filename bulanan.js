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
    const worksheet = workbook.addWorksheet("Monitoring Pemanduan ");

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
    worksheet.getCell("A1").value = "LAPORAN BULANAN KEGIATAN OPERASIONAL PELIMPAHAN PEMANDUAN DAN PENUNDAAN KAPAL";
    worksheet.getCell("A1").font = fontStyleBold;
    worksheet.getCell("A1").alignment = alignmentStyle;

    worksheet.mergeCells("A2:C2");
    worksheet.mergeCells("A3:C3");
    worksheet.mergeCells("A4:C4");
    worksheet.mergeCells("A5:C5");


    worksheet.getCell("A2").value = "PELAKSANA PEMANDUAN (TERSUS/BUP)";
    worksheet.getCell("D2").value = ":";
    worksheet.getCell("E2").value = "PT. PELABUHAN INDONESIA II (PERSERO)";

    worksheet.getCell("A3").value = "LOKASI PELABUHAN";
    worksheet.getCell("D3").value = ":";
    worksheet.getCell("E3").value = "BENGKULU";


    worksheet.getCell("A4").value = "BULAN / TAHUN";
    worksheet.getCell("D4").value = ":";
    worksheet.getCell("E4").value = "APRIL 2022";

    worksheet.getCell("A5").value = "KETERANGAN";
    worksheet.getCell("D5").value = ":";
    worksheet.getCell("E5").value = "KEGIATAN PEMANDUAN DAN PENUNDAAN PT. PELABUHAN INDONESIA II (PERSERO)";

    // styling
    worksheet.mergeCells("A6:A9"); // no
    worksheet.mergeCells("B6:B9"); // nama kapal tunda
    worksheet.mergeCells("C6:C9"); // bendera

    worksheet.mergeCells("D6:F7"); // kunjungan
    worksheet.mergeCells("G6:I7"); // gerakan
    worksheet.mergeCells("J6:K7"); // tunda
    worksheet.mergeCells("L6:O7"); // pendapatan

    worksheet.mergeCells("L8:M8"); // pendapatan
    worksheet.mergeCells("N8:O8"); // pendapatan


    worksheet.mergeCells("P6:S6"); // PNBP
    worksheet.mergeCells("P7:Q7"); // PEMANDUAN
    worksheet.mergeCells("R7:S7"); // PENUNDAAN

    worksheet.mergeCells("T6:U8"); // JUMLAH PNBP


    worksheet.mergeCells("D8:D9"); // JUMLAH PNBP
    worksheet.mergeCells("E8:E9"); // JUMLAH PNBP
    worksheet.mergeCells("F8:F9"); // JUMLAH PNBP
    worksheet.mergeCells("G8:G9"); // JUMLAH PNBP
    worksheet.mergeCells("H8:H9"); // JUMLAH PNBP
    worksheet.mergeCells("I8:I9"); // JUMLAH PNBP
    worksheet.mergeCells("J8:J9"); // JUMLAH PNBP
    worksheet.mergeCells("K8:K9"); // JUMLAH PNBP

    const arrayHeaderBorder = [
        "A6",
        "B6",
        "C6",
        "D6",
        "G6",
        "J6",
        "L6",
        "P6",
        "T6",
        "P7",
        "R7",
        "D8",
        "E8",
        "F8",
        "G8",
        "H8",
        "I8",
        "J8",
        "K8",
        "L8",
        "N8",
        "P8",
        "R8",
        "M9",
        "N9",
        "O9",
        "P9",
        "Q9",
        "R9",
        "S9",
        "T9",
        "U9",


    ]
    for (let i = 0; i < arrayHeaderBorder.length; i++) {
        worksheet.getCell(arrayHeaderBorder[i]).border = borderStyle
        worksheet.getCell(arrayHeaderBorder[i]).alignment = alignmentStyle
        worksheet.getCell(arrayHeaderBorder[i]).font = fontStyleBold
    }

    worksheet.getCell("A6").value = "NO";
    worksheet.getCell("B6").value = "NAMA KAPAL";
    worksheet.getCell("C6").value = "BENDERA";
    worksheet.getCell("D6").value = "Kunjungan Kapal";
    worksheet.getCell("G6").value = "GERAKAN KAPAL";
    worksheet.getCell("J6").value = "TUNDA";
    worksheet.getCell("L6").value = "PENDAPATAN";

    worksheet.getCell("P6").value = "PNBP";
    worksheet.getCell("T6").value = "JUMLAH PNBP";

    worksheet.getCell("P7").value = "PEMANDUAN";
    worksheet.getCell("R7").value = "PENUNDAAN";

    worksheet.getCell("D8").value = "GRT";
    worksheet.getCell("E8").value = "DWT";
    worksheet.getCell("F8").value = "LOA";
    worksheet.getCell("G8").value = "GERAKAN KAPAL";
    worksheet.getCell("H8").value = "MULAI";
    worksheet.getCell("I8").value = "SELESAI";
    worksheet.getCell("J8").value = "UNIT";
    worksheet.getCell("K8").value = "JAM";

    worksheet.getCell("L8").value = "PEMANDUAN";
    worksheet.getCell("N8").value = "PENUNDAAN";

    worksheet.getCell("P8").value = "5%";
    worksheet.getCell("R8").value = "5%";

    worksheet.getCell("L9").value = "Rp";
    worksheet.getCell("M9").value = "US $";

    worksheet.getCell("N9").value = "Rp";
    worksheet.getCell("O9").value = "US $";

    worksheet.getCell("P9").value = "Rp";
    worksheet.getCell("Q9").value = "US $";
    worksheet.getCell("R9").value = "Rp";
    worksheet.getCell("S9").value = "US $";
    worksheet.getCell("T9").value = "Rp";
    worksheet.getCell("U9").value = "US $";


    let startCell = 10;

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
        worksheet.getCell(`N${startCell}`).value = `i + 1`;
        worksheet.getCell(`O${startCell}`).value = `i + 1`;
        worksheet.getCell(`P${startCell}`).value = `i + 1`;
        worksheet.getCell(`Q${startCell}`).value = `i + 1`;
        worksheet.getCell(`R${startCell}`).value = `i + 1`;
        worksheet.getCell(`S${startCell}`).value = `i + 1`;
        worksheet.getCell(`T${startCell}`).value = `i + 1`;
        worksheet.getCell(`U${startCell}`).value = `i + 1`;

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
        worksheet.getCell(`N${startCell}`).border = borderStyle;
        worksheet.getCell(`O${startCell}`).border = borderStyle;
        worksheet.getCell(`P${startCell}`).border = borderStyle;
        worksheet.getCell(`Q${startCell}`).border = borderStyle;
        worksheet.getCell(`R${startCell}`).border = borderStyle;
        worksheet.getCell(`S${startCell}`).border = borderStyle;
        worksheet.getCell(`T${startCell}`).border = borderStyle;
        worksheet.getCell(`U${startCell}`).border = borderStyle;


        startCell++;
    }

    workbook.xlsx.writeFile(
      "Laporan Bulanan.xlsx"
    );
}

test()