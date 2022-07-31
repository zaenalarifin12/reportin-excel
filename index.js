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
    // const worksheet2 = workbook.addWorksheet("Rincian Biaya");

    const calibri = "Calibri"
    
    const fontStyle12 = {
        name: calibri,
        size: 12,
        bold: true
    };
    
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

    const alignmentStyleStart = {
        vertical: "middle",
        horizontal: "start",
        wrapText: "true"
    };

    const borderStyle = {
        top: {style: "thin"},
        bottom: {style: "thin"},
        left: {style: "thin"},
        right: {style: "thin"}
    }

    worksheet.mergeCells("A1:E1");
    worksheet.getCell("A1").value = "PT. (PERSERO) PELABUHAN INDONESIA";
    worksheet.getCell("A1").alignment = alignmentStyleStart;

    worksheet.mergeCells("A2:E2");
    worksheet.getCell("A2").value = "CABANG PELABUHAN PONTIANAK";
    worksheet.getCell("A2").alignment = alignmentStyleStart;

    worksheet.mergeCells("A3:E3");
    worksheet.getCell("A3").value = "CABANG PONTIANAK";
    worksheet.getCell("A3").alignment = alignmentStyleStart;

    worksheet.getCell("G3").value = "Form";
    worksheet.getCell("G3").alignment = alignmentStyleStart;

    worksheet.getCell("G4").value = "FM";
    worksheet.getCell("G4").alignment = alignmentStyleStart;

    worksheet.getCell("H3").value = ":LHGK";
    worksheet.getCell("H3").alignment = alignmentStyleStart;

    worksheet.getCell("H4").value = "01/01/01/27";
    worksheet.getCell("H4").alignment = alignmentStyleStart;

    worksheet.mergeCells("A5:AG5");
    worksheet.getCell("A5").value = "Laporan Harian Pemanduan Gerakan Kapal dan Keterlambatan Pelayanan Pemanduan";
    worksheet.getCell("A5").font = fontStyle12;
    worksheet.getCell("A5").alignment = alignmentStyle;

    worksheet.mergeCells("A6:AG6");
    worksheet.getCell("A6").value = "Hari : Kamis Tanggal : 01-Juli-2021 Pukul : 00:00 s/d 24:00 WIB";
    worksheet.getCell("A6").alignment = alignmentStyle;

    worksheet.mergeCells("A7:A8");
    worksheet.getCell("A7").value = "NO";
    worksheet.getCell("A7").font = fontStyleBold;
    worksheet.getCell("A7").alignment = alignmentStyle;

    worksheet.getCell("A9").value = "1";
    worksheet.getCell("A9").font = fontStyleBold;
    worksheet.getCell("A9").alignment = alignmentStyle;

    worksheet.getCell("B7").value = "No PKK";
    worksheet.getCell("B7").font = fontStyleBold;
    worksheet.getCell("B7").alignment = alignmentStyle;

    worksheet.getCell("B8").value = "Nama Kapal";
    worksheet.getCell("B8").font = fontStyleBold;
    worksheet.getCell("B8").alignment = alignmentStyle;

    worksheet.getCell("B9").value = "2";
    worksheet.getCell("B9").font = fontStyleBold;
    worksheet.getCell("B9").alignment = alignmentStyle;

    worksheet.mergeCells("C7:C8");
    worksheet.getCell("C7").value = "LOA";
    worksheet.getCell("C7").font = fontStyleBold;
    worksheet.getCell("C7").alignment = alignmentStyle;

    worksheet.getCell("C9").value = "3";
    worksheet.getCell("C9").font = fontStyleBold;
    worksheet.getCell("C9").alignment = alignmentStyle;

    worksheet.mergeCells("D7:D8");
    worksheet.getCell("D7").value = "GRT";
    worksheet.getCell("D7").font = fontStyleBold;
    worksheet.getCell("D7").alignment = alignmentStyle;

    worksheet.getCell("D9").value = "3";
    worksheet.getCell("D9").font = fontStyleBold;
    worksheet.getCell("D9").alignment = alignmentStyle;

    worksheet.mergeCells("E7:E8");
    worksheet.getCell("E7").value = "FLAG";
    worksheet.getCell("E7").font = fontStyleBold;
    worksheet.getCell("E7").alignment = alignmentStyle;

    worksheet.getCell("E9").value = "4";
    worksheet.getCell("E9").font = fontStyleBold;
    worksheet.getCell("E9").alignment = alignmentStyle;

    worksheet.mergeCells("F7:F8");
    worksheet.getCell("F7").value = "DRAFT";
    worksheet.getCell("F7").font = fontStyleBold;
    worksheet.getCell("F7").alignment = alignmentStyle;

    worksheet.getCell("F9").value = "5";
    worksheet.getCell("F9").font = fontStyleBold;
    worksheet.getCell("F9").alignment = alignmentStyle;

    worksheet.mergeCells("G7:G8");
    worksheet.getCell("G7").value = "KODE PPKB/PPKB KE ";
    worksheet.getCell("G7").font = fontStyleBold;
    worksheet.getCell("G7").alignment = alignmentStyle;

    worksheet.getCell("G9").value = "6";
    worksheet.getCell("G9").font = fontStyleBold;
    worksheet.getCell("G9").alignment = alignmentStyle;

    worksheet.mergeCells("H7:H8");
    worksheet.getCell("H7").value = "NO SPK 2A1";
    worksheet.getCell("H7").font = fontStyleBold;
    worksheet.getCell("H7").alignment = alignmentStyle;

    worksheet.getCell("H9").value = "7";
    worksheet.getCell("H9").font = fontStyleBold;
    worksheet.getCell("H9").alignment = alignmentStyle;

    worksheet.mergeCells("I7:I8");
    worksheet.getCell("I7").value = "PANDU";
    worksheet.getCell("I7").font = fontStyleBold;
    worksheet.getCell("I7").alignment = alignmentStyle;

    worksheet.getCell("I9").value = "8";
    worksheet.getCell("I9").font = fontStyleBold;
    worksheet.getCell("I9").alignment = alignmentStyle;

    worksheet.mergeCells("J7:K7");
    worksheet.getCell("J7").value = "KAPAL TIBA";
    worksheet.getCell("J7").font = fontStyleBold;
    worksheet.getCell("J7").alignment = alignmentStyle;

    worksheet.getCell("J8").value = "TANGGAL";
    worksheet.getCell("J8").font = fontStyleBold;
    worksheet.getCell("J8").alignment = alignmentStyle;

    worksheet.getCell("K8").value = "JAM";
    worksheet.getCell("K8").font = fontStyleBold;
    worksheet.getCell("K8").alignment = alignmentStyle;

    worksheet.mergeCells("J9:K9");
    worksheet.getCell("J9").value = "9";
    worksheet.getCell("J9").font = fontStyleBold;
    worksheet.getCell("J9").alignment = alignmentStyle;

    worksheet.mergeCells("L7:M7");
    worksheet.getCell("L7").value = "PERMINTAAN";
    worksheet.getCell("L7").font = fontStyleBold;
    worksheet.getCell("L7").alignment = alignmentStyle;

    worksheet.getCell("L8").value = "TGL";
    worksheet.getCell("L8").font = fontStyleBold;
    worksheet.getCell("L8").alignment = alignmentStyle;

    worksheet.getCell("M8").value = "JAM";
    worksheet.getCell("M8").font = fontStyleBold;
    worksheet.getCell("M8").alignment = alignmentStyle;

    worksheet.mergeCells("L9:M9");
    worksheet.getCell("L9").value = "10";
    worksheet.getCell("L9").font = fontStyleBold;
    worksheet.getCell("L9").alignment = alignmentStyle;

    worksheet.mergeCells("N7:O7");
    worksheet.getCell("N7").value = "PERSIAPAN OG";
    worksheet.getCell("N7").font = fontStyleBold;
    worksheet.getCell("N7").alignment = alignmentStyle;

    worksheet.getCell("N8").value = "PNK";
    worksheet.getCell("N8").font = fontStyleBold;
    worksheet.getCell("N8").alignment = alignmentStyle;

    worksheet.getCell("O8").value = "KB";
    worksheet.getCell("O8").font = fontStyleBold;
    worksheet.getCell("O8").alignment = alignmentStyle;

    worksheet.mergeCells("N9:O9");
    worksheet.getCell("N9").value = "11";
    worksheet.getCell("N9").font = fontStyleBold;
    worksheet.getCell("N9").alignment = alignmentStyle;

    worksheet.mergeCells("P7:Q7");
    worksheet.getCell("P7").value = "PELAKSANAAN";
    worksheet.getCell("P7").font = fontStyleBold;
    worksheet.getCell("P7").alignment = alignmentStyle;

    worksheet.getCell("P8").value = "MULAI";
    worksheet.getCell("P8").font = fontStyleBold;
    worksheet.getCell("P8").alignment = alignmentStyle;

    worksheet.getCell("Q8").value = "SELESAI";
    worksheet.getCell("Q8").font = fontStyleBold;
    worksheet.getCell("Q8").alignment = alignmentStyle;

    worksheet.mergeCells("P9:Q9");
    worksheet.getCell("P9").value = "12";
    worksheet.getCell("P9").font = fontStyleBold;
    worksheet.getCell("P9").alignment = alignmentStyle;

    worksheet.getCell("R7").value = "LAMA";
    worksheet.getCell("R7").font = fontStyleBold;
    worksheet.getCell("R7").alignment = alignmentStyle;

    worksheet.getCell("R8").value = "PND";
    worksheet.getCell("R8").font = fontStyleBold;
    worksheet.getCell("R8").alignment = alignmentStyle;

    worksheet.getCell("R9").value = "13";
    worksheet.getCell("R9").font = fontStyleBold;
    worksheet.getCell("R9").alignment = alignmentStyle;

    worksheet.mergeCells("S7:S8");
    worksheet.getCell("S7").value = "WT";
    worksheet.getCell("S7").font = fontStyleBold;
    worksheet.getCell("S7").alignment = alignmentStyle;

    worksheet.getCell("S9").value = "14";
    worksheet.getCell("S9").font = fontStyleBold;
    worksheet.getCell("S9").alignment = alignmentStyle;

    worksheet.mergeCells("T7:T8");
    worksheet.getCell("T7").value = "AT";
    worksheet.getCell("T7").font = fontStyleBold;
    worksheet.getCell("T7").alignment = alignmentStyle;

    worksheet.getCell("T9").value = "15";
    worksheet.getCell("T9").font = fontStyleBold;
    worksheet.getCell("T9").alignment = alignmentStyle;

    worksheet.mergeCells("U7:U8");
    worksheet.getCell("U7").value = "PT";
    worksheet.getCell("U7").font = fontStyleBold;
    worksheet.getCell("U7").alignment = alignmentStyle;

    worksheet.getCell("U9").value = "16";
    worksheet.getCell("U9").font = fontStyleBold;
    worksheet.getCell("U9").alignment = alignmentStyle;

    worksheet.mergeCells("V7:V8");
    worksheet.getCell("V7").value = "TRT";
    worksheet.getCell("V7").font = fontStyleBold;
    worksheet.getCell("V7").alignment = alignmentStyle;

    worksheet.getCell("V9").value = "17";
    worksheet.getCell("V9").font = fontStyleBold;
    worksheet.getCell("V9").alignment = alignmentStyle;

    worksheet.mergeCells("W7:X7");
    worksheet.getCell("W7").value = "LOKASI";
    worksheet.getCell("W7").font = fontStyleBold;
    worksheet.getCell("W7").alignment = alignmentStyle;


    // worksheet.mergeCells("V7:W8");
    worksheet.getCell("W8").value = "DARI";
    worksheet.getCell("W8").font = fontStyleBold;
    worksheet.getCell("W8").alignment = alignmentStyle;

    worksheet.getCell("W9").value = "18";
    worksheet.getCell("W9").font = fontStyleBold;
    worksheet.getCell("W9").alignment = alignmentStyle;

    // worksheet.mergeCells("V7:W8");
    worksheet.getCell("X8").value = "KE";
    worksheet.getCell("X8").font = fontStyleBold;
    worksheet.getCell("X8").alignment = alignmentStyle;

    worksheet.getCell("X9").value = "19";
    worksheet.getCell("X9").font = fontStyleBold;
    worksheet.getCell("X9").alignment = alignmentStyle;

    worksheet.mergeCells("Y7:Y8");
    worksheet.getCell("Y7").value = "GRK";
    worksheet.getCell("Y7").font = fontStyleBold;
    worksheet.getCell("Y7").alignment = alignmentStyle;

    worksheet.getCell("Y9").value = "20";
    worksheet.getCell("Y9").font = fontStyleBold;
    worksheet.getCell("Y9").alignment = alignmentStyle;

    worksheet.mergeCells("Z7:Z8");
    worksheet.getCell("Z7").value = "AGEN";
    worksheet.getCell("Z7").font = fontStyleBold;
    worksheet.getCell("Z7").alignment = alignmentStyle;

    worksheet.getCell("Z9").value = "18";
    worksheet.getCell("Z9").font = fontStyleBold;
    worksheet.getCell("Z9").alignment = alignmentStyle;

    worksheet.mergeCells("AA7:AE7");
    worksheet.getCell("AA7").value = "PEMAKAIAN TUNDA";
    worksheet.getCell("AA7").font = fontStyleBold;
    worksheet.getCell("AA7").alignment = alignmentStyle;

    worksheet.mergeCells("AA8:AC8");
    worksheet.getCell("AA8").value = "NAMA";
    worksheet.getCell("AA8").font = fontStyleBold;
    worksheet.getCell("AA8").alignment = alignmentStyle;

    worksheet.getCell("AA9").value = "19";
    worksheet.getCell("AA9").font = fontStyleBold;
    worksheet.getCell("AA9").alignment = alignmentStyle;

    worksheet.getCell("AB9").value = "20";
    worksheet.getCell("AB9").font = fontStyleBold;
    worksheet.getCell("AB9").alignment = alignmentStyle;

    worksheet.getCell("AC9").value = "21";
    worksheet.getCell("AC9").font = fontStyleBold;
    worksheet.getCell("AC9").alignment = alignmentStyle;

    worksheet.mergeCells("AD8:AE8");
    worksheet.getCell("AD8").value = "JAM";
    worksheet.getCell("AD8").font = fontStyleBold;
    worksheet.getCell("AD8").alignment = alignmentStyle;

    worksheet.getCell("AD9").value = "22";
    worksheet.getCell("AD9").font = fontStyleBold;
    worksheet.getCell("AD9").alignment = alignmentStyle;

    worksheet.getCell("AE9").value = "23";
    worksheet.getCell("AE9").font = fontStyleBold;
    worksheet.getCell("AE9").alignment = alignmentStyle;

    worksheet.mergeCells("AF7:AF8");
    worksheet.getCell("AF7").value = "LAMA";
    worksheet.getCell("AF7").font = fontStyleBold;
    worksheet.getCell("AF7").alignment = alignmentStyle;

    worksheet.getCell("AF9").value = "23";
    worksheet.getCell("AF9").font = fontStyleBold;
    worksheet.getCell("AF9").alignment = alignmentStyle;

    worksheet.mergeCells("AG7:AG8");
    worksheet.getCell("AG7").value = "KETERANGAN";
    worksheet.getCell("AG7").font = fontStyleBold;
    worksheet.getCell("Ag7").alignment = alignmentStyle;

    worksheet.getCell("AG9").value = "23";
    worksheet.getCell("AG9").font = fontStyleBold;
    worksheet.getCell("AG9").alignment = alignmentStyle;

    const arrayHeaderBorder = [
        "A7","A9",
        "B7", "B8", 'B9',
        "C7", "C8", "C9",
        "D7", "D9",
        "E7", "E9",
        "F7", "F9",
        "G7", "G9",
        "H7", "H9",
        "I7", "I9",
        "J7", "J8", "J9",
        "K7", "K8", "K9",
        "L7", "L8", "L9",
        "M7", "M8", "M9",
        "N7", "N8", "N9",
        "O8", "O9",
        "P7", "P8", "P9",
        "Q7", "Q8",
        "R7", "R8", "R9",
        "S7", "S9",
        "T7", "T9",
        "U7", "U9",
        "V7", "V9",
        "W7", "W8", "W9",
        "X7", "X8", "X9",
        "Y7", "Y8", "Y9",
        "Z7", "Z8", "Z9",
        "Z7", "Z9",
        "AA7", "AA8", "AA9",
        "AB7",  "AB9",
        "AC7",  "AC9",
        "AD7", "AD8", "AE9",
        "AE7", "AE8", "AF9",
        "AF7", "AF8", "AG9",
        "AG7", "AG8", "AH9",
    ]
    for (let i = 0; i < arrayHeaderBorder.length; i++) {
        worksheet.getCell(arrayHeaderBorder[i]).border = borderStyle
        worksheet.getCell(arrayHeaderBorder[i]).alignment = alignmentStyle
        worksheet.getCell(arrayHeaderBorder[i]).font = fontStyleBold
    }

    // worksheet.getCell("A5").value = "NO";
    // worksheet.getCell("B5").value = "NOMOR NOTA";
    // worksheet.getCell("C5").value = "TANGGAL INVOICE";
    // worksheet.getCell("D5").value = "NAMA KAPAL";
    // worksheet.getCell("E5").value = "BENDERA";
    // worksheet.getCell("F5").value = "AGENT";
    // worksheet.getCell("G5").value = "KURS USD";
    // worksheet.getCell("H5").value = "GT";
    // worksheet.getCell("I5").value = "LOA";
    // worksheet.getCell("J5").value = "TGL";
    // worksheet.getCell("K5").value = "GERAKAN KAPAL";

    // worksheet.getCell("L5").value = "PEMANDUAN";
    // worksheet.getCell("O5").value = "PENUNDAAN";

    // worksheet.getCell("L7").value = "Luar Negeri";
    // worksheet.getCell("N7").value = "Dalam Negeri";

    // worksheet.getCell("O7").value = "Luar Negeri";
    // worksheet.getCell("Q7").value = "Dalam Negeri";

    // //
    // worksheet.getCell("L8").value = "US";
    // worksheet.getCell("M8").value = "Dalam Rp";
    // worksheet.getCell("N8").value = "RP";

    // worksheet.getCell("O8").value = "US";
    // worksheet.getCell("P8").value = "Dalam Rp";
    // worksheet.getCell("Q8").value = "RP";


    let startCell = 10;
    // perbulan
    // const jumlahKapal = 1000
    // const gerakanKapal = 50

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
        worksheet.getCell(`V${startCell}`).value = `i + 1`;
        worksheet.getCell(`W${startCell}`).value = `i + 1`;
        worksheet.getCell(`X${startCell}`).value = `i + 1`;
        worksheet.getCell(`Y${startCell}`).value = `i + 1`;
        worksheet.getCell(`Z${startCell}`).value = `i + 1`;
        worksheet.getCell(`AA${startCell}`).value = `i + 1`;
        worksheet.getCell(`AB${startCell}`).value = `i + 1`;
        worksheet.getCell(`AC${startCell}`).value = `i + 1`;
        worksheet.getCell(`AD${startCell}`).value = `i + 1`;
        worksheet.getCell(`AE${startCell}`).value = `i + 1`;
        worksheet.getCell(`AF${startCell}`).value = `i + 1`;
        worksheet.getCell(`AG${startCell}`).value = `i + 1`;

        // styling
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
        worksheet.getCell(`Q${startCell}`).border = borderStyle;
        worksheet.getCell(`R${startCell}`).border = borderStyle;
        worksheet.getCell(`S${startCell}`).border = borderStyle;
        worksheet.getCell(`T${startCell}`).border = borderStyle;
        worksheet.getCell(`U${startCell}`).border = borderStyle;
        worksheet.getCell(`V${startCell}`).border = borderStyle;
        worksheet.getCell(`W${startCell}`).border = borderStyle;
        worksheet.getCell(`X${startCell}`).border = borderStyle;
        worksheet.getCell(`Y${startCell}`).border = borderStyle;
        worksheet.getCell(`Z${startCell}`).border = borderStyle;
        worksheet.getCell(`AA${startCell}`).border = borderStyle;
        worksheet.getCell(`AB${startCell}`).border = borderStyle;
        worksheet.getCell(`AC${startCell}`).border = borderStyle;
        worksheet.getCell(`AD${startCell}`).border = borderStyle;
        worksheet.getCell(`AE${startCell}`).border = borderStyle;
        worksheet.getCell(`AF${startCell}`).border = borderStyle;
        worksheet.getCell(`AG${startCell}`).border = borderStyle;

        startCell++;
    }

    const startCellPlus1 = startCell + 1
    const startCellPlus2 = startCell + 2
    const startCellPlus3 = startCell + 3
    const startCellPlus4 = startCell + 4
    const startCellPlus5 = startCell + 5
    const startCellPlus6 = startCell + 6
    const startCellPlus7 = startCell + 7
    const startCellPlus8 = startCell + 8
    const startCellPlus9 = startCell + 9

    worksheet.getCell(`A${startCellPlus2}`).value = "1"
    worksheet.getCell(`A${startCellPlus3}`).value = "2"
    worksheet.getCell(`A${startCellPlus4}`).value = "3"
    worksheet.getCell(`A${startCellPlus5}`).value = "4"
    worksheet.getCell(`A${startCellPlus6}`).value = "5"
    worksheet.getCell(`A${startCellPlus7}`).value = "6"
    worksheet.getCell(`A${startCellPlus8}`).value = "7"

    worksheet.getCell(`B${startCellPlus1}`).value = "Keterangan"
    worksheet.getCell(`B${startCellPlus2}`).value = "Kapal Dilayani Masuk (M)"
    worksheet.getCell(`B${startCellPlus3}`).value = "Kapal Dilayani Keluar (K)"
    worksheet.getCell(`B${startCellPlus4}`).value = "Kapal Dilayani Pindah (P)"
    worksheet.getCell(`B${startCellPlus5}`).value = "Kapal Batal (B)"
    worksheet.getCell(`B${startCellPlus6}`).value = "Kapal yang Menggunakan Tunda"
    worksheet.getCell(`B${startCellPlus7}`).value = "Kapal Langsung Masuk (LM)"
    worksheet.getCell(`B${startCellPlus8}`).value = "Jumlah Gerakan"

    worksheet.getCell(`A${startCellPlus1}`).value = ":"
    worksheet.getCell(`A${startCellPlus2}`).value = ":"
    worksheet.getCell(`A${startCellPlus3}`).value = ":"
    worksheet.getCell(`A${startCellPlus4}`).value = ":"
    worksheet.getCell(`A${startCellPlus5}`).value = ":"
    worksheet.getCell(`A${startCellPlus6}`).value = ":"
    worksheet.getCell(`A${startCellPlus7}`).value = ":"
    worksheet.getCell(`A${startCellPlus8}`).value = ":"

    worksheet.getCell(`C${startCellPlus1}`).value = "0"
    worksheet.getCell(`C${startCellPlus2}`).value = "0"
    worksheet.getCell(`C${startCellPlus3}`).value = "0"
    worksheet.getCell(`C${startCellPlus4}`).value = "0"
    worksheet.getCell(`C${startCellPlus5}`).value = "0"
    worksheet.getCell(`C${startCellPlus6}`).value = "0"
    worksheet.getCell(`C${startCellPlus7}`).value = "0"
    worksheet.getCell(`C${startCellPlus8}`).value = "0"

    worksheet.getCell(`D${startCellPlus1}`).value = ""
    worksheet.getCell(`D${startCellPlus2}`).value = ""
    worksheet.getCell(`D${startCellPlus3}`).value = ""
    worksheet.getCell(`D${startCellPlus4}`).value = ""
    worksheet.getCell(`D${startCellPlus5}`).value = ""
    worksheet.getCell(`D${startCellPlus6}`).value = ""
    worksheet.getCell(`D${startCellPlus7}`).value = ""
    worksheet.getCell(`D${startCellPlus8}`).value = "0%"


    worksheet.getCell(`E${startCellPlus1}`).value = "KAPAL"
    worksheet.getCell(`E${startCellPlus2}`).value = "KAPAL"
    worksheet.getCell(`E${startCellPlus3}`).value = "KAPAL"
    worksheet.getCell(`E${startCellPlus4}`).value = "KAPAL"
    worksheet.getCell(`E${startCellPlus5}`).value = "KAPAL"
    worksheet.getCell(`E${startCellPlus6}`).value = "KAPAL"
    worksheet.getCell(`E${startCellPlus7}`).value = "KAPAL"
    worksheet.getCell(`E${startCellPlus8}`).value = "KAPAL"

    worksheet.mergeCells(`H${startCellPlus2}:O${startCellPlus2}`)
    worksheet.mergeCells(`H${startCellPlus3}:O${startCellPlus3}`)
    worksheet.mergeCells(`H${startCellPlus9}:O${startCellPlus9}`)

    worksheet.getCell(`H${startCellPlus2}`).value = "Mengetahui"
    worksheet.getCell(`H${startCellPlus3}`).value = "PONTIANAK, 29-07-2022"
    worksheet.getCell(`H${startCellPlus9}`).value = "NIPP. "

    worksheet.getCell(`H${startCellPlus2}`).alignment = alignmentStyle
    worksheet.getCell(`H${startCellPlus3}`).alignment = alignmentStyle
    worksheet.getCell(`H${startCellPlus9}`).alignment = alignmentStyle

    workbook.xlsx.writeFile(
      "LAPORAN-LHGK-2.xlsx"
    );
}

test()