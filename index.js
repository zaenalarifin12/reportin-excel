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
    const worksheet2 = workbook.addWorksheet("Rincian Biaya");

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
    worksheet.getCell("A1").value = "DAFTAR PERHITUNGAN BIAYA KONTRIBUSI JASA PEMANDUAN & PENUNDAAN";
    worksheet.getCell("A1").font = fontStyleBold;
    worksheet.getCell("A1").alignment = alignmentStyle;

    // title 2
    worksheet.mergeCells("A2:Q2");
    worksheet.getCell("A2").value = "PT PELABUHAN INDONESIA (PERSERO) REGIONAL 2 BENGKULU";
    worksheet.getCell("A2").font = fontStyleBold;
    worksheet.getCell("A2").alignment = alignmentStyle;


    // title 3
    worksheet.mergeCells("A3:Q3");
    worksheet.getCell("A3").value = "PERIODE BULAN MARET 2022";
    worksheet.getCell("A3").font = fontStyleBold;
    worksheet.getCell("A3").alignment = alignmentStyle;


    // styling
    worksheet.mergeCells("A5:A8");
    worksheet.mergeCells("B5:B8");
    worksheet.mergeCells("C5:C8");
    worksheet.mergeCells("D5:D8");
    worksheet.mergeCells("E5:E8");
    worksheet.mergeCells("F5:F8");
    worksheet.mergeCells("G5:G8");
    worksheet.mergeCells("H5:H8");
    worksheet.mergeCells("I5:I8");
    worksheet.mergeCells("J5:J8");
    worksheet.mergeCells("K5:K8");

    worksheet.mergeCells("L5:N6");
    worksheet.mergeCells("O5:Q6");

    worksheet.mergeCells("L7:M7");

    worksheet.mergeCells("O7:P7");


    const arrayHeaderBorder = [
        "A5",
        "B5",
        "C5",
        "D5",
        "E5",
        "F5",
        "G5",
        "H5",
        "I5",
        "J5",
        "K5",
        "L5", "L7", "L8",
        "M7", "M8",
        "N7", "N8",
        "O5", "O7", "O8",
        "P7", "P8",
        "Q7", "Q8"
    ]
    for (let i = 0; i < arrayHeaderBorder.length; i++) {
        worksheet.getCell(arrayHeaderBorder[i]).border = borderStyle
        worksheet.getCell(arrayHeaderBorder[i]).alignment = alignmentStyle
        worksheet.getCell(arrayHeaderBorder[i]).font = fontStyleBold
    }

    worksheet.getCell("A5").value = "NO";
    worksheet.getCell("B5").value = "NOMOR NOTA";
    worksheet.getCell("C5").value = "TANGGAL INVOICE";
    worksheet.getCell("D5").value = "NAMA KAPAL";
    worksheet.getCell("E5").value = "BENDERA";
    worksheet.getCell("F5").value = "AGENT";
    worksheet.getCell("G5").value = "KURS USD";
    worksheet.getCell("H5").value = "GT";
    worksheet.getCell("I5").value = "LOA";
    worksheet.getCell("J5").value = "TGL";
    worksheet.getCell("K5").value = "GERAKAN KAPAL";

    worksheet.getCell("L5").value = "PEMANDUAN";
    worksheet.getCell("O5").value = "PENUNDAAN";

    worksheet.getCell("L7").value = "Luar Negeri";
    worksheet.getCell("N7").value = "Dalam Negeri";

    worksheet.getCell("O7").value = "Luar Negeri";
    worksheet.getCell("Q7").value = "Dalam Negeri";

    //
    worksheet.getCell("L8").value = "US";
    worksheet.getCell("M8").value = "Dalam Rp";
    worksheet.getCell("N8").value = "RP";

    worksheet.getCell("O8").value = "US";
    worksheet.getCell("P8").value = "Dalam Rp";
    worksheet.getCell("Q8").value = "RP";


    let startCell = 9;
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
        startCell++;
    }

    const startCellPlus1 = startCell + 1
    const startCellPlus2 = startCell + 2
    const startCellPlus3 = startCell + 3
    const startCellPlus4 = startCell + 4

    // merge
    worksheet.mergeCells(`L${startCellPlus1}:O${startCellPlus1}`)
    worksheet.mergeCells(`L${startCellPlus2}:O${startCellPlus2}`)
    worksheet.mergeCells(`L${startCellPlus3}:O${startCellPlus3}`)
    worksheet.mergeCells(`L${startCellPlus4}:O${startCellPlus4}`)

    //value
    worksheet.getCell(`L${startCellPlus1}`).value = "JENIS PENDAPATAN"
    worksheet.getCell(`L${startCellPlus2}`).value = "PENDAPATAN KOTOR PANDU"
    worksheet.getCell(`L${startCellPlus3}`).value = "PENDAPATAN KOTOR TUNDA"
    worksheet.getCell(`L${startCellPlus4}`).value = "JUMLAH"

    worksheet.getCell(`P${startCellPlus1}`).value = "PEND. KOTOR"
    worksheet.getCell(`P${startCellPlus2}`).value = "PENDAPATAN KOTOR PANDU"
    worksheet.getCell(`P${startCellPlus3}`).value = "PENDAPATAN KOTOR TUNDA"
    worksheet.getCell(`P${startCellPlus4}`).value = "JUMLAH"

    worksheet.getCell(`Q${startCellPlus1}`).value = "PNBP 5%"
    worksheet.getCell(`Q${startCellPlus2}`).value = "PENDAPATAN KOTOR PANDU"
    worksheet.getCell(`Q${startCellPlus3}`).value = "PENDAPATAN KOTOR TUNDA"
    worksheet.getCell(`Q${startCellPlus4}`).value = "JUMLAH"

    //BORDER
    worksheet.getCell(`L${startCellPlus1}`).border = borderStyle
    worksheet.getCell(`L${startCellPlus2}`).border = borderStyle
    worksheet.getCell(`L${startCellPlus3}`).border = borderStyle
    worksheet.getCell(`L${startCellPlus4}`).border = borderStyle

    worksheet.getCell(`P${startCellPlus1}`).border = borderStyle
    worksheet.getCell(`P${startCellPlus2}`).border = borderStyle
    worksheet.getCell(`P${startCellPlus3}`).border = borderStyle
    worksheet.getCell(`P${startCellPlus4}`).border = borderStyle

    worksheet.getCell(`Q${startCellPlus1}`).border = borderStyle
    worksheet.getCell(`Q${startCellPlus2}`).border = borderStyle
    worksheet.getCell(`Q${startCellPlus3}`).border = borderStyle
    worksheet.getCell(`Q${startCellPlus4}`).border = borderStyle

    //font
    worksheet.getCell(`L${startCellPlus1}`).font = fontStyleBold
    worksheet.getCell(`P${startCellPlus1}`).font = fontStyleBold
    worksheet.getCell(`Q${startCellPlus1}`).font = fontStyleBold

    const startCellPlus6 = startCell + 6

    worksheet.getCell(`F${startCellPlus6}`).value = "Delapan Puluh Delapan Juta Enam Ratus Lima Puluh Satu Ribu Tiga Ratus Sembilan Rupiah"
    worksheet.getCell(`F${startCellPlus6}`).font = {italic: true, bold: true}

    const startCellPlus8 = startCell + 8
    const startCellPlus9 = startCell + 9
    const startCellPlus10 = startCell + 10
    const startCellPlus11 = startCell + 11
    const startCellPlus12 = startCell + 12

    // MERGE
    worksheet.mergeCells(`A${startCellPlus8}:E${startCellPlus8}`)
    worksheet.mergeCells(`A${startCellPlus9}:E${startCellPlus9}`)
    worksheet.mergeCells(`A${startCellPlus10}:E${startCellPlus10}`)
    worksheet.mergeCells(`A${startCellPlus11}:E${startCellPlus11}`)
    worksheet.mergeCells(`A${startCellPlus12}:E${startCellPlus12}`)

    worksheet.mergeCells(`M${startCellPlus8}:Q${startCellPlus8}`)
    worksheet.mergeCells(`M${startCellPlus9}:Q${startCellPlus9}`)
    worksheet.mergeCells(`M${startCellPlus10}:Q${startCellPlus10}`)
    worksheet.mergeCells(`M${startCellPlus11}:Q${startCellPlus11}`)
    worksheet.mergeCells(`M${startCellPlus12}:Q${startCellPlus12}`)

    // STYLE
    worksheet.getCell(`A${startCellPlus8}`).font = fontStyleBold
    worksheet.getCell(`A${startCellPlus9}`).font = fontStyleBold
    worksheet.getCell(`A${startCellPlus10}`).font = fontStyleBold
    worksheet.getCell(`A${startCellPlus11}`).font = fontStyleBold
    worksheet.getCell(`A${startCellPlus12}`).font = fontStyleBold

    worksheet.getCell(`M${startCellPlus8}`).font = fontStyleBold
    worksheet.getCell(`M${startCellPlus9}`).font = fontStyleBold
    worksheet.getCell(`M${startCellPlus10}`).font = fontStyleBold
    worksheet.getCell(`M${startCellPlus11}`).font = fontStyleBold
    worksheet.getCell(`M${startCellPlus12}`).font = fontStyleBold

    // alignment
    worksheet.getCell(`A${startCellPlus8}`).alignment = alignmentStyle
    worksheet.getCell(`A${startCellPlus9}`).alignment = alignmentStyle
    worksheet.getCell(`A${startCellPlus10}`).alignment = alignmentStyle
    worksheet.getCell(`A${startCellPlus11}`).alignment = alignmentStyle
    worksheet.getCell(`A${startCellPlus12}`).alignment = alignmentStyle

    worksheet.getCell(`M${startCellPlus8}`).alignment = alignmentStyle
    worksheet.getCell(`M${startCellPlus9}`).alignment = alignmentStyle
    worksheet.getCell(`M${startCellPlus10}`).alignment = alignmentStyle
    worksheet.getCell(`M${startCellPlus11}`).alignment = alignmentStyle
    worksheet.getCell(`M${startCellPlus12}`).alignment = alignmentStyle

    //
    worksheet.getCell(`A${startCellPlus8}`).value = "MENGETAHUI,"
    worksheet.getCell(`A${startCellPlus9}`).value = "KANTOR KESYAHBANDARAN DAN OTORITAS"
    worksheet.getCell(`A${startCellPlus10}`).value = "PELABUHAN KELAS III PULAU BAAI BENGKULU"
    worksheet.getCell(`A${startCellPlus11}`).value = "KEPALA SEKSI LALU LINTAS DAN UK KANTOR"
    worksheet.getCell(`A${startCellPlus12}`).value = "PELAKSANA HARIAN"

    worksheet.getCell(`M${startCellPlus8}`).value = "MENGETAHUI,"
    worksheet.getCell(`M${startCellPlus9}`).value = "PT PELABUHAN INDONESIA (PERSERO) REGIONAL 2 BENGKULU"
    worksheet.getCell(`M${startCellPlus10}`).value = "DGM KEUANGAN & SDM"
    worksheet.getCell(`M${startCellPlus11}`).value = ""
    worksheet.getCell(`M${startCellPlus12}`).value = ""

    const startCellPlus16 = startCell + 16
    const startCellPlus17 = startCell + 17

    //
    worksheet.mergeCells(`A${startCellPlus16}:E${startCellPlus16}`)
    worksheet.mergeCells(`A${startCellPlus17}:E${startCellPlus17}`)

    worksheet.mergeCells(`M${startCellPlus16}:Q${startCellPlus16}`)
    worksheet.mergeCells(`M${startCellPlus17}:Q${startCellPlus17}`)

    //
    worksheet.getCell(`A${startCellPlus16}`).font = {...fontStyleBold, underline: "single"}
    worksheet.getCell(`A${startCellPlus17}`).font = fontStyleBold

    worksheet.getCell(`M${startCellPlus16}`).font = {...fontStyleBold, underline: "single"}
    worksheet.getCell(`M${startCellPlus17}`).font = fontStyleBold

    //
    worksheet.getCell(`A${startCellPlus16}`).alignment = alignmentStyle
    worksheet.getCell(`A${startCellPlus17}`).alignment = alignmentStyle

    worksheet.getCell(`M${startCellPlus16}`).alignment = alignmentStyle
    worksheet.getCell(`M${startCellPlus17}`).alignment = alignmentStyle

    //
    worksheet.getCell(`A${startCellPlus16}`).value = "ALKISMAN, SE"
    worksheet.getCell(`A${startCellPlus17}`).value = "NIP. 19750727 201001 1 005"

    worksheet.getCell(`M${startCellPlus16}`).value = "SARIPUDIN"
    worksheet.getCell(`M${startCellPlus17}`).value = "NIPP. 101809"


    /**
     * SHEET 2
     */
    // title 1
    worksheet2.mergeCells("A1:H1");
    worksheet2.getCell("A1").value = "RINCIAN PERHITUNGAN PNBP JASA PEMANDUAN & PENUNDAAN";
    worksheet2.getCell("A1").font = fontStyleBold;
    worksheet2.getCell("A1").alignment = alignmentStyle;

    // title 2
    worksheet2.mergeCells("A2:H2");
    worksheet2.getCell("A2").value = "PT PELABUHAN INDONESIA ( PERSERO ) REGIONAL 2 BENGKULU";
    worksheet2.getCell("A2").font = fontStyleBold;
    worksheet2.getCell("A2").alignment = alignmentStyle;


    // title 3
    worksheet2.mergeCells("A3:H3");
    worksheet2.getCell("A3").value = "Periode Bulan Maret Tahun 2022";
    worksheet2.getCell("A3").font = fontStyleBold;
    worksheet2.getCell("A3").alignment = alignmentStyle;


    // styling
    worksheet2.mergeCells("C5:D5");

    const arrayHeaderBorderWS2 = [
        "A5",
        "B5",
        "C5",
        "E5",
        "F5",
        "G5",
        "H5",
        "I5"
    ]
    for (let i = 0; i < arrayHeaderBorderWS2.length; i++) {
        worksheet2.getCell(arrayHeaderBorder[i]).border = borderStyle
        worksheet2.getCell(arrayHeaderBorder[i]).alignment = alignmentStyle
        worksheet2.getCell(arrayHeaderBorder[i]).font = fontStyleBold
    }

    worksheet2.getCell("A5").value = "NO";
    worksheet2.getCell("B5").value = "Jasa Kepelabuhan";
    worksheet2.getCell("C5").value = "Segmen";
    worksheet2.getCell("E5").value = "Objek PNBP";
    worksheet2.getCell("F5").value = "Jumlah Pendapatan Kotor";
    worksheet2.getCell("G5").value = "Persentase PNBP";
    worksheet2.getCell("H5").value = "Jumlah Dibayar";

    let startCellWS2 = 6;
    // perbulan
    // const jumlahKapal = 1000
    // const gerakanKapal = 50

    const jumlahKapalWS2 = 10
    const gerakanKapalWS2 = 50

    for (let i = 0; i < (jumlahKapalWS2 * gerakanKapalWS2); i++) {
        worksheet2.getCell(`A${startCellWS2}`).value = `i + 1`;
        worksheet2.getCell(`B${startCellWS2}`).value = `i + 1`;
        worksheet2.getCell(`C${startCellWS2}`).value = `i + 1`;
        worksheet2.getCell(`D${startCellWS2}`).value = `i + 1`;
        worksheet2.getCell(`E${startCellWS2}`).value = `i + 1`;
        worksheet2.getCell(`F${startCellWS2}`).value = `i + 1`;
        worksheet2.getCell(`G${startCellWS2}`).value = `i + 1`;
        worksheet2.getCell(`H${startCellWS2}`).value = `i + 1`;

        // styling
        worksheet2.getCell(`A${startCellWS2}`).border = borderStyle;
        worksheet2.getCell(`B${startCellWS2}`).border = borderStyle;
        worksheet2.getCell(`C${startCellWS2}`).border = borderStyle;
        worksheet2.getCell(`D${startCellWS2}`).border = borderStyle;
        worksheet2.getCell(`E${startCellWS2}`).border = borderStyle;
        worksheet2.getCell(`F${startCellWS2}`).border = borderStyle;
        worksheet2.getCell(`G${startCellWS2}`).border = borderStyle;
        worksheet2.getCell(`H${startCellWS2}`).border = borderStyle;
        startCellWS2++;
    }


    // merge
    worksheet2.mergeCells(`A${startCellWS2}:E${startCellWS2}`)
    // style
    worksheet2.getCell(`A${startCellWS2}`).font = fontStyleBold
    worksheet2.getCell(`A${startCellWS2}`).alignment = alignmentStyle
    worksheet2.getCell(`H${startCellWS2}`).font = {italic: true, bold: true, underline: "doubleAccounting"}

    worksheet2.getCell(`A${startCellWS2}`).border = fontStyleBold
    worksheet2.getCell(`F${startCellWS2}`).border = fontStyleBold
    worksheet2.getCell(`H${startCellWS2}`).border = fontStyleBold

    // value
    worksheet2.getCell(`A${startCellWS2}`).value = "JUMLAH"
    worksheet2.getCell(`F${startCellWS2}`).value = "88888"
    worksheet2.getCell(`H${startCellWS2}`).value = "888888"

    const startCellWS2Plus4 = startCellWS2 + 4

    worksheet2.getCell(`B${startCellWS2Plus4}`).font = {italic: true, bold: true, underline: "singleAccounting"}
    //value
    worksheet2.getCell(`B${startCellWS2Plus4}`).value = "TERBILANG"
    worksheet2.getCell(`C${startCellWS2Plus4}`).value = ":"
    worksheet2.getCell(`D${startCellWS2Plus4}`).value = "JENIS PENDAPATAN"


    const startCellWS2Plus7 = startCell + 7
    const startCellWS2Plus8 = startCell + 8
    const startCellWS2Plus9 = startCell + 9
    const startCellWS2Plus10 = startCell + 10
    const startCellWS2Plus11 = startCell + 11

    // MERGE
    worksheet2.mergeCells(`B${startCellWS2Plus7}:D${startCellWS2Plus7}`)
    worksheet2.mergeCells(`B${startCellWS2Plus8}:D${startCellWS2Plus8}`)
    worksheet2.mergeCells(`B${startCellWS2Plus9}:D${startCellWS2Plus9}`)
    worksheet2.mergeCells(`B${startCellWS2Plus10}:D${startCellWS2Plus10}`)
    worksheet2.mergeCells(`B${startCellWS2Plus11}:D${startCellWS2Plus11}`)

    worksheet2.mergeCells(`F${startCellWS2Plus7}:H${startCellWS2Plus7}`)
    worksheet2.mergeCells(`F${startCellWS2Plus8}:H${startCellWS2Plus8}`)
    worksheet2.mergeCells(`F${startCellWS2Plus9}:H${startCellWS2Plus9}`)
    worksheet2.mergeCells(`F${startCellWS2Plus10}:H${startCellWS2Plus10}`)
    worksheet2.mergeCells(`F${startCellWS2Plus11}:H${startCellWS2Plus11}`)

    // STYLE
    worksheet2.getCell(`B${startCellWS2Plus7}`).font = fontStyleBold
    worksheet2.getCell(`B${startCellWS2Plus8}`).font = fontStyleBold
    worksheet2.getCell(`B${startCellWS2Plus9}`).font = fontStyleBold
    worksheet2.getCell(`B${startCellWS2Plus10}`).font = fontStyleBold
    worksheet2.getCell(`B${startCellWS2Plus11}`).font = fontStyleBold

    worksheet2.getCell(`F${startCellWS2Plus7}`).font = fontStyleBold
    worksheet2.getCell(`F${startCellWS2Plus8}`).font = fontStyleBold
    worksheet2.getCell(`F${startCellWS2Plus9}`).font = fontStyleBold
    worksheet2.getCell(`F${startCellWS2Plus10}`).font = fontStyleBold
    worksheet2.getCell(`F${startCellWS2Plus11}`).font = fontStyleBold

    // alignment
    worksheet2.getCell(`B${startCellWS2Plus7}`).alignment = alignmentStyle
    worksheet2.getCell(`B${startCellWS2Plus8}`).alignment = alignmentStyle
    worksheet2.getCell(`B${startCellWS2Plus9}`).alignment = alignmentStyle
    worksheet2.getCell(`B${startCellWS2Plus10}`).alignment = alignmentStyle
    worksheet2.getCell(`B${startCellWS2Plus11}`).alignment = alignmentStyle

    worksheet2.getCell(`F${startCellWS2Plus7}`).alignment = alignmentStyle
    worksheet2.getCell(`F${startCellWS2Plus8}`).alignment = alignmentStyle
    worksheet2.getCell(`F${startCellWS2Plus9}`).alignment = alignmentStyle
    worksheet2.getCell(`F${startCellWS2Plus10}`).alignment = alignmentStyle
    worksheet2.getCell(`F${startCellWS2Plus11}`).alignment = alignmentStyle

    //
    worksheet2.getCell(`B${startCellWS2Plus7}`).value = "MENGETAHUI,"
    worksheet2.getCell(`B${startCellWS2Plus8}`).value = "KANTOR KESYAHBANDARAN DAN OTORITAS"
    worksheet2.getCell(`B${startCellWS2Plus9}`).value = "PELABUHAN KELAS III PULAU BAAI BENGKULU"
    worksheet2.getCell(`B${startCellWS2Plus10}`).value = "KEPALA SEKSI LALU LINTAS DAN UK KANTOR"
    worksheet2.getCell(`B${startCellWS2Plus11}`).value = "PELAKSANA HARIAN"

    worksheet2.getCell(`F${startCellWS2Plus7}`).value = "MENGETAHUI,"
    worksheet2.getCell(`F${startCellWS2Plus8}`).value = "PT PELABUHAN INDONESIA (PERSERO) REGIONAL 2 BENGKULU"
    worksheet2.getCell(`F${startCellWS2Plus9}`).value = "DGM KEUANGAN & SDM"
    worksheet2.getCell(`F${startCellWS2Plus10}`).value = ""
    worksheet2.getCell(`F${startCellWS2Plus11}`).value = ""

    const startCellWS2Plus16 = startCellWS2 + 17
    const startCellWS2Plus17 = startCellWS2 + 18

    //
    worksheet2.mergeCells(`B${startCellWS2Plus16}:D${startCellWS2Plus16}`)
    worksheet2.mergeCells(`B${startCellWS2Plus17}:D${startCellWS2Plus17}`)

    worksheet2.mergeCells(`F${startCellWS2Plus16}:H${startCellWS2Plus16}`)
    worksheet2.mergeCells(`F${startCellWS2Plus17}:H${startCellWS2Plus17}`)

    //
    worksheet2.getCell(`B${startCellWS2Plus16}`).font = {...fontStyleBold, underline: "single"}
    worksheet2.getCell(`B${startCellWS2Plus17}`).font = fontStyleBold

    worksheet2.getCell(`F${startCellWS2Plus16}`).font = {...fontStyleBold, underline: "single"}
    worksheet2.getCell(`F${startCellWS2Plus17}`).font = fontStyleBold

    //
    worksheet2.getCell(`B${startCellWS2Plus16}`).alignment = alignmentStyle
    worksheet2.getCell(`B${startCellWS2Plus17}`).alignment = alignmentStyle

    worksheet2.getCell(`F${startCellWS2Plus16}`).alignment = alignmentStyle
    worksheet2.getCell(`F${startCellWS2Plus17}`).alignment = alignmentStyle

    //
    worksheet2.getCell(`B${startCellWS2Plus16}`).value = "ALKISMAN, SE"
    worksheet2.getCell(`B${startCellWS2Plus17}`).value = "NIP. 19750727 201001 1 005"

    worksheet2.getCell(`F${startCellWS2Plus16}`).value = "SARIPUDIN"
    worksheet2.getCell(`F${startCellWS2Plus17}`).value = "NIPP. 101809"

    workbook.xlsx.writeFile(
      "LAPORAN.xlsx"
    );
}

test()