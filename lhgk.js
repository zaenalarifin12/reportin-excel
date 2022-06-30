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
    const worksheet = workbook.addWorksheet("LHGK BENGKULU MAR-2022");

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


    const arrayHeaderBorder = []
    for (let i = 0; i < arrayHeaderBorder.length; i++) {
        worksheet.getCell(arrayHeaderBorder[i]).border = borderStyle
        worksheet.getCell(arrayHeaderBorder[i]).alignment = alignmentStyle
        worksheet.getCell(arrayHeaderBorder[i]).font = fontStyleBold
    }

    worksheet.getCell("A1").value = "NM_CABANG";
    worksheet.getCell("B1").value = "KD_PPKB";
    worksheet.getCell("C1").value = "PERIODE";
    worksheet.getCell("D1").value = "NO_UKK";
    worksheet.getCell("E1").value = "NO_BKT_PANDU";
    worksheet.getCell("F1").value = "TGL_JAM_TIBA";
    worksheet.getCell("G1").value = "PPKB_KE";
    worksheet.getCell("H1").value = "DRAFT_DEPAN";
    worksheet.getCell("I1").value = "DRAFT_BELAKANG";
    worksheet.getCell("J1").value = "NM_TUGBOAT1";
    worksheet.getCell("K1").value = "KP_GRT_TUGBOAT1";
    worksheet.getCell("L1").value = "KP_LOA_TUGBOAT1";
    worksheet.getCell("M1").value = "FLAG_TUGBOAT1";
    worksheet.getCell("N1").value = "DRAFT_DEPAN_TUGBOAT1";
    worksheet.getCell("O1").value = "DRAFT_BELAKANG_TUGBOAT1";
    worksheet.getCell("P1").value = "NM_TUGBOAT2";
    worksheet.getCell("Q1").value = "KP_GRT_TUGBOAT2";
    worksheet.getCell("R1").value = "KP_LOA_TUGBOAT2";
    worksheet.getCell("S1").value = "FLAG_TUGBOAT2";
    worksheet.getCell("T1").value = "DRAFT_DEPAN_TUGBOAT2";
    worksheet.getCell("U1").value = "DRAFT_BELAKANG_TUGBOAT2";
    worksheet.getCell("V1").value = "NM_KAPAL";
    worksheet.getCell("W1").value = "JN_KAPAL";
    worksheet.getCell("X1").value = "KP_GRT";
    worksheet.getCell("Y1").value = "KP_LOA";
    worksheet.getCell("Z1").value = "KD_BENDERA";

    worksheet.getCell(`AA1`).value = `KD_AGEN`;
    worksheet.getCell(`AB1`).value = `NM_PERS_PANDU`;
    worksheet.getCell(`AC1`).value = `TGL_TIBA`;
    worksheet.getCell(`AD1`).value = `JAM_TIBA`;
    worksheet.getCell(`AE1`).value = `TGL_PMT`;
    worksheet.getCell(`AF1`).value = `JAM_PMT`;
    worksheet.getCell(`AG1`).value = `PNK`;
    worksheet.getCell(`AH1`).value = `KB`;
    worksheet.getCell(`AI1`).value = `MULAI_PELAKSANAAN`;
    worksheet.getCell(`AJ1`).value = `SELESAI_PELAKSANAAN`;
    worksheet.getCell(`AK1`).value = `PND`;
    worksheet.getCell(`AL1`).value = `WT`;
    worksheet.getCell(`AM1`).value = `PT`;
    worksheet.getCell(`AN1`).value = `TRT`;
    worksheet.getCell(`AO1`).value = `AT_JAM`;
    worksheet.getCell(`AP1`).value = `PANDU_DARI`;
    worksheet.getCell(`AQ1`).value = `PANDU_KE`;
    worksheet.getCell(`AR1`).value = `KD_GERAKAN`;
    worksheet.getCell(`AS1`).value = `GERAKAN`;
    worksheet.getCell(`AT1`).value = `TGL_MPANDU`;
    worksheet.getCell(`AU1`).value = `NM_AGEN`;
    worksheet.getCell(`AV1`).value = `NM_KAPAL_1`;
    worksheet.getCell(`AW1`).value = `NM_KAPAL_2`;
    worksheet.getCell(`AX1`).value = `NM_KAPAL_3`;
    worksheet.getCell(`AY1`).value = `NM_KAPAL_4`;
    worksheet.getCell(`AZ1`).value = `NM_KAPAL_5`;

    worksheet.getCell(`BA1`).value = `NM_KAPAL_6`;
    worksheet.getCell(`BB1`).value = `MULAI_TUNDA`;
    worksheet.getCell(`BC1`).value = `SELESAI_TUNDA`;
    worksheet.getCell(`BD1`).value = `LAMA_TUNDA`;
    worksheet.getCell(`BE1`).value = `KET_PANDU`;
    worksheet.getCell(`BF1`).value = `KET_PANDU`;
    worksheet.getCell(`BG1`).value = `KETERANGAN_PANDU`;
    worksheet.getCell(`BH1`).value = `PELAYARAN`;
    worksheet.getCell(`BI1`).value = `PENDAPATAN_PANDU`;
    worksheet.getCell(`BJ1`).value = `PENDAPATAN_TUNDA`;
    worksheet.getCell(`BK1`).value = `PNBP_PANDU`;
    worksheet.getCell(`BL1`).value = `PNBP_TUNDA`;
    worksheet.getCell(`BM1`).value = `JUMLAH_PNBP`;

    let startCell = 2;

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

        worksheet.getCell(`AA${startCell}`).value = `i + 1`;
        worksheet.getCell(`AB${startCell}`).value = `i + 1`;
        worksheet.getCell(`AC${startCell}`).value = `i + 1`;
        worksheet.getCell(`AD${startCell}`).value = `i + 1`;
        worksheet.getCell(`AE${startCell}`).value = `i + 1`;
        worksheet.getCell(`AF${startCell}`).value = `i + 1`;
        worksheet.getCell(`AG${startCell}`).value = `i + 1`;
        worksheet.getCell(`AH${startCell}`).value = `i + 1`;
        worksheet.getCell(`AI${startCell}`).value = `i + 1`;
        worksheet.getCell(`AJ${startCell}`).value = `i + 1`;
        worksheet.getCell(`AK${startCell}`).value = `i + 1`;
        worksheet.getCell(`AL${startCell}`).value = `i + 1`;
        worksheet.getCell(`AM${startCell}`).value = `i + 1`;
        worksheet.getCell(`AN${startCell}`).value = `i + 1`;
        worksheet.getCell(`AO${startCell}`).value = `i + 1`;
        worksheet.getCell(`AP${startCell}`).value = `i + 1`;
        worksheet.getCell(`AQ${startCell}`).value = `i + 1`;
        worksheet.getCell(`AR${startCell}`).value = `i + 1`;
        worksheet.getCell(`AS${startCell}`).value = `i + 1`;
        worksheet.getCell(`AT${startCell}`).value = `i + 1`;
        worksheet.getCell(`AU${startCell}`).value = `i + 1`;
        worksheet.getCell(`AV${startCell}`).value = `i + 1`;
        worksheet.getCell(`AW${startCell}`).value = `i + 1`;
        worksheet.getCell(`AX${startCell}`).value = `i + 1`;
        worksheet.getCell(`AY${startCell}`).value = `i + 1`;
        worksheet.getCell(`AZ${startCell}`).value = `i + 1`;

        worksheet.getCell(`BA${startCell}`).value = `i + 1`;
        worksheet.getCell(`BB${startCell}`).value = `i + 1`;
        worksheet.getCell(`BC${startCell}`).value = `i + 1`;
        worksheet.getCell(`BD${startCell}`).value = `i + 1`;
        worksheet.getCell(`BE${startCell}`).value = `i + 1`;
        worksheet.getCell(`BF${startCell}`).value = `i + 1`;
        worksheet.getCell(`BG${startCell}`).value = `i + 1`;
        worksheet.getCell(`BH${startCell}`).value = `i + 1`;
        worksheet.getCell(`BI${startCell}`).value = `i + 1`;
        worksheet.getCell(`BJ${startCell}`).value = `i + 1`;
        worksheet.getCell(`BK${startCell}`).value = `i + 1`;
        worksheet.getCell(`BL${startCell}`).value = `i + 1`;
        worksheet.getCell(`BM${startCell}`).value = `i + 1`;
        // styling
        // worksheet.getCell(`A${startCell}`).border = borderStyle;
        // worksheet.getCell(`B${startCell}`).border = borderStyle;
        // worksheet.getCell(`C${startCell}`).border = borderStyle;
        // worksheet.getCell(`D${startCell}`).border = borderStyle;
        // worksheet.getCell(`E${startCell}`).border = borderStyle;
        // worksheet.getCell(`F${startCell}`).border = borderStyle;
        // worksheet.getCell(`G${startCell}`).border = borderStyle;
        // worksheet.getCell(`H${startCell}`).border = borderStyle;
        // worksheet.getCell(`I${startCell}`).border = borderStyle;
        // worksheet.getCell(`J${startCell}`).border = borderStyle;
        // worksheet.getCell(`K${startCell}`).border = borderStyle;
        // worksheet.getCell(`L${startCell}`).border = borderStyle;
        // worksheet.getCell(`M${startCell}`).border = borderStyle;
        // worksheet.getCell(`N${startCell}`).border = borderStyle;
        // worksheet.getCell(`O${startCell}`).border = borderStyle;
        // worksheet.getCell(`P${startCell}`).border = borderStyle;
        // worksheet.getCell(`Q${startCell}`).border = borderStyle;
        startCell++;
    }

    workbook.xlsx.writeFile(
      "LAPORAN-LHGK.xlsx"
    );
}

test()