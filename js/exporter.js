const CURRENCY_FMT = "#,##0.0000";
const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } };
const TOTAL_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
const THIN_BORDER = {
  top: { style: "thin", color: { argb: "FFB4B4B4" } },
  left: { style: "thin", color: { argb: "FFB4B4B4" } },
  bottom: { style: "thin", color: { argb: "FFB4B4B4" } },
  right: { style: "thin", color: { argb: "FFB4B4B4" } },
};

function colLetter(col) {
  let letter = "";
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

function styleHeaderRow(worksheet, row, startCol, endCol) {
  for (let col = startCol; col <= endCol; col++) {
    const cell = worksheet.getCell(row, col);
    cell.fill = HEADER_FILL;
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { horizontal: "center", wrapText: true };
  }
}

function applyBorder(worksheet, row, startCol, endCol) {
  for (let col = startCol; col <= endCol; col++) {
    worksheet.getCell(row, col).border = THIN_BORDER;
  }
}

function setCurrency(worksheet, row, cols) {
  for (const col of cols) {
    worksheet.getCell(row, col).numFmt = CURRENCY_FMT;
  }
}

function writeSummaryPanel(worksheet, summary, sectionTotalRows, startCol, startRow) {
  const valueCol = startCol + 1;
  const valueLetter = colLetter(valueCol);
  const summaryRowMap = {};

  worksheet.mergeCells(startRow, startCol, startRow, valueCol);
  const titleCell = worksheet.getCell(startRow, startCol);
  titleCell.value = "COST SUMMARY";
  titleCell.font = { bold: true, size: 12 };

  const headerRow = startRow + 1;
  worksheet.getCell(headerRow, startCol).value = "Line Item";
  worksheet.getCell(headerRow, valueCol).value = "Amount";
  styleHeaderRow(worksheet, headerRow, startCol, valueCol);
  applyBorder(worksheet, headerRow, startCol, valueCol);

  const materialFormula = sectionTotalRows.length
    ? `SUM(${sectionTotalRows.map((r) => `G${r}`).join(",")})`
    : "0";
  const wasteFormula = sectionTotalRows.length
    ? `SUM(${sectionTotalRows.map((r) => `H${r}`).join(",")})`
    : "0";

  let sumRow = headerRow + 1;

  for (const entry of summary) {
    const labelCell = worksheet.getCell(sumRow, startCol);
    const valueCell = worksheet.getCell(sumRow, valueCol);
    labelCell.value = entry.label;

    switch (entry.kind) {
      case "material":
        valueCell.value = { formula: materialFormula };
        break;
      case "waste":
        valueCell.value = { formula: wasteFormula };
        break;
      case "sum_misc_pallet":
        valueCell.value = {
          formula: `${valueLetter}${summaryRowMap["MISC MATERIAL"]}+${valueLetter}${summaryRowMap["PALLET MATERIAL"]}`,
        };
        break;
      case "sum_material_block":
        valueCell.value = {
          formula: `${["MATERIAL", "VEND-FRT & DISC", "TOTAL MISC & PALLET", "WASTE"]
            .map((key) => `${valueLetter}${summaryRowMap[key]}`)
            .join("+")}`,
        };
        break;
      case "sum_labor":
        valueCell.value = {
          formula: `${["LABOR FROM ROUTING", "RECEIVING", "WAREHOUSING"]
            .map((key) => `${valueLetter}${summaryRowMap[key]}`)
            .join("+")}`,
        };
        break;
      case "sum_labor_block":
        valueCell.value = {
          formula: `${valueLetter}${summaryRowMap["TOTAL LABOR"]}+${valueLetter}${summaryRowMap["TAXES & BENEFITS"]}`,
        };
        break;
      case "sum_inventory":
        valueCell.value = {
          formula: `${valueLetter}${summaryRowMap["TOTAL MATERIAL"]}+${valueLetter}${summaryRowMap["TOTAL LABOR, TAXES & BENEFITS"]}`,
        };
        break;
      case "total_cost":
        valueCell.value = {
          formula: `${valueLetter}${summaryRowMap["TOTAL INVENTORY COST"]}+${valueLetter}${summaryRowMap.FREIGHT}+${valueLetter}${summaryRowMap["SALES ALLOWANCE"]}`,
        };
        labelCell.font = { bold: true, size: 12 };
        valueCell.font = { bold: true, size: 12 };
        labelCell.fill = TOTAL_FILL;
        valueCell.fill = TOTAL_FILL;
        break;
      default:
        valueCell.value = entry.value;
    }

    setCurrency(worksheet, sumRow, [valueCol]);
    applyBorder(worksheet, sumRow, startCol, valueCol);
    summaryRowMap[entry.label] = sumRow;
    sumRow += 1;
  }
}

export async function buildWorkbook(bom, ExcelJS) {
  const { meta, sections, summary } = bom;
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("BOM", {
    views: [{ state: "frozen", ySplit: 5, xSplit: 0 }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1 },
  });

  worksheet.mergeCells("A1:H1");
  worksheet.getCell("A1").value = "COSTED BILL OF MATERIAL";
  worksheet.getCell("A1").font = { bold: true, size: 14 };

  const metaPairs = [
    ["Item No", meta.itemNo],
    ["Entity", meta.entity],
    ["Warehouse", meta.warehouse],
    ["Generic", meta.generic],
    ["Date", meta.date],
  ];

  let metaCol = 1;
  for (const [label, value] of metaPairs) {
    worksheet.getCell(2, metaCol).value = label;
    worksheet.getCell(2, metaCol).font = { bold: true };
    worksheet.getCell(2, metaCol + 1).value = cleanText(value);
    metaCol += 2;
  }

  let headerRow = 4;
  if (meta.description) {
    worksheet.getCell(3, 1).value = "Description";
    worksheet.getCell(3, 1).font = { bold: true };
    worksheet.mergeCells("B3:H3");
    worksheet.getCell("B3").value = cleanText(meta.description);
    worksheet.getCell("B3").alignment = { wrapText: true, vertical: "top" };
    headerRow = 5;
  }

  const headers = [
    "Category",
    "Part #",
    "Description",
    "Qty",
    "UOM",
    "Unit Cost",
    "Extended Cost",
    "Waste",
  ];

  headers.forEach((header, index) => {
    worksheet.getCell(headerRow, index + 1).value = header;
  });
  styleHeaderRow(worksheet, headerRow, 1, headers.length);

  let row = headerRow + 1;
  const sectionTotalRows = [];

  for (const section of sections) {
    if (!section.items.length) continue;

    const itemRows = [];
    for (const item of section.items) {
      if (item.type === "section_total") continue;

      worksheet.getCell(row, 1).value = section.name;
      worksheet.getCell(row, 2).value = item.part;
      worksheet.getCell(row, 3).value = item.description;
      worksheet.getCell(row, 4).value = item.qty;
      worksheet.getCell(row, 5).value = item.uom;
      if (item.currPer != null) worksheet.getCell(row, 6).value = item.currPer;
      worksheet.getCell(row, 7).value = { formula: `D${row}*F${row}` };
      if (item.currWaste != null) worksheet.getCell(row, 8).value = item.currWaste;

      setCurrency(worksheet, row, [4, 6, 7, 8]);
      applyBorder(worksheet, row, 1, headers.length);
      itemRows.push(row);
      row += 1;
    }

    if (itemRows.length) {
      const first = Math.min(...itemRows);
      const last = Math.max(...itemRows);
      worksheet.getCell(row, 1).value = section.name;
      worksheet.getCell(row, 1).font = { bold: true };
      worksheet.getCell(row, 2).value = `TOTAL ${section.name}`;
      worksheet.mergeCells(row, 2, row, 6);
      worksheet.getCell(row, 7).value = { formula: `SUM(G${first}:G${last})` };
      worksheet.getCell(row, 8).value = { formula: `SUM(H${first}:H${last})` };
      setCurrency(worksheet, row, [7, 8]);
      applyBorder(worksheet, row, 1, headers.length);
      sectionTotalRows.push(row);
      row += 1;
    }
  }

  writeSummaryPanel(worksheet, summary, sectionTotalRows, 10, headerRow);

  worksheet.getColumn(1).width = 12;
  worksheet.getColumn(2).width = 10;
  worksheet.getColumn(3).width = 42;
  worksheet.getColumn(4).width = 10;
  worksheet.getColumn(5).width = 6;
  worksheet.getColumn(6).width = 11;
  worksheet.getColumn(7).width = 13;
  worksheet.getColumn(8).width = 10;
  worksheet.getColumn(9).width = 2;
  worksheet.getColumn(10).width = 28;
  worksheet.getColumn(11).width = 14;

  return workbook;
}

function cleanText(value) {
  return String(value || "").replace(/\(cid:\d+\)/g, "").trim();
}

export async function downloadExcel(bom, ExcelJS, filename) {
  const workbook = await buildWorkbook(bom, ExcelJS);
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
