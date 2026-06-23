const UOM_PATTERN =
  /^(EA|M3|GA|RL|FT|PC|BOOT|PCS|LB|KG|IN|SF|SY|LF|CY|TON|GAL|QT|PT|OZ|ML|L|MM|CM|M)$/;
const NUM_OR_STAR = /^(\d+\.?\d*|\*{8})$/;

export function cleanText(value) {
  return String(value || "").replace(/\(cid:\d+\)/g, "").trim();
}

function parseValue(token) {
  if (token === "********") return null;
  return parseFloat(token);
}

export function parseLineItem(line) {
  line = line.trim();
  if (!line || line.startsWith("Desc:") || line.startsWith("--------")) {
    return null;
  }

  if (line.startsWith("TOTAL ")) {
    const parts = line.split(/\s+/);
    if (
      parts.length >= 5 &&
      NUM_OR_STAR.test(parts[parts.length - 4]) &&
      NUM_OR_STAR.test(parts[parts.length - 3])
    ) {
      return {
        type: "section_total",
        section: parts.slice(1, -4).join(" "),
        currCost: parseValue(parts[parts.length - 4]),
        currWaste: parseValue(parts[parts.length - 3]),
      };
    }
    return null;
  }

  if (!/^\d{5}\s/.test(line)) return null;

  const tokens = line.split(/\s+/);
  if (tokens.length < 9) return null;

  const tail = tokens.slice(-8);
  if (!tail.slice(2).every((t) => NUM_OR_STAR.test(t))) return null;
  if (!UOM_PATTERN.test(tail[1])) return null;

  const head = tokens.slice(0, -8);
  return {
    type: "item",
    part: head[0],
    description: head.slice(1).join(" "),
    qty: parseFloat(tail[0]),
    uom: tail[1],
    currPer: parseValue(tail[2]),
    currCost: parseValue(tail[3]),
    currWaste: parseValue(tail[4]),
  };
}

export function extractMetadata(text) {
  const meta = {};
  const patterns = [
    ["itemNo", /ITEM NO\s*:\s*(\S+)/i],
    ["entity", /ENTITY\s*:\s*(\S+)/i],
    ["warehouse", /WAREHOUSE\s*:\s*(\S+)/i],
    ["generic", /Generic\s*:\s*(\S+)/i],
    ["date", /(\d{1,2}\/\d{1,2}\/\d{2})\s+\d{2}:\d{2}/],
  ];

  for (const [key, pattern] of patterns) {
    const match = text.match(pattern);
    if (match) meta[key] = cleanText(match[1]);
  }

  const descLines = [];
  const lines = text.split(/\r?\n/);
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (line.startsWith("Desc:")) {
      descLines.push(line.slice(5).trim());
    } else if (descLines.length && !/^(CHEM|HDWR|\d{5})/.test(line)) {
      if (line && !line.startsWith("PART#")) descLines.push(line);
      else break;
    } else if (descLines.length) {
      break;
    }
  }
  if (descLines.length) meta.description = descLines.join(" ");

  return meta;
}

const SUMMARY_PATTERNS = [
  [/^MATERIAL\s+([\d.]+)/, "MATERIAL", "material"],
  [/^VEND-FRT & DISC\s+([\d.]+)/, "VEND-FRT & DISC", "editable"],
  [/^MISC MATERIAL\s+([\d.]+)/, "MISC MATERIAL", "editable"],
  [/^PALLET MATERIAL\s+([\d.]+)/, "PALLET MATERIAL", "editable"],
  [/^TOTAL MISC & PALLET\s+([\d.]+)/, "TOTAL MISC & PALLET", "sum_misc_pallet"],
  [/^WASTE\s+([\d.]+)/, "WASTE", "waste"],
  [/^TOTAL MATERIAL\([^)]*\)\s+([\d.]+)/, "TOTAL MATERIAL", "sum_material_block"],
  [/^LABOR FROM ROUTINGS?\s+([\d.]+)/, "LABOR FROM ROUTING", "editable"],
  [/^RECEIVING\([^)]*\)\s+([\d.]+)/, "RECEIVING", "editable"],
  [/^WAREHOUSING\([^)]*\)\s+([\d.]+)/, "WAREHOUSING", "editable"],
  [/^TOTAL LABOR\([^)]*\)\s+([\d.]+)/, "TOTAL LABOR", "sum_labor"],
  [/^TAXES & BENEFITS\([^)]*\)\s+([\d.]+)/, "TAXES & BENEFITS", "editable"],
  [
    /^TOTAL LABOR, TAXES & BENEFITS\s+([\d.]+)/,
    "TOTAL LABOR, TAXES & BENEFITS",
    "sum_labor_block",
  ],
  [/^TOTAL INVENTORY COST\s+([\d.]+)/, "TOTAL INVENTORY COST", "sum_inventory"],
  [/^FREIGHT\([^)]*\)\s+([\d.]+)/, "FREIGHT", "editable"],
  [/^SALES ALLOWANCE\([^)]*\)\s+([\d.]+)/, "SALES ALLOWANCE", "editable"],
  [/^TOTAL COST\s+([\d.]+)/, "TOTAL COST", "total_cost"],
];

export function parseSummary(text) {
  const match = text.match(/Totals for Finished Item:[\s\S]*/);
  const block = match ? match[0] : text;
  const rows = [];

  for (const rawLine of block.split(/\r?\n/)) {
    const line = cleanText(rawLine);
    for (const [pattern, label, kind] of SUMMARY_PATTERNS) {
      const m = line.match(pattern);
      if (m) {
        rows.push({ label, kind, value: parseFloat(m[1]) });
        break;
      }
    }
  }

  return rows;
}

export function parseBomFromText(fullText, pageTexts) {
  const sections = [];
  let currentSection = null;

  for (const pageText of pageTexts) {
    for (const rawLine of pageText.split(/\r?\n/)) {
      const line = cleanText(rawLine);
      if (!line) continue;

      if (/^[A-Z][A-Z0-9/-]*$/.test(line) && !line.startsWith("Desc")) {
        if (line !== "PAGE" && line.length < 20) {
          currentSection = line;
          sections.push({ name: currentSection, items: [] });
        }
        continue;
      }

      const parsed = parseLineItem(line);
      if (parsed && currentSection) {
        sections[sections.length - 1].items.push(parsed);
      }
    }
  }

  return {
    meta: extractMetadata(fullText),
    sections,
    summary: parseSummary(fullText),
  };
}

function pageHasLineItems(pageText) {
  return /^\d{5}\s/m.test(pageText);
}

function pageStartsBom(pageText) {
  return (
    pageText.includes("COSTED BILL OF MATERIAL") &&
    pageText.includes("PART#") &&
    pageHasLineItems(pageText)
  );
}

export function splitIntoBomBlocks(pageTexts) {
  const blocks = [];
  let itemPages = [];
  let bomTextParts = [];

  const pushBlock = (summaryText) => {
    if (!itemPages.length) return;

    const fullText = [...bomTextParts, summaryText].filter(Boolean).join("\n");
    const totalsMatch = summaryText?.match(/Totals for Finished Item:\s*(\S+)/);
    const itemNo =
      cleanText(totalsMatch?.[1]) ||
      cleanText(fullText.match(/ITEM NO\s*:\s*(\S+)/i)?.[1]) ||
      `BOM ${blocks.length + 1}`;

    blocks.push({
      itemNo,
      itemPages: [...itemPages],
      summaryText,
      fullText,
    });
    itemPages = [];
    bomTextParts = [];
  };

  for (const pageText of pageTexts) {
    const totalsMatch = pageText.match(/Totals for Finished Item:\s*(\S+)/);
    const hasLineItems = pageHasLineItems(pageText);

    if (pageStartsBom(pageText) && itemPages.length > 0) {
      pushBlock(null);
    }

    if (hasLineItems) {
      itemPages.push(pageText);
      bomTextParts.push(pageText);
    } else if (
      pageText.includes("COSTED BILL OF MATERIAL") &&
      pageText.includes("ITEM NO")
    ) {
      bomTextParts.push(pageText);
    }

    if (totalsMatch) {
      pushBlock(pageText);
    }
  }

  if (itemPages.length > 0) {
    pushBlock(null);
  }

  return blocks;
}

export function parseBomReport(fullText, pageTexts) {
  const blocks = splitIntoBomBlocks(pageTexts);

  if (!blocks.length) {
    const fallback = parseBomFromText(fullText, pageTexts.slice(0, 2));
    return {
      boms: [fallback],
      bomCount: 1,
      itemCount: countItems(fallback),
    };
  }

  const boms = blocks.map((block) =>
    parseBomFromText(
      block.fullText || block.itemPages.join("\n"),
      block.itemPages,
    ),
  );

  for (let index = 0; index < boms.length; index++) {
    boms[index].meta.itemNo = cleanText(
      blocks[index].itemNo || boms[index].meta.itemNo,
    );
    boms[index].sheetName = boms[index].meta.itemNo;
  }

  const itemCount = boms.reduce((sum, bom) => sum + countItems(bom), 0);

  return { boms, bomCount: boms.length, itemCount };
}

function countItems(bom) {
  return bom.sections.reduce(
    (sum, section) =>
      sum + section.items.filter((item) => item.type === "item").length,
    0,
  );
}

function groupTextIntoLines(items) {
  const tolerance = 3;
  const lines = [];

  const sorted = items
    .filter((item) => item.str && item.str.trim())
    .map((item) => ({
      x: item.transform[4],
      y: item.transform[5],
      str: item.str,
    }))
    .sort((a, b) => b.y - a.y || a.x - b.x);

  for (const item of sorted) {
    let line = lines.find((entry) => Math.abs(entry.y - item.y) <= tolerance);
    if (!line) {
      line = { y: item.y, parts: [] };
      lines.push(line);
    }
    line.parts.push(item);
  }

  return lines
    .sort((a, b) => b.y - a.y)
    .map((line) =>
      line.parts
        .sort((a, b) => a.x - b.x)
        .map((part) => part.str)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim(),
    )
    .filter(Boolean);
}

export async function extractPdfText(file, pdfjsLib) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const pageTexts = [];
  let fullText = "";

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const content = await page.getTextContent();
    const text = groupTextIntoLines(content.items).join("\n");
    pageTexts.push(text);
    fullText += `${text}\n`;
  }

  return { fullText, pageTexts, pageCount: pdf.numPages };
}

export async function parseBomPdf(file, pdfjsLib) {
  const { fullText, pageTexts, pageCount } = await extractPdfText(file, pdfjsLib);
  const report = parseBomReport(fullText, pageTexts);

  if (report.itemCount === 0) {
    throw new Error(
      "No line items found. Make sure this is a costed BOM PDF with part numbers and quantities.",
    );
  }

  return { ...report, pageCount };
}
