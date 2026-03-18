import * as XLSX from "xlsx";

// ──────────────────────────────────────────────
// Type definitions
// ──────────────────────────────────────────────
export interface InstitutionRecord {
  code: string;
  name: string;
  postalCode: string;
  address: string;
  phone: string;
  category: string; // 病院 / 診療所 / 特定機能 / 地域支援
  status: string; // 現存 / 休止
  bedsGeneral: number | "";
  bedsPsychiatric: number | "";
  bedsNursing: number | "";
  bedsTuberculosis: number | "";
  fullTimeDoctors: number | "";
  partTimeDoctors: number | "";
  departments: string;
  founder: string;
  manager: string;
  designatedDate: string;
  renewalDate: string;
}

// Row type from sheet_to_json with header:1
type RawRow = (string | number | boolean | null | undefined)[];

// ──────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────

/** Replace full-width chars with half-width equivalents */
function normalizeFullWidth(str: string): string {
  if (!str) return "";
  // full-width digits / letters → half-width
  str = str.replace(/[\uFF01-\uFF5E]/g, (ch) =>
    String.fromCharCode(ch.charCodeAt(0) - 0xfee0)
  );
  // full-width hyphen/minus variants → hyphen
  str = str.replace(/[\u2212\uFF0D\u30FC\u2014\u2013]/g, "-");
  // full-width space → half-width
  str = str.replace(/\u3000/g, " ");
  return str.trim();
}

/** Get cell value as trimmed string */
function cellStr(row: RawRow | undefined, colIdx: number): string {
  if (!row) return "";
  const cell = row[colIdx];
  if (cell === undefined || cell === null) return "";
  const raw = String(cell);
  return raw.replace(/\u3000/g, " ").trim();
}

/** Parse medical institution code: remove commas */
function parseCode(raw: string): string {
  return raw.replace(/,/g, "");
}

/** Split postal code and address from combined D column */
function parseAddress(raw: string): { postalCode: string; address: string } {
  const normalized = normalizeFullWidth(raw);
  // Match 〒XXX-XXXX (with possible full-width chars already normalized)
  const match = normalized.match(/〒(\d{3}[-]\d{4})(.*)/);
  if (match) {
    return {
      postalCode: match[1],
      address: match[2].trim(),
    };
  }
  // Try without 〒
  const match2 = normalized.match(/(\d{3}[-]\d{4})(.*)/);
  if (match2) {
    return {
      postalCode: match2[1],
      address: match2[2].trim(),
    };
  }
  return { postalCode: "", address: normalized };
}

/** Convert Japanese era date string to YYYY/MM/DD */
function parseJapaneseDate(raw: string): string {
  if (!raw) return "";
  const normalized = normalizeFullWidth(raw);

  // Non-date values → empty
  const nonDate = ["組織変更", "移動", "新規", "交代", "その他"];
  if (nonDate.some((nd) => normalized.includes(nd))) return "";

  const eraMatch = normalized.match(
    /(昭|平|令|大)\s*(元|\d+)\.\s*(\d+)\.\s*(\d+)/
  );
  if (!eraMatch) return "";

  const [, era, yearStr, monthStr, dayStr] = eraMatch;
  let baseYear = 0;
  if (era === "昭") baseYear = 1925;
  else if (era === "平") baseYear = 1988;
  else if (era === "令") baseYear = 2018;
  else if (era === "大") baseYear = 1911;

  const eraYear = yearStr === "元" ? 1 : parseInt(yearStr, 10);
  const year = baseYear + eraYear;
  const month = parseInt(monthStr, 10);
  const day = parseInt(dayStr, 10);

  return `${year}/${String(month).padStart(2, "0")}/${String(day).padStart(2, "0")}`;
}

/** Extract bed count for a specific type from a string */
function extractBedCount(
  str: string,
  type: string
): number | "" {
  if (!str) return "";
  const normalized = normalizeFullWidth(str);
  const regex = new RegExp(type + "\\s*(\\d+)");
  const match = normalized.match(regex);
  if (match) return parseInt(match[1], 10);
  return "";
}

/** Extract all bed types from I column value */
function extractAllBeds(str: string): {
  general: number | "";
  psychiatric: number | "";
  nursing: number | "";
  tuberculosis: number | "";
} {
  const normalized = normalizeFullWidth(str);
  return {
    general: extractBedCount(normalized, "一般"),
    psychiatric: extractBedCount(normalized, "精神"),
    nursing: extractBedCount(normalized, "療養"),
    tuberculosis: extractBedCount(normalized, "結核"),
  };
}

/** Check if a string represents bed info */
function isBedInfo(str: string): boolean {
  if (!str) return false;
  const normalized = normalizeFullWidth(str);
  return /(一般|精神|療養|結核|感染)\s*\d+/.test(normalized);
}

/** Extract doctor counts from E column value */
function extractDoctors(str: string): {
  fullTime: number | "";
  partTime: number | "";
} {
  if (!str) return { fullTime: "", partTime: "" };
  const normalized = normalizeFullWidth(str);

  let fullTime: number | "" = "";
  let partTime: number | "" = "";

  // Part-time: contains 非常勤
  if (normalized.includes("非常勤")) {
    const m = normalized.match(/非常勤[^\d]*(\d+(?:\.\d+)?)/);
    if (m) partTime = parseFloat(m[1]);
  }
  // Full-time: contains 常 and 勤 but NOT 非
  else if (normalized.includes("常") && normalized.includes("勤") && !normalized.includes("非")) {
    const m = normalized.match(/常\s*勤[^\d]*(\d+(?:\.\d+)?)/);
    if (m) fullTime = parseFloat(m[1]);
  }

  return { fullTime, partTime };
}

// ──────────────────────────────────────────────
// Main processing function
// ──────────────────────────────────────────────

export function processExcelWithPreview(buffer: ArrayBuffer): { output: ArrayBuffer; records: InstitutionRecord[] } {
  const records = parseRecords(buffer);
  const output = buildExcel(records);
  return { output, records };
}

export function processExcel(buffer: ArrayBuffer): ArrayBuffer {
  return buildExcel(parseRecords(buffer));
}

function parseRecords(buffer: ArrayBuffer): InstitutionRecord[] {
  const workbook = XLSX.read(buffer, { type: "array", cellText: false });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Convert to array of arrays (raw values)
  const rawData: RawRow[] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    defval: undefined,
  }) as RawRow[];

  if (!rawData || rawData.length === 0) {
    throw new Error("Excelファイルにデータがありません。");
  }

  const records: InstitutionRecord[] = [];

  // Find first data row: A column is a number
  let startRow = -1;
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row) continue;
    const aVal = row[0];
    if (aVal !== undefined && aVal !== null && String(aVal) !== "") {
      const num = Number(aVal);
      if (!isNaN(num) && num > 0 && Number.isInteger(num)) {
        startRow = i;
        break;
      }
    }
  }

  if (startRow === -1) {
    throw new Error(
      "データの開始行が見つかりませんでした。A列に連番が含まれているか確認してください。"
    );
  }

  let i = startRow;
  while (i < rawData.length) {
    const firstRow = rawData[i];
    if (!firstRow) {
      i++;
      continue;
    }

    // Check if A column is a number (start of new institution)
    const aVal = firstRow[0];
    if (aVal === undefined || aVal === null || String(aVal) === "") {
      i++;
      continue;
    }
    const num = Number(aVal);
    if (isNaN(num) || num <= 0 || !Number.isInteger(num)) {
      i++;
      continue;
    }

    // ── First row parsing ──
    const rawCode = cellStr(firstRow, 1);
    const code = parseCode(rawCode);

    const name = cellStr(firstRow, 2);
    const rawAddress = cellStr(firstRow, 3);
    const { postalCode, address } = parseAddress(rawAddress);
    const phone = normalizeFullWidth(cellStr(firstRow, 4));
    const founder = cellStr(firstRow, 5).replace(/\u3000/g, " ").trim();
    const manager = cellStr(firstRow, 6).replace(/\u3000/g, " ").trim();
    const rawDesignatedDate = cellStr(firstRow, 7);
    const designatedDate = parseJapaneseDate(rawDesignatedDate);
    const firstRowBedStr = cellStr(firstRow, 8);
    const category = cellStr(firstRow, 9);

    // Beds from first row
    const firstBeds = extractAllBeds(firstRowBedStr);

    // ── Gather subsequent rows (until next institution or end) ──
    const subRows: RawRow[] = [];
    let j = i + 1;
    while (j < rawData.length) {
      const nextRow = rawData[j];
      if (!nextRow) {
        j++;
        continue;
      }
      // If A column has a positive integer, it's a new institution
      const nextA = nextRow[0];
      if (nextA !== undefined && nextA !== null && String(nextA) !== "") {
        const nextNum = Number(nextA);
        if (!isNaN(nextNum) && nextNum > 0 && Number.isInteger(nextNum)) {
          break;
        }
      }
      subRows.push(nextRow);
      j++;
    }

    // ── Process sub-rows ──
    let bedsGeneral: number | "" = firstBeds.general;
    let bedsPsychiatric: number | "" = firstBeds.psychiatric;
    let bedsNursing: number | "" = firstBeds.nursing;
    let bedsTuberculosis: number | "" = firstBeds.tuberculosis;
    let fullTimeDoctors: number | "" = "";
    let partTimeDoctors: number | "" = "";
    const departmentsList: string[] = [];
    let status = "";
    let renewalDate = "";

    for (const subRow of subRows) {
      // E column (index 4) - doctor counts
      const eVal = cellStr(subRow, 4);
      if (eVal) {
        const { fullTime, partTime } = extractDoctors(eVal);
        if (fullTime !== "") {
          if (fullTimeDoctors === "") fullTimeDoctors = 0;
          fullTimeDoctors = (fullTimeDoctors as number) + (fullTime as number);
        }
        if (partTime !== "") {
          if (partTimeDoctors === "") partTimeDoctors = 0;
          partTimeDoctors = (partTimeDoctors as number) + (partTime as number);
        }
      }

      // I column (index 8) - additional beds or departments
      const iVal = cellStr(subRow, 8);
      if (iVal) {
        const addBeds = extractAllBeds(iVal);
        if (addBeds.general !== "") {
          if (bedsGeneral === "") bedsGeneral = 0;
          bedsGeneral = (bedsGeneral as number) + (addBeds.general as number);
        }
        if (addBeds.psychiatric !== "") {
          if (bedsPsychiatric === "") bedsPsychiatric = 0;
          bedsPsychiatric =
            (bedsPsychiatric as number) + (addBeds.psychiatric as number);
        }
        if (addBeds.nursing !== "") {
          if (bedsNursing === "") bedsNursing = 0;
          bedsNursing = (bedsNursing as number) + (addBeds.nursing as number);
        }
        if (addBeds.tuberculosis !== "") {
          if (bedsTuberculosis === "") bedsTuberculosis = 0;
          bedsTuberculosis =
            (bedsTuberculosis as number) + (addBeds.tuberculosis as number);
        }

        // If not bed info, treat as department
        if (!isBedInfo(iVal)) {
          const normalized = normalizeFullWidth(iVal);
          if (normalized && !normalized.match(/^\s*$/)) {
            departmentsList.push(normalized);
          }
        }
      }

      // J column (index 9) - status
      const jVal = cellStr(subRow, 9);
      if (jVal) {
        if (jVal.includes("現存") || jVal.includes("休止")) {
          status = jVal.includes("休止") ? "休止" : "現存";
        }
      }

      // H column (index 7) - renewal date (rows 2+)
      const hVal = cellStr(subRow, 7);
      if (hVal) {
        const parsed = parseJapaneseDate(hVal);
        if (parsed) {
          renewalDate = parsed;
        }
      }
    }

    const record: InstitutionRecord = {
      code,
      name,
      postalCode,
      address,
      phone,
      category: category || "",
      status: status || "",
      bedsGeneral,
      bedsPsychiatric,
      bedsNursing,
      bedsTuberculosis,
      fullTimeDoctors,
      partTimeDoctors,
      departments: departmentsList.join(" / "),
      founder,
      manager,
      designatedDate,
      renewalDate,
    };

    records.push(record);
    i = j;
  }

  if (records.length === 0) {
    throw new Error("処理できるデータが見つかりませんでした。");
  }

  return records;
}

function buildExcel(records: InstitutionRecord[]): ArrayBuffer {
  // ──────────────────────────────────────────────
  // Build output Excel
  // ──────────────────────────────────────────────
  const headers = [
    "医療機関コード",
    "医療機関名",
    "郵便番号",
    "所在地",
    "電話番号",
    "種別",
    "状態",
    "一般病床",
    "精神病床",
    "療養病床",
    "結核病床",
    "常勤医数",
    "非常勤医数",
    "診療科目",
    "開設者",
    "管理者",
    "指定年月日",
    "指定更新日",
  ];

  const rows = records.map((r) => [
    r.code,
    r.name,
    r.postalCode,
    r.address,
    r.phone,
    r.category,
    r.status,
    r.bedsGeneral,
    r.bedsPsychiatric,
    r.bedsNursing,
    r.bedsTuberculosis,
    r.fullTimeDoctors,
    r.partTimeDoctors,
    r.departments,
    r.founder,
    r.manager,
    r.designatedDate,
    r.renewalDate,
  ]);

  const wsData = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Column widths
  const colWidths = [
    { wch: 14 }, // 医療機関コード
    { wch: 30 }, // 医療機関名
    { wch: 12 }, // 郵便番号
    { wch: 40 }, // 所在地
    { wch: 16 }, // 電話番号
    { wch: 12 }, // 種別
    { wch: 8 },  // 状態
    { wch: 8 },  // 一般病床
    { wch: 8 },  // 精神病床
    { wch: 8 },  // 療養病床
    { wch: 8 },  // 結核病床
    { wch: 8 },  // 常勤医数
    { wch: 10 }, // 非常勤医数
    { wch: 40 }, // 診療科目
    { wch: 24 }, // 開設者
    { wch: 16 }, // 管理者
    { wch: 14 }, // 指定年月日
    { wch: 14 }, // 指定更新日
  ];
  ws["!cols"] = colWidths;

  // Freeze header row
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };

  // Apply styles
  const headerStyle = {
    fill: { fgColor: { rgb: "1E3A5F" } },
    font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: {
      bottom: { style: "thin", color: { rgb: "FFFFFF" } },
    },
  };

  const rowStyleEven = {
    fill: { fgColor: { rgb: "EBF3FB" } },
    font: { sz: 10 },
    alignment: { vertical: "center" },
  };

  const rowStyleOdd = {
    fill: { fgColor: { rgb: "FFFFFF" } },
    font: { sz: 10 },
    alignment: { vertical: "center" },
  };

  const totalCols = headers.length;
  const totalRows = wsData.length;

  for (let row = 0; row < totalRows; row++) {
    for (let col = 0; col < totalCols; col++) {
      const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
      if (!ws[cellRef]) {
        ws[cellRef] = { v: "", t: "s" };
      }
      if (row === 0) {
        ws[cellRef].s = headerStyle;
      } else {
        ws[cellRef].s = row % 2 === 0 ? rowStyleEven : rowStyleOdd;
      }
    }
  }

  // Auto-filter on header row
  ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 0, c: totalCols - 1 } }) };

  // Row heights
  const rowHeights: XLSX.RowInfo[] = [{ hpt: 28 }]; // header height
  for (let r = 1; r < totalRows; r++) {
    rowHeights.push({ hpt: 20 });
  }
  ws["!rows"] = rowHeights;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "医療機関一覧");

  const output = XLSX.write(wb, {
    bookType: "xlsx",
    type: "array",
    cellStyles: true,
  });

  return output as ArrayBuffer;
}
