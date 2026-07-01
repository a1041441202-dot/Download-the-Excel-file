import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { bitable } from "@lark-base-open/js-sdk";
import type { IOpenAttachment } from "@lark-base-open/js-sdk";

export interface SelectOption {
  id: string;
  name: string;
}

export interface FieldMeta {
  id: string;
  name: string;
  type: number;
  options?: SelectOption[];
}

export interface FilterCondition {
  fieldId: string;
  fieldName: string;
  operator: string;
  value: string;
}

const OPERATOR_MAP: Record<string, (cv: string, fv: string) => boolean> = {
  equals: (cv, fv) => cv === fv,
  not_equals: (cv, fv) => cv !== fv,
  contains: (cv, fv) => cv.includes(fv),
  not_contains: (cv, fv) => !cv.includes(fv),
  starts_with: (cv, fv) => cv.startsWith(fv),
  ends_with: (cv, fv) => cv.endsWith(fv),
  gt: (cv, fv) => Number(cv) > Number(fv),
  lt: (cv, fv) => Number(cv) < Number(fv),
  gte: (cv, fv) => Number(cv) >= Number(fv),
  lte: (cv, fv) => Number(cv) <= Number(fv),
  is_empty: (cv) => cv === "" || cv === "undefined" || cv === "null",
  is_not_empty: (cv) => cv !== "" && cv !== "undefined" && cv !== "null",
};

function matchesFilters(row: Record<string, string>, filters: FilterCondition[]): boolean {
  if (filters.length === 0) return true;
  return filters.every((f) => {
    const cv = row[f.fieldId] ?? "";
    return OPERATOR_MAP[f.operator]?.(cv, f.value) ?? true;
  });
}

function formatCellValue(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number" || typeof value === "boolean" || typeof value === "string") {
    return String(value);
  }
  if (Array.isArray(value)) {
    return value
      .map((v) => {
        if (typeof v === "object" && v !== null) {
          if ("text" in v) return String(v.text);
          if ("name" in v) return String(v.name);
          if ("link" in v) return String(v.link);
        }
        return String(v);
      })
      .join(", ");
  }
  if (typeof value === "object") {
    if ("text" in (value as object)) return String((value as { text: string }).text);
    if ("name" in (value as object)) return String((value as { name: string }).name);
    if ("link" in (value as object)) return String((value as { link: string }).link);
  }
  return String(value);
}

function extractAttachments(value: unknown): IOpenAttachment[] {
  if (Array.isArray(value)) {
    return value.filter(
      (v) => typeof v === "object" && v !== null && "token" in v
    ) as IOpenAttachment[];
  }
  return [];
}

function getExtension(name: string): string {
  const lower = name.toLowerCase();
  if (lower.endsWith(".png")) return "png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "jpeg";
  if (lower.endsWith(".gif")) return "gif";
  if (lower.endsWith(".webp")) return "png"; // Convert webp to png for Excel compat
  return "png";
}

export async function exportToExcel(options: {
  headers: FieldMeta[];
  rows: Array<{ recordId: string; fields: Record<string, unknown> }>;
  filters: FilterCondition[];
  tableName: string;
  tableId?: string;
  onProgress?: (msg: string) => void;
}): Promise<void> {
  const { headers, rows, filters, tableName, tableId, onProgress } = options;

  const table = tableId
    ? await bitable.base.getTable(tableId)
    : await bitable.base.getActiveTable();

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "多维表格导出插件";
  workbook.created = new Date();

  const sheet = workbook.addWorksheet(tableName || "Sheet1");

  // Build text representation + track attachment data
  type RowData = {
    textValues: Record<string, string>;
    attachments: Record<string, IOpenAttachment[]>;
    recordId: string;
  };
  const allRows: RowData[] = rows.map((r) => ({
    textValues: Object.fromEntries(
      headers.map((h) => [h.id, formatCellValue(r.fields[h.id])])
    ),
    attachments: Object.fromEntries(
      headers.map((h) => [h.id, extractAttachments(r.fields[h.id])])
    ),
    recordId: r.recordId,
  }));

  const filtered = allRows.filter((r) => matchesFilters(r.textValues, filters));

  onProgress?.(`已筛选出 ${filtered.length} 条记录，正在生成 Excel...`);

  // Style definitions
  const HEADER_FILL: ExcelJS.FillPattern = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF3370FF" },
  };
  const HEADER_FONT: Partial<ExcelJS.Font> = {
    bold: true,
    size: 12,
    color: { argb: "FFFFFFFF" },
    name: "Microsoft YaHei",
  };
  const DATA_FONT: Partial<ExcelJS.Font> = {
    size: 10,
    color: { argb: "FF333333" },
    name: "Microsoft YaHei",
  };
  const THIN_BORDER: Partial<ExcelJS.Borders> = {
    top: { style: "thin", color: { argb: "FFD9D9D9" } },
    bottom: { style: "thin", color: { argb: "FFD9D9D9" } },
    left: { style: "thin", color: { argb: "FFD9D9D9" } },
    right: { style: "thin", color: { argb: "FFD9D9D9" } },
  };
  const CENTER_ALIGN: Partial<ExcelJS.Alignment> = {
    horizontal: "center",
    vertical: "middle",
    wrapText: true,
  };
  const DATA_ALIGN: Partial<ExcelJS.Alignment> = {
    vertical: "middle",
    wrapText: true,
  };
  const ALT_FILL: ExcelJS.FillPattern = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF5F7FA" },
  };

  // === Header Row ===
  const headerRow = sheet.addRow(headers.map((h) => h.name));
  headerRow.height = 32;
  headerRow.eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = CENTER_ALIGN;
    cell.border = THIN_BORDER;
  });

  // === Data Rows ===
  for (let i = 0; i < filtered.length; i++) {
    const r = filtered[i];
    const values = headers.map((h) => r.textValues[h.id] ?? "");
    const row = sheet.addRow(values);
    row.eachCell((cell) => {
      cell.font = DATA_FONT;
      cell.alignment = DATA_ALIGN;
      cell.border = THIN_BORDER;
    });
    if (i % 2 === 1) {
      row.eachCell((cell) => {
        cell.fill = ALT_FILL;
      });
    }
  }

  // === Find which columns have attachment/image data ===
  const attachmentCols = new Set<number>();
  for (let c = 0; c < headers.length; c++) {
    const fieldId = headers[c].id;
    if (filtered.some((r) => r.attachments[fieldId]?.length > 0)) {
      attachmentCols.add(c);
    }
  }

  // === Embed images ===
  let totalImages = 0;
  for (const r of filtered) {
    for (const fieldId of Object.keys(r.attachments)) {
      totalImages += r.attachments[fieldId].length;
    }
  }

  let processedImages = 0;
  for (const colIdx of attachmentCols) {
    const fieldId = headers[colIdx].id;
    for (let rowIdx = 0; rowIdx < filtered.length; rowIdx++) {
      const attachments = filtered[rowIdx].attachments[fieldId];
      if (!attachments || attachments.length === 0) continue;

      const excelRow = rowIdx + 2; // 1-indexed, skip header
      const xlRow = sheet.getRow(excelRow);
      xlRow.height = Math.max(xlRow.height || 20, 80);

      for (let imgIdx = 0; imgIdx < Math.min(attachments.length, 3); imgIdx++) {
        const att = attachments[imgIdx];
        processedImages++;
        onProgress?.(
          `正在下载图片 (${processedImages}/${totalImages})...`
        );

        try {
          const url = await table.getAttachmentUrl(
            att.token,
            fieldId,
            filtered[rowIdx].recordId
          );
          const resp = await fetch(url);
          if (!resp.ok) continue;
          const buffer = await resp.arrayBuffer();
          if (buffer.byteLength === 0) continue;

          const ext = getExtension(att.name || "image.png") as "png" | "jpeg" | "gif";
          const imageId = workbook.addImage({
            buffer,
            extension: ext,
          });

          sheet.addImage(imageId, {
            tl: { col: colIdx + 0.05, row: excelRow - 1 + 0.05 },
            br: { col: colIdx + 1 - 0.05, row: excelRow - 1 + 0.95 },
            editAs: "oneCell",
          } as ExcelJS.ImageRange & { editAs: string });
        } catch (e) {
          console.warn(`图片下载失败: ${att.name}`, e);
        }
      }
    }
  }

  // === Auto-fit column widths ===
  for (let c = 0; c < headers.length; c++) {
    const col = sheet.getColumn(c + 1);
    let maxLen = 0;
    for (const r of filtered) {
      const len = (r.textValues[headers[c].id] ?? "").length;
      if (len > maxLen) maxLen = len;
    }
    // Chinese chars are roughly 2x width
    const charLen = [...(headers[c].name)].length;
    const dataLen = maxLen;
    col.width = Math.min(Math.max(charLen * 2 + 2, dataLen * 1.2 + 2, 10), 50);
    if (attachmentCols.has(c)) {
      col.width = Math.max(col.width, 15);
    }
  }

  // Freeze header
  sheet.views = [{ state: "frozen", ySplit: 1 }];

  // Generate and download
  const buf = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const fileName = `${tableName || "export"}_${new Date().toISOString().slice(0, 10)}.xlsx`;
  saveAs(blob, fileName);

  onProgress?.(`导出完成！共 ${filtered.length} 条记录已保存为 ${fileName}`);
}
