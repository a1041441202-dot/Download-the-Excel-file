import { bitable, FieldType } from "@lark-base-open/js-sdk";
import type { FieldMeta } from "./excelExporter";

export interface TableInfo {
  id: string;
  name: string;
}

export async function getTableList(): Promise<TableInfo[]> {
  const metaList = await bitable.base.getTableMetaList();
  return metaList.map((m) => ({
    id: m.id,
    name: m.name,
  }));
}

export async function getFieldList(tableId?: string): Promise<FieldMeta[]> {
  const table = tableId
    ? await bitable.base.getTable(tableId)
    : await bitable.base.getActiveTable();

  // Use active view's getFieldMetaList — returns fields in visual column order.
  // ITable.getFieldMetaList() does NOT guarantee ordering.
  let metaList: Array<{ id?: string; name?: string; type: number }>;
  try {
    const view = await table.getActiveView();
    metaList = await view.getFieldMetaList();
  } catch {
    metaList = await table.getFieldMetaList();
  }

  const fields: FieldMeta[] = metaList.map((m) => ({
    id: m.id || "",
    name: m.name || "",
    type: m.type as number,
  }));

  // Fetch options for select fields via table.getField(fieldId).getMeta()
  // because view.getFieldMetaList() may not include property/options at runtime.
  const selectFieldIds = fields
    .filter((f) => f.type === FieldType.SingleSelect || f.type === FieldType.MultiSelect)
    .map((f) => f.id);

  if (selectFieldIds.length > 0) {
    const optionsMap = new Map<string, Array<{ id: string; name: string }>>();
    await Promise.all(
      selectFieldIds.map(async (fid) => {
        try {
          const field = await table.getField(fid);
          const meta = await field.getMeta();
          const raw = meta as unknown as {
            property?: { options?: Array<{ id: string; name: string }> };
          };
          if (raw.property?.options) {
            optionsMap.set(fid, raw.property.options.map((o) => ({ id: o.id, name: o.name })));
          }
        } catch {
          // ignore individual field errors
        }
      })
    );
    for (const f of fields) {
      if (optionsMap.has(f.id)) {
        f.options = optionsMap.get(f.id);
      }
    }
  }

  return fields;
}

export async function getRecordCount(tableId?: string): Promise<number> {
  const table = tableId
    ? await bitable.base.getTable(tableId)
    : await bitable.base.getActiveTable();
  const recordIdList = await table.getRecordIdList();
  return recordIdList.length;
}

export async function getTableRecords(
  tableId?: string
): Promise<Array<{ recordId: string; fields: Record<string, unknown> }>> {
  const table = tableId
    ? await bitable.base.getTable(tableId)
    : await bitable.base.getActiveTable();
  const recordIdList = await table.getRecordIdList();
  const records: Array<{
    recordId: string;
    fields: Record<string, unknown>;
  }> = [];
  for (const recordId of recordIdList) {
    const record = await table.getRecordById(recordId);
    records.push({
      recordId,
      fields: record.fields as Record<string, unknown>,
    });
  }
  return records;
}

export const FIELD_TYPE_LABELS: Record<number, string> = {
  [FieldType.Number]: "数字",
  [FieldType.Text]: "文本",
  [FieldType.DateTime]: "日期时间",
  [FieldType.Checkbox]: "复选框",
  [FieldType.Phone]: "电话",
  [FieldType.Url]: "邮箱",
  [FieldType.Currency]: "货币",
  [FieldType.Attachment]: "附件/图片",
  [FieldType.DuplexLink]: "关联",
  [FieldType.Rating]: "评分",
  [FieldType.Formula]: "公式",
  [FieldType.Progress]: "进度",
  [FieldType.CreatedTime]: "创建时间",
  [FieldType.ModifiedTime]: "修改时间",
  [FieldType.AutoNumber]: "自动编号",
  [FieldType.Lookup]: "引用",
  [FieldType.GroupChat]: "群聊",
  [FieldType.SingleSelect]: "单选项",
  [FieldType.MultiSelect]: "多选项",
  [FieldType.CreatedUser]: "创建人",
  [FieldType.ModifiedUser]: "最后修改人",
};
