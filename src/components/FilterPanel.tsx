import React, { useState, useEffect, useCallback } from "react";
import type { FieldMeta, FilterCondition, SelectOption } from "../utils/excelExporter";

interface FilterPanelProps {
  fields: FieldMeta[];
  filters: FilterCondition[];
  onFiltersChange: (filters: FilterCondition[]) => void;
}

interface SavedRule {
  name: string;
  filters: FilterCondition[];
}

const STORAGE_KEY = "bitable_export_filter_rules";

function loadSavedRules(): SavedRule[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function saveRules(rules: SavedRule[]) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
}

const OPERATORS = [
  { value: "equals", label: "等于" },
  { value: "not_equals", label: "不等于" },
  { value: "contains", label: "包含" },
  { value: "not_contains", label: "不包含" },
  { value: "starts_with", label: "开头是" },
  { value: "ends_with", label: "结尾是" },
  { value: "gt", label: "大于" },
  { value: "lt", label: "小于" },
  { value: "gte", label: "大于等于" },
  { value: "lte", label: "小于等于" },
  { value: "is_empty", label: "为空" },
  { value: "is_not_empty", label: "不为空" },
];

export default function FilterPanel({
  fields,
  filters,
  onFiltersChange,
}: FilterPanelProps) {
  const [collapsed, setCollapsed] = useState(true);
  const [savedRules, setSavedRules] = useState<SavedRule[]>(loadSavedRules());
  const [saveDialogOpen, setSaveDialogOpen] = useState(false);
  const [newRuleName, setNewRuleName] = useState("");

  const getOptions = useCallback(
    (fieldId: string): SelectOption[] | undefined => {
      return fields.find((f) => f.id === fieldId)?.options;
    },
    [fields]
  );

  const isSelectField = useCallback(
    (fieldId: string): boolean => {
      const f = fields.find((fd) => fd.id === fieldId);
      return f?.type === 3 || f?.type === 4; // SingleSelect=3, MultiSelect=4
    },
    [fields]
  );

  const addFilter = () => {
    if (fields.length === 0) return;
    onFiltersChange([
      ...filters,
      {
        fieldId: fields[0].id,
        fieldName: fields[0].name,
        operator: "contains",
        value: "",
      },
    ]);
  };

  const removeFilter = (index: number) => {
    const next = filters.filter((_, i) => i !== index);
    onFiltersChange(next);
  };

  const updateFilter = (index: number, patch: Partial<FilterCondition>) => {
    const next = [...filters];
    const f = next[index];
    if (patch.fieldId && patch.fieldId !== f.fieldId) {
      const field = fields.find((fd) => fd.id === patch.fieldId);
      next[index] = {
        ...f,
        fieldId: patch.fieldId!,
        fieldName: field?.name || "",
        value: "",
      };
    } else {
      next[index] = { ...f, ...patch };
    }
    onFiltersChange(next);
  };

  const handleSaveRule = () => {
    const name = newRuleName.trim();
    if (!name || filters.length === 0) return;
    const rule: SavedRule = { name, filters: [...filters] };
    const updated = [...savedRules, rule];
    setSavedRules(updated);
    saveRules(updated);
    setNewRuleName("");
    setSaveDialogOpen(false);
  };

  const handleLoadRule = (rule: SavedRule) => {
    onFiltersChange(rule.filters.map((f) => ({ ...f })));
  };

  const handleDeleteRule = (index: number) => {
    const updated = savedRules.filter((_, i) => i !== index);
    setSavedRules(updated);
    saveRules(updated);
  };

  return (
    <div style={styles.container}>
      <div
        style={styles.header}
        onClick={() => setCollapsed(!collapsed)}
      >
        <span style={styles.headerIcon}>{collapsed ? "▶" : "▼"}</span>
        <span style={styles.headerTitle}>
          筛选条件
          {filters.length > 0 && (
            <span style={styles.badge}>{filters.length}</span>
          )}
        </span>
      </div>

      {!collapsed && (
        <div style={styles.body}>
          {/* Saved rules */}
          {savedRules.length > 0 && (
            <div style={styles.savedSection}>
              <div style={styles.savedLabel}>已保存的规则</div>
              {savedRules.map((rule, idx) => (
                <div key={idx} style={styles.savedRow}>
                  <button
                    style={styles.savedLoadBtn}
                    onClick={() => handleLoadRule(rule)}
                    title="应用此规则"
                  >
                    📋 {rule.name}
                  </button>
                  <button
                    style={styles.savedDelBtn}
                    onClick={() => handleDeleteRule(idx)}
                    title="删除此规则"
                  >
                    ×
                  </button>
                </div>
              ))}
            </div>
          )}

          {/* Save current filters */}
          {filters.length > 0 && !saveDialogOpen && (
            <button
              style={styles.saveBtn}
              onClick={() => setSaveDialogOpen(true)}
            >
              💾 保存当前规则
            </button>
          )}

          {saveDialogOpen && (
            <div style={styles.saveDialog}>
              <input
                style={styles.saveInput}
                value={newRuleName}
                placeholder="输入规则名称"
                onChange={(e) => setNewRuleName(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleSaveRule()}
              />
              <div style={styles.saveDialogBtns}>
                <button style={styles.saveConfirmBtn} onClick={handleSaveRule}>
                  保存
                </button>
                <button
                  style={styles.saveCancelBtn}
                  onClick={() => {
                    setSaveDialogOpen(false);
                    setNewRuleName("");
                  }}
                >
                  取消
                </button>
              </div>
            </div>
          )}

          {/* Filter rows */}
          {filters.map((filter, idx) => (
            <div key={idx} style={styles.filterRow}>
              <select
                style={styles.select}
                value={filter.fieldId}
                onChange={(e) =>
                  updateFilter(idx, { fieldId: e.target.value })
                }
              >
                {fields.map((f) => (
                  <option key={f.id} value={f.id}>
                    {f.name}
                  </option>
                ))}
              </select>
              <select
                style={styles.selectSmall}
                value={filter.operator}
                onChange={(e) =>
                  updateFilter(idx, { operator: e.target.value })
                }
              >
                {OPERATORS.map((op) => (
                  <option key={op.value} value={op.value}>
                    {op.label}
                  </option>
                ))}
              </select>
              {filter.operator !== "is_empty" &&
                filter.operator !== "is_not_empty" &&
                (isSelectField(filter.fieldId) ? (
                  <select
                    style={styles.input}
                    value={filter.value}
                    onChange={(e) =>
                      updateFilter(idx, { value: e.target.value })
                    }
                  >
                    <option value="">请选择</option>
                    {getOptions(filter.fieldId)?.map((opt) => (
                      <option key={opt.id} value={opt.name}>
                        {opt.name}
                      </option>
                    ))}
                  </select>
                ) : (
                  <input
                    style={styles.input}
                    value={filter.value}
                    placeholder="筛选值"
                    onChange={(e) =>
                      updateFilter(idx, { value: e.target.value })
                    }
                  />
                ))}
              <button
                style={styles.removeBtn}
                onClick={() => removeFilter(idx)}
                title="移除此筛选"
              >
                ×
              </button>
            </div>
          ))}

          <button style={styles.addBtn} onClick={addFilter}>
            + 添加筛选条件
          </button>

          {filters.length > 0 && (
            <button
              style={styles.clearBtn}
              onClick={() => onFiltersChange([])}
            >
              清空所有筛选
            </button>
          )}
        </div>
      )}
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  container: {
    backgroundColor: "#fff",
    borderRadius: 8,
    border: "1px solid #e8e8e8",
    marginBottom: 12,
    overflow: "hidden",
  },
  header: {
    display: "flex",
    alignItems: "center",
    padding: "10px 14px",
    cursor: "pointer",
    userSelect: "none",
    gap: 8,
  },
  headerIcon: { fontSize: 12, color: "#999", flexShrink: 0 },
  headerTitle: {
    fontSize: 14,
    fontWeight: 600,
    color: "#333",
    display: "flex",
    alignItems: "center",
    gap: 6,
  },
  badge: {
    backgroundColor: "#3370ff",
    color: "#fff",
    borderRadius: 10,
    padding: "1px 6px",
    fontSize: 11,
    fontWeight: 500,
  },
  body: {
    padding: "0 14px 12px",
    borderTop: "1px solid #f0f0f0",
  },
  savedSection: {
    marginTop: 10,
    marginBottom: 4,
  },
  savedLabel: {
    fontSize: 12,
    color: "#999",
    marginBottom: 6,
  },
  savedRow: {
    display: "flex",
    alignItems: "center",
    gap: 4,
    marginBottom: 4,
  },
  savedLoadBtn: {
    flex: 1,
    padding: "5px 8px",
    borderRadius: 4,
    border: "1px solid #d6e4ff",
    backgroundColor: "#f0f5ff",
    fontSize: 12,
    color: "#3370ff",
    cursor: "pointer",
    textAlign: "left",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  savedDelBtn: {
    width: 22,
    height: 22,
    borderRadius: 4,
    border: "none",
    backgroundColor: "#ffccc7",
    color: "#ff4d4f",
    fontSize: 13,
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
  },
  saveBtn: {
    marginTop: 10,
    width: "100%",
    padding: "6px 0",
    borderRadius: 6,
    border: "1px solid #d6e4ff",
    backgroundColor: "#f0f5ff",
    color: "#3370ff",
    fontSize: 13,
    cursor: "pointer",
  },
  saveDialog: {
    marginTop: 8,
    display: "flex",
    gap: 6,
    alignItems: "center",
  },
  saveInput: {
    flex: 1,
    padding: "6px 8px",
    borderRadius: 6,
    border: "1px solid #ddd",
    fontSize: 13,
    color: "#333",
    outline: "none",
  },
  saveDialogBtns: {
    display: "flex",
    gap: 4,
    flexShrink: 0,
  },
  saveConfirmBtn: {
    padding: "6px 12px",
    borderRadius: 6,
    border: "none",
    backgroundColor: "#3370ff",
    color: "#fff",
    fontSize: 13,
    cursor: "pointer",
  },
  saveCancelBtn: {
    padding: "6px 12px",
    borderRadius: 6,
    border: "1px solid #ddd",
    backgroundColor: "#fff",
    color: "#666",
    fontSize: 13,
    cursor: "pointer",
  },
  filterRow: {
    display: "flex",
    gap: 6,
    marginTop: 10,
    alignItems: "center",
    flexWrap: "wrap",
  },
  select: {
    flex: "1 1 40%",
    minWidth: 0,
    padding: "6px 8px",
    borderRadius: 6,
    border: "1px solid #ddd",
    fontSize: 13,
    color: "#333",
    backgroundColor: "#fff",
    outline: "none",
  },
  selectSmall: {
    flex: "1 1 35%",
    minWidth: 0,
    padding: "6px 8px",
    borderRadius: 6,
    border: "1px solid #ddd",
    fontSize: 13,
    color: "#333",
    backgroundColor: "#fff",
    outline: "none",
  },
  input: {
    flex: "1 1 40%",
    minWidth: 0,
    padding: "6px 8px",
    borderRadius: 6,
    border: "1px solid #ddd",
    fontSize: 13,
    color: "#333",
    outline: "none",
  },
  removeBtn: {
    width: 26,
    height: 26,
    borderRadius: 6,
    border: "none",
    backgroundColor: "#ff4d4f",
    color: "#fff",
    fontSize: 14,
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
  },
  addBtn: {
    marginTop: 10,
    width: "100%",
    padding: "7px 0",
    borderRadius: 6,
    border: "1px dashed #bbb",
    backgroundColor: "#fafafa",
    color: "#666",
    fontSize: 13,
    cursor: "pointer",
  },
  clearBtn: {
    marginTop: 6,
    width: "100%",
    padding: "7px 0",
    borderRadius: 6,
    border: "1px solid #ffccc7",
    backgroundColor: "#fff2f0",
    color: "#ff4d4f",
    fontSize: 13,
    cursor: "pointer",
  },
};
