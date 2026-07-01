import { useEffect, useState, useCallback, useRef } from "react";
import { bitable } from "@lark-base-open/js-sdk";
import type { FilterCondition } from "./utils/excelExporter";
import { exportToExcel } from "./utils/excelExporter";
import {
  getTableList,
  getFieldList,
  getTableRecords,
  getRecordCount,
  type TableInfo,
  FIELD_TYPE_LABELS,
} from "./utils/bitbleHelper";
import FilterPanel from "./components/FilterPanel";

type TableStatus = "loading" | "ready" | "error";

function App() {
  const [tables, setTables] = useState<TableInfo[]>([]);
  const [selectedTableId, setSelectedTableId] = useState<string>("");
  const [fields, setFields] = useState<
    Array<{ id: string; name: string; type: number }>
  >([]);
  const [filters, setFilters] = useState<FilterCondition[]>([]);
  const [exporting, setExporting] = useState(false);
  const [statusMsg, setStatusMsg] = useState("");
  const [recordCount, setRecordCount] = useState<number | null>(null);
  const [tableStatus, setTableStatus] = useState<TableStatus>("loading");
  const loadingRef = useRef(false);

  const loadTableInfo = useCallback(async (tableId: string) => {
    setFields([]);
    setRecordCount(null);
    setFilters([]);
    setStatusMsg("");
    setTableStatus("loading");
    if (!tableId) return;
    loadingRef.current = true;

    try {
      const [fl, count] = await Promise.all([
        getFieldList(tableId),
        getRecordCount(tableId),
      ]);
      if (!loadingRef.current) return; // stale request guard
      setFields(fl);
      setRecordCount(count);
      setTableStatus("ready");
      const tableName = tables.find((t) => t.id === tableId)?.name || "";
      setStatusMsg(`当前表「${tableName}」共 ${fl.length} 个字段，${count} 条记录`);
    } catch (err) {
      console.error("获取表信息失败", err);
      if (!loadingRef.current) return;
      setTableStatus("error");
      setStatusMsg("获取表信息失败");
    } finally {
      loadingRef.current = false;
    }
  }, [tables]);

  // Load table list on mount
  useEffect(() => {
    const init = async () => {
      try {
        const list = await getTableList();
        setTables(list);
        if (list.length > 0) {
          const activeTable = await bitable.base.getActiveTable();
          const meta = await activeTable.getMeta();
          const match = meta.id ? list.find((t) => t.id === meta.id) : null;
          const initialId = match ? match.id : list[0].id;
          setSelectedTableId(initialId);
        }
      } catch (err) {
        console.error("获取表列表失败", err);
      }
    };
    init();
  }, []);

  // Load fields when selectedTableId changes
  useEffect(() => {
    if (selectedTableId) {
      loadTableInfo(selectedTableId);
    }
  }, [selectedTableId, loadTableInfo]);

  // Listen for table switch events from Bitable host
  useEffect(() => {
    const off = bitable.base.onSelectionChange((e) => {
      const newTableId = e.data?.tableId;
      if (newTableId && newTableId !== selectedTableId) {
        setTables((prev) => {
          // Refresh table list to catch any additions
          getTableList().then((list) => {
            if (list.some((t) => t.id !== prev[0]?.id)) {
              setTables(list);
            }
          });
          return prev;
        });
        setSelectedTableId(newTableId);
      }
    });
    return () => {
      off?.();
    };
  }, [selectedTableId]);

  // Export
  const handleExport = async () => {
    if (!selectedTableId || fields.length === 0) {
      setStatusMsg("请先选择一个数据表");
      return;
    }
    setExporting(true);
    setStatusMsg("正在读取数据...");
    try {
      const records = await getTableRecords(selectedTableId);
      const tableName =
        tables.find((t) => t.id === selectedTableId)?.name || "导出";

      await exportToExcel({
        headers: fields,
        rows: records,
        filters,
        tableName,
        onProgress: (msg) => setStatusMsg(msg),
      });
    } catch (err) {
      console.error("导出失败", err);
      setStatusMsg("导出失败：" + String(err));
    } finally {
      setExporting(false);
    }
  };

  const selectedTable = tables.find((t) => t.id === selectedTableId);

  return (
    <div style={styles.root}>
      {/* Header */}
      <div style={styles.header}>
        <div style={styles.headerLeft}>
          <span style={styles.logo}>📊</span>
          <span style={styles.title}>多维表格导出</span>
        </div>
      </div>

      {/* Table selector */}
      <div style={styles.section}>
        <label style={styles.label}>选择数据表</label>
        <select
          style={styles.fullSelect}
          value={selectedTableId}
          onChange={(e) => setSelectedTableId(e.target.value)}
        >
          {tables.length === 0 && <option value="">加载中...</option>}
          {tables.map((t) => (
            <option key={t.id} value={t.id}>
              {t.name}
            </option>
          ))}
        </select>
      </div>

      {/* Table info card */}
      {selectedTableId && tableStatus === "ready" && (
        <div style={styles.infoCard}>
          <div style={styles.infoTitle}>{selectedTable?.name || ""}</div>
          <div style={styles.infoRow}>
            <div style={styles.infoItem}>
              <span style={styles.infoValue}>{fields.length}</span>
              <span style={styles.infoLabel}>个字段</span>
            </div>
            <div style={styles.infoDivider} />
            <div style={styles.infoItem}>
              <span style={styles.infoValue}>
                {recordCount !== null ? recordCount : "-"}
              </span>
              <span style={styles.infoLabel}>条记录</span>
            </div>
          </div>
          <div style={styles.fieldList}>
            {fields.map((f) => (
              <div key={f.id} style={styles.fieldTag}>
                <span style={styles.fieldName}>{f.name}</span>
                <span style={styles.fieldType}>
                  {FIELD_TYPE_LABELS[f.type] || "未知"}
                </span>
              </div>
            ))}
          </div>
        </div>
      )}

      {tableStatus === "loading" && (
        <div style={styles.loadingCard}>正在加载表信息...</div>
      )}

      {tableStatus === "error" && (
        <div style={styles.errorCard}>加载失败，请重试</div>
      )}

      {/* Filter panel */}
      <FilterPanel
        fields={fields}
        filters={filters}
        onFiltersChange={setFilters}
      />

      {/* Export button */}
      <div style={styles.section}>
        <button
          style={{
            ...styles.exportBtn,
            ...((exporting || tableStatus !== "ready")
              ? styles.exportBtnDisabled
              : {}),
          }}
          onClick={handleExport}
          disabled={exporting || tableStatus !== "ready"}
        >
          {exporting ? "导出中..." : "导出 Excel"}
        </button>
      </div>

      {/* Status */}
      {statusMsg && (
        <div style={styles.statusBar}>
          <span style={styles.statusDot} />
          {statusMsg}
        </div>
      )}

      {/* Footer */}
      <div style={styles.footer}>
        <p>导出格式: .xlsx (兼容 WPS / Microsoft Excel)</p>
        <p>图片自动嵌入单元格 · 表头加粗美化 · 隔行变色</p>
      </div>
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  root: {
    padding: 16,
    fontFamily:
      '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Microsoft YaHei", sans-serif',
    color: "#333",
    minHeight: "100vh",
    backgroundColor: "#f7f8fa",
  },
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: 16,
    paddingBottom: 12,
    borderBottom: "2px solid #3370ff",
  },
  headerLeft: {
    display: "flex",
    alignItems: "center",
    gap: 8,
  },
  logo: { fontSize: 22 },
  title: {
    fontSize: 17,
    fontWeight: 700,
    color: "#1f2329",
  },
  section: {
    marginBottom: 12,
  },
  label: {
    display: "block",
    fontSize: 13,
    fontWeight: 600,
    color: "#555",
    marginBottom: 6,
  },
  fullSelect: {
    width: "100%",
    padding: "8px 10px",
    borderRadius: 6,
    border: "1px solid #ddd",
    fontSize: 13,
    color: "#333",
    backgroundColor: "#fff",
    outline: "none",
    boxSizing: "border-box",
  },
  infoCard: {
    backgroundColor: "#fff",
    borderRadius: 8,
    border: "1px solid #e8e8e8",
    padding: "12px 14px",
    marginBottom: 12,
  },
  infoTitle: {
    fontSize: 15,
    fontWeight: 700,
    color: "#1f2329",
    marginBottom: 10,
  },
  infoRow: {
    display: "flex",
    alignItems: "center",
    gap: 0,
    marginBottom: 12,
  },
  infoItem: {
    flex: 1,
    textAlign: "center",
    display: "flex",
    alignItems: "baseline",
    justifyContent: "center",
    gap: 4,
  },
  infoValue: {
    fontSize: 20,
    fontWeight: 700,
    color: "#3370ff",
  },
  infoLabel: {
    fontSize: 12,
    color: "#999",
  },
  infoDivider: {
    width: 1,
    height: 28,
    backgroundColor: "#eee",
  },
  fieldList: {
    display: "flex",
    flexWrap: "wrap",
    gap: 6,
    maxHeight: 200,
    overflowY: "auto",
    paddingTop: 2,
  },
  fieldTag: {
    display: "inline-flex",
    alignItems: "center",
    gap: 4,
    padding: "3px 8px",
    borderRadius: 4,
    backgroundColor: "#f0f5ff",
    border: "1px solid #d6e4ff",
    fontSize: 12,
    lineHeight: "18px",
    maxWidth: "100%",
  },
  fieldName: {
    color: "#1f2329",
    fontWeight: 500,
    maxWidth: 120,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  fieldType: {
    color: "#3370ff",
    fontSize: 11,
    flexShrink: 0,
  },
  loadingCard: {
    backgroundColor: "#fff",
    borderRadius: 8,
    border: "1px solid #e8e8e8",
    padding: "16px 14px",
    marginBottom: 12,
    textAlign: "center",
    color: "#999",
    fontSize: 13,
  },
  errorCard: {
    backgroundColor: "#fff2f0",
    borderRadius: 8,
    border: "1px solid #ffccc7",
    padding: "12px 14px",
    marginBottom: 12,
    textAlign: "center",
    color: "#ff4d4f",
    fontSize: 13,
  },
  exportBtn: {
    width: "100%",
    padding: "10px 0",
    borderRadius: 8,
    border: "none",
    backgroundColor: "#3370ff",
    color: "#fff",
    fontSize: 15,
    fontWeight: 600,
    cursor: "pointer",
  },
  exportBtnDisabled: {
    backgroundColor: "#99b8ff",
    cursor: "not-allowed",
  },
  statusBar: {
    marginTop: 10,
    padding: "8px 12px",
    backgroundColor: "#fff",
    borderRadius: 6,
    fontSize: 12,
    color: "#666",
    border: "1px solid #e8e8e8",
    display: "flex",
    alignItems: "center",
    gap: 6,
  },
  statusDot: {
    width: 6,
    height: 6,
    borderRadius: 3,
    backgroundColor: "#3370ff",
    flexShrink: 0,
  },
  footer: {
    marginTop: 20,
    paddingTop: 12,
    borderTop: "1px solid #eee",
    fontSize: 11,
    color: "#aaa",
    lineHeight: "20px",
  },
};

export default App;
