import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";
import { mergeExcelsToPeopleJson } from "./lib/roster";
import { buildPaymentsExportRows, buildPaymentsRows, exportPaymentsXlsx } from "./lib/payments";

type PeopleDict = Record<string, any>;

export default function App() {
  const [files, setFiles] = useState<FileList | null>(null);
  const [people, setPeople] = useState<PeopleDict | null>(null);
  const [jsonUrl, setJsonUrl] = useState<string | null>(null);
  const [paymentsReady, setPaymentsReady] = useState(false);

  const selectedFiles = useMemo(() => Array.from(files ?? []), [files]);
  const canMerge = selectedFiles.length > 0;
  const paymentPreviewRows = useMemo(() => {
    if (!people) return [];
    return buildPaymentsExportRows(buildPaymentsRows(people));
  }, [people]);
  const paymentPreviewColumns = useMemo(
    () => (paymentPreviewRows[0] ? Object.keys(paymentPreviewRows[0]) : []),
    [paymentPreviewRows],
  );

  async function onMerge() {
    if (!files || files.length === 0) return;
    try {
      const workbooks: XLSX.WorkBook[] = [];
      for (const file of Array.from(files)) {
        const buf = await file.arrayBuffer();
        workbooks.push(XLSX.read(buf, { type: "array" }));
      }
      const peopleDict = mergeExcelsToPeopleJson(workbooks);
      setPeople(peopleDict);

      const blob = new Blob([JSON.stringify(peopleDict, null, 2)], {
        type: "application/json;charset=utf-8",
      });
      const url = URL.createObjectURL(blob);
      setJsonUrl(url);
      setPaymentsReady(true);
    } catch (e: any) {
      alert("合併失敗：" + (e && e.message ? e.message : e.toString()));
    }
  }

  function onExportJson() {
    if (!jsonUrl) return;
    const a = document.createElement("a");
    a.href = jsonUrl;
    a.download = "roster.json";
    a.click();
  }

  function onExportPayments() {
    if (!people) return;
    const rows = buildPaymentsRows(people);
    exportPaymentsXlsx(rows, "payments.xlsx");
  }

  return (
    <main className="app-shell">
      <section className="hero">
        <div>
          <p className="eyebrow">Volunteer Payment Calculator</p>
          <h1>志工給付計算器</h1>
          <p className="hero-copy">整併值勤 Excel，產生可下載的給付明細表。</p>
        </div>
        <div className="status-strip" aria-label="目前狀態">
          <div>
            <span className="status-label">檔案</span>
            <strong>{selectedFiles.length}</strong>
          </div>
          <div>
            <span className="status-label">筆數</span>
            <strong>{paymentPreviewRows.length}</strong>
          </div>
          <div>
            <span className="status-label">狀態</span>
            <strong>{paymentsReady ? "可匯出" : "待合併"}</strong>
          </div>
        </div>
      </section>

      <section className="workspace">
        <div className="panel upload-panel">
          <div className="panel-header">
            <div>
              <h2>Excel 檔案</h2>
              <p>支援一次選取多個 .xlsx 檔。</p>
            </div>
          </div>

          <label className="file-drop">
            <input
              type="file"
              accept=".xlsx"
              multiple
              onChange={(e) => setFiles(e.target.files)}
            />
            <span className="file-drop-title">選擇 Excel 檔</span>
            <span className="file-drop-meta">
              {selectedFiles.length > 0
                ? `已選擇 ${selectedFiles.length} 個檔案`
                : "尚未選擇檔案"}
            </span>
          </label>

          {selectedFiles.length > 0 && (
            <ul className="file-list" aria-label="已選擇檔案">
              {selectedFiles.map((file) => (
                <li key={`${file.name}-${file.lastModified}`}>
                  <span>{file.name}</span>
                  <small>{Math.max(1, Math.round(file.size / 1024))} KB</small>
                </li>
              ))}
            </ul>
          )}

          <div className="actions">
            <button className="primary-button" disabled={!canMerge} onClick={onMerge}>
              合併資料
            </button>
            <button className="secondary-button" disabled={!paymentsReady} onClick={onExportPayments}>
              下載 Excel
            </button>
            <button hidden disabled={!jsonUrl} onClick={onExportJson}>
              下載 roster.json
            </button>
          </div>
        </div>

        <div className="panel preview-panel">
          <div className="panel-header">
            <div>
              <h2>匯出預覽</h2>
              <p>{paymentPreviewRows.length > 0 ? "內容與下載的 Excel 相同。" : "合併後會顯示給付明細。"}</p>
            </div>
          </div>

          {paymentPreviewRows.length > 0 ? (
            <div className="table-wrap">
              <table className="preview-table">
                <thead>
                  <tr>
                    {paymentPreviewColumns.map((column) => (
                      <th key={column}>{column}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {paymentPreviewRows.map((row, rowIndex) => {
                    const isTotal = row["姓名"] === "總計";
                    return (
                      <tr key={rowIndex} className={isTotal ? "total-row" : undefined}>
                        {paymentPreviewColumns.map((column) => (
                          <td key={column}>{row[column]}</td>
                        ))}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="empty-state">
              <strong>尚無預覽資料</strong>
              <span>選取檔案並合併後，表格會出現在這裡。</span>
            </div>
          )}
        </div>
      </section>
    </main>
  );
}
