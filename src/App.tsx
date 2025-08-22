import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { mergeExcelsToPeopleJson } from "./lib/roster";
import { buildPaymentsRows, exportPaymentsXlsx } from "./lib/payments";

type PeopleDict = Record<string, any>;

export default function App() {
  const [files, setFiles] = useState<FileList | null>(null);
  const [people, setPeople] = useState<PeopleDict | null>(null);
  const [jsonUrl, setJsonUrl] = useState<string | null>(null);
  const [paymentsReady, setPaymentsReady] = useState(false);

  const canMerge = useMemo(() => files && files.length > 0, [files]);

  async function onMerge() {
    if (!files || files.length === 0) return;
    try {
      // 讀進多個 Excel 並合併
      const workbooks: XLSX.WorkBook[] = [];
      for (const f of Array.from(files)) {
        const buf = await f.arrayBuffer();
        workbooks.push(XLSX.read(buf, { type: "array" }));
      }
      const peopleDict = mergeExcelsToPeopleJson(workbooks);
      setPeople(peopleDict);

      // 下載 roster.json
      const blob = new Blob([JSON.stringify(peopleDict, null, 2)], {
        type: "application/json;charset=utf-8",
      });
      const url = URL.createObjectURL(blob);
      setJsonUrl(url);
      setPaymentsReady(true);
    } catch (e: any) {
      alert("合併失敗：" + e?.message ?? e);
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
    <div style={{ maxWidth: 960, margin: "40px auto", padding: 16, fontFamily: "system-ui, sans-serif" }}>
      <h1>Roster Tool（前端版）</h1>
      <p>上傳多個 Excel（.xlsx），合併→去重→輸出 roster.json 與 payments.xlsx。</p>

      <div style={{ marginTop: 16 }}>
        <input
          type="file"
          accept=".xlsx"
          multiple
          onChange={(e) => setFiles(e.target.files)}
        />
      </div>

      <div style={{ marginTop: 16, display: "flex", gap: 8 }}>
        <button disabled={!canMerge} onClick={onMerge}>合併並產生 JSON</button>
        <button disabled={!jsonUrl} onClick={onExportJson}>下載 roster.json</button>
        <button disabled={!paymentsReady} onClick={onExportPayments}>下載 payments.xlsx</button>
      </div>

      {people && (
        <details style={{ marginTop: 16 }}>
          <summary>預覽（僅前幾筆）</summary>
          <pre style={{ whiteSpace: "pre-wrap" }}>
            {JSON.stringify(Object.fromEntries(Object.entries(people).slice(0, 3)), null, 2)}
          </pre>
        </details>
      )}
    </div>
  );
}
