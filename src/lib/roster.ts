import * as XLSX from "xlsx";

type Row = Record<string, any>;
type PeopleDict = Record<string, any>;

const COLUMN_MAP: Record<string, string> = {
  "排班名稱": "schedule_name",
  "服務位置": "location",
  "服務日期": "service_date",
  "班別(早班、午班、晚班)": "shift",
  "服務時段": "service_range",
  "志工姓名": "name",
  "身分證字號": "id_number",
  "標準時數": "std_hours",
  "實際簽到": "actual_checkin",
  "簽到": "planned_checkin",
  "實際簽退": "actual_checkout",
  "簽退": "planned_checkout",
  "實際時數": "actual_hours",
  "出勤狀態": "status",
};

const ALT_COLUMN_MAP: Record<string, string> = {
  "班別": "shift",
  "服務時段(起訖)": "service_range",
  "志工代號": "id_number",
  "實簽到": "actual_checkin",
  "實簽退": "actual_checkout",
  "應簽到": "planned_checkin",
  "應簽退": "planned_checkout",
};

function normalizeColumns(row: Row): Row {
  const out: Row = {};
  for (const [k, v] of Object.entries(row)) {
    const key = (COLUMN_MAP[k] ?? ALT_COLUMN_MAP[k] ?? k).trim();
    out[key] = v;
  }
  return out;
}

function parseTimeRange(s?: string | null): [string | null, string | null] {
  if (!s) return [null, null];
  const m = String(s).match(/^\s*(\d{1,2}:\d{2})\s*[~-]\s*(\d{1,2}:\d{2})\s*$/);
  if (!m) return [null, null];
  return [m[1], m[2]];
}

function cleanTime(t: any): string | null {
  if (t == null || t === "" || t === "-" || t === "—") return null;
  const s = String(t).trim();
  // Excel fraction (0~1) -> HH:MM
  const num = Number(s);
  if (!Number.isNaN(num) && num >= 0 && num < 1) {
    const minutes = Math.round(num * 24 * 60);
    const hh = String(Math.floor(minutes / 60)).padStart(2, "0");
    const mm = String(minutes % 60).padStart(2, "0");
    return `${hh}:${mm}`;
  }
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return `${String(Number(m[1])).padStart(2, "0")}:${m[2]}`;
  return s;
}

function cleanHours(x: any): number | null {
  if (x == null || x === "" || x === "-") return null;
  const n = Number(x);
  return Number.isNaN(n) ? null : n;
}

function sheetToDF(wb: XLSX.WorkBook): Row[] {
  // 取第一個非空工作表
  const sheetName = wb.SheetNames.find((n) => {
    const ws = wb.Sheets[n];
    return ws && ws["!ref"];
  });
  if (!sheetName) return [];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json<Row>(ws, { raw: false, defval: "" });
  return rows.map(normalizeColumns);
}

export function mergeExcelsToPeopleJson(workbooks: XLSX.WorkBook[]): PeopleDict {
  // 合成一個大的 row 陣列
  const rows: Row[] = [];
  for (const wb of workbooks) {
    rows.push(...sheetToDF(wb));
  }

  // 補欄位、拆時段
  const ensured = rows.map((r) => {
    const rr: Row = { ...r };
    if ("service_range" in rr) {
      const [a, b] = parseTimeRange(rr["service_range"]);
      rr["service_start"] = cleanTime(a);
      rr["service_end"] = cleanTime(b);
    }
    for (const c of ["actual_checkin","planned_checkin","actual_checkout","planned_checkout"]) {
      if (c in rr) rr[c] = cleanTime(rr[c]);
    }
    rr["actual_hours"] = cleanHours(rr["actual_hours"]);
    rr["std_hours"] = cleanHours(rr["std_hours"]);
    rr["name"] = rr["name"] ? String(rr["name"]).trim() : null;
    rr["id_number"] = rr["id_number"] ? String(rr["id_number"]).trim() : null;
    return rr;
  });

  // 依人彙整 + 去重（同人＋同日＋同時段＋班別）
  const people: PeopleDict = {};
  function personKey(r: Row): string {
    if (r.id_number) return `id:${r.id_number}`;
    if (r.name) return `name:${r.name}`;
    return `anon:${r.service_date ?? ""}-${r.shift ?? ""}`;
    }

  for (const r of ensured) {
    const key = personKey(r);
    if (!people[key]) {
      people[key] = {
        name: r.name ?? null,
        id_number: r.id_number ?? null,
        locations: new Set<string>(),
        duty_records: [] as Row[],
        _seen: new Set<string>(), // for dedup
      };
    }
    const p = people[key];

    if (!p.name && r.name) p.name = r.name;
    if (!p.id_number && r.id_number) p.id_number = r.id_number;
    if (r.location) p.locations.add(String(r.location));

    const svc_date = r.service_date ? String(r.service_date) : "";
    const dedupKey = JSON.stringify([
      svc_date,
      r.shift ?? null,
      r.service_start ?? null,
      r.service_end ?? null,
    ]);
    if (p._seen.has(dedupKey)) continue;
    p._seen.add(dedupKey);

    p.duty_records.push({
      schedule_name: r.schedule_name ?? null,
      location: r.location ?? null,
      service_date: svc_date || null,
      shift: r.shift ?? null,
      service_start: r.service_start ?? null,
      service_end: r.service_end ?? null,
      planned_checkin: r.planned_checkin ?? null,
      planned_checkout: r.planned_checkout ?? null,
      actual_checkin: r.actual_checkin ?? null,
      actual_checkout: r.actual_checkout ?? null,
      std_hours: r.std_hours ?? null,
      actual_hours: r.actual_hours ?? null,
      status: r.status ?? null,
    });
  }

  // 整理輸出
  for (const k of Object.keys(people)) {
    const p = people[k];
    p.locations = Array.from(p.locations).sort();
    p.duty_records.sort((a: Row, b: Row) => {
      const da = (a.service_date ?? "") + (a.service_start ?? "");
      const db = (b.service_date ?? "") + (b.service_start ?? "");
      return da.localeCompare(db);
    });
    delete p._seen;
  }

  return people;
}
