import * as XLSX from "xlsx";

type PeopleDict = Record<string, any>;

function isLeave(status: any): boolean {
  return String(status ?? "").trim() !== "正常";
}
function safeFloat(x: any): number {
  const n = Number(x);
  return Number.isNaN(n) ? 0 : n;
}
function fmtMMDD(dateStr: string): string {
  if (!dateStr) return "";
  const s = String(dateStr).replace(/\//g, "-");
  const parts = s.split("-");
  if (parts.length >= 3) {
    const m = String(Number(parts[1])).padStart(2, "0");
    const d = String(Number(parts[2])).padStart(2, "0");
    return `${m}/${d}`;
  }
  return dateStr;
}
function hoursToDisplay(h: number): string {
  const s = h.toFixed(2).replace(/\.?0+$/, "");
  return s;
}

export function buildPaymentsRows(people: PeopleDict) {
  // people 可能是 dict keyed by id/name
  const persons = Array.isArray(people) ? people : Object.values(people);
  const rows = persons.map((person: any) => {
    const name = person.name ?? "";
    const pid = person.id_number ?? "";
    const duty = Array.isArray(person.duty_records) ? person.duty_records : [];

    const daily: Record<string, number> = {};
    let totalShifts = 0;

    for (const rec of duty) {
      if (isLeave(rec.status)) continue;
      const d = String(rec.service_date ?? "").trim();
      if (!d) continue;
      const hours = rec.actual_hours ?? rec.std_hours;
      const h = safeFloat(hours);
      daily[d] = (daily[d] ?? 0) + h;
      totalShifts += 1;
    }

    const parts: string[] = [];
    for (const d of Object.keys(daily).sort()) {
      parts.push(`${fmtMMDD(d)}(${hoursToDisplay(daily[d])})`);
    }
    const dutyHoursStr = parts.join("，");

    const uniqueDays = Object.keys(daily).length;
    const traffic = uniqueDays * 50;
    const meals = totalShifts * 100;
    const total = traffic + meals;

    return {
      序號: 0, // 稍後填
      姓名: name,
      身分證字號: pid,
      值勤時數: dutyHoursStr,
      誤餐費: meals,
      交通費: traffic,
      給付總額: total,
      蓋章: "",
    };
  });

  // 排序 & 序號
  rows.sort((a, b) => (a.姓名 + a.身分證字號).localeCompare(b.姓名 + b.身分證字號));
  rows.forEach((r, i) => (r.序號 = i + 1));
  return rows;
}

export function exportPaymentsXlsx(rows: any[], filename: string) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "payments");
  XLSX.writeFile(wb, filename);
}
