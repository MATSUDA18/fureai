// 利用日誌Excel（目印方式）の差し込みロジック。ExcelJSの最小限のAPIだけを使う（テストで差し替え可能にするため）。
export const CATS: [string, string[]][] = [
  ["来所", ["来所", "到着", "入室", "登所", "お迎え", "迎えに"]],
  ["ベッドメイク", ["ベッドメイク", "ベッドメーキング", "布団敷", "寝具準備"]],
  ["夕食準備（調理）", ["夕食準備", "夕飯準備", "夕食の準備", "夕食づくり", "調理", "下ごしらえ", "下準備"]],
  ["朝食準備", ["朝食準備", "朝食の準備", "朝食づくり"]],
  ["食器洗い", ["食器洗い", "皿洗い", "洗い物", "食器を洗", "食器片"]],
  ["夕食", ["夕食", "夕飯", "夕ご飯", "夕ごはん"]],
  ["朝食", ["朝食", "朝ご飯", "朝ごはん"]],
  ["入浴", ["入浴", "お風呂", "風呂", "シャワー", "清拭"]],
  ["洗濯", ["洗濯", "ランドリー", "洗濯物"]],
  ["口腔ケア", ["口腔", "歯磨き", "歯みがき", "うがい", "義歯"]],
  ["就寝", ["就寝", "消灯", "寝る", "床に就", "入眠", "おやすみ"]],
  ["起床", ["起床", "起きる", "目覚め", "覚醒"]],
  ["巡回", ["巡回", "巡視", "見回り", "ラウンド", "安否確認"]],
  ["シーツ外し", ["シーツ外し", "シーツを外", "シーツ剥", "シーツ交換", "シーツはがし", "シーツ畳"]],
  ["体操（ダンス）", ["体操", "ダンス", "ストレッチ", "運動"]],
  ["身支度", ["身支度", "着替え", "更衣", "整容", "洗面", "身だしなみ"]],
  ["清掃", ["清掃", "掃除", "片付け", "片づけ", "モップ", "ゴミ"]],
  ["余暇時間", ["余暇", "YouTube", "ユーチューブ", "テレビ", "TV", "休憩", "自由時間", "読書", "ゲーム", "音楽", "談笑", "レクリエーション", "レク", "番組", "動画"]],
  ["出勤or帰宅", ["出勤", "帰宅", "退勤", "降所", "退所", "お帰り"]],
];
export function detectTitles(text: string): string[] {
  let w = text; const hits: string[] = [];
  for (const [label, kws] of CATS) {
    if (kws.some((k) => w.includes(k))) {
      if (label === "出勤or帰宅") hits.push(/帰宅|退勤|降所|退所|お帰り/.test(w) ? "帰宅" : "出勤");
      else hits.push(label);
      if (label === "夕食準備（調理）") w = w.split("夕食準備").join("").split("夕飯準備").join("");
      if (label === "朝食準備") w = w.split("朝食準備").join("");
    }
  }
  return [...new Set(hits)];
}
export function dispLen(s: string): number { let n = 0; for (const ch of s) n += /[\x00-\xff｡-ﾟ]/.test(ch) ? 0.5 : 1; return n; }
export function wrapLines(text: string, widthUnits: number, fontSize: number, div: number): number {
  const perLine = Math.max(4, (widthUnits / div) * (11 / fontSize));
  let lines = 0; for (const seg of text.split("\n")) lines += Math.max(1, Math.ceil(dispLen(seg) / perLine));
  return Math.max(1, lines);
}
export function hardWrap(text: string, widthUnits: number, fontSize: number, div: number): string {
  const perLine = Math.max(4, (widthUnits / div) * (11 / fontSize)); const out: string[] = [];
  for (const seg of text.split("\n")) { let line = "", n = 0; for (const ch of seg) { const w = dispLen(ch); if (n + w > perLine && line) { out.push(line); line = ""; n = 0; } line += ch; n += w; } out.push(line); }
  return out.join("\n");
}
function parseTime(l: string): number | null { const m = l.match(/(\d{1,2})[:：](\d{2})/); return m ? parseInt(m[1], 10) * 60 + parseInt(m[2], 10) : null; }
function eatLine(main: any, side: any): string {
  const m = (main === "" || main == null) ? null : Number(main); const s = (side === "" || side == null) ? null : Number(side);
  if (m == null && s == null) return "";
  const part = (v: number | null, label: string) => v == null ? null : (v <= 0 ? label + "完食" : label + v + "割残し");
  const parts = [part(m, "主食"), part(s, "副食")].filter(Boolean); return parts.length ? "喫食　" + parts.join("／") : "";
}
export function colNum(s: string): number { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
export const MK = /\{\{([^{}|]+?)(?:\|([^{}]*))?\}\}/g;
export const KNOWN_MARKERS = ["利用者名", "利用日時", "測定時間1", "体温1", "血圧1", "脈拍1", "測定時間2", "体温2", "血圧2", "脈拍2", "排尿", "排便", "投薬", "担当職員", "他利用者",
  "活動16", "活動17", "活動18", "活動19", "活動20", "活動21", "活動22", "活動23", "活動0", "活動1", "活動2", "活動3", "活動4", "活動5", "活動6", "活動7", "活動8", "活動9",
  "夕食メニュー", "朝食メニュー", "夕食写真", "朝食写真", "ご様子", "次回利用予定日", "次回へのお願い", "写真", "記録者", "日付"];
const DIV_NARROW = 1.75, DIV_WIDE = 1.9, DIV_MATOME = 1.9, LH = 1.9;
type Rng = { r1: number; c1: number; r2: number; c2: number };
export function cellText(v: any): string {
  if (v == null) return ""; if (typeof v === "string") return v; if (typeof v === "number") return String(v);
  if (Array.isArray(v.richText)) return v.richText.map((x: any) => x.text || "").join("");
  if (v.text != null) return String(v.text); if (v.result != null) return String(v.result); return "";
}
export function mergeMap(ws: any) {
  const list: Rng[] = [];
  for (const m of ((ws.model && ws.model.merges) || [])) { const x = String(m).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/); if (x) list.push({ c1: colNum(x[1]), r1: +x[2], c2: colNum(x[3]), r2: +x[4] }); }
  const slave = new Set<string>(); for (const g of list) for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (!(r === g.r1 && c === g.c1)) slave.add(r + "_" + c);
  const at = (r: number, c: number) => list.find((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2);
  return { list, slave, at };
}
export function scanMarkers(ws: any) {
  const mm = mergeMap(ws); const out: { name: string; def: string; r: number; c: number; raw: string }[] = [];
  const R = Math.min(ws.rowCount || 0, 300), C = Math.min(ws.columnCount || 0, 60);
  for (let r = 1; r <= R; r++) for (let c = 1; c <= C; c++) {
    if (mm.slave.has(r + "_" + c)) continue; const t = cellText(ws.getCell(r, c).value); if (!t || t.indexOf("{{") < 0) continue;
    MK.lastIndex = 0; let m: RegExpExecArray | null; while ((m = MK.exec(t))) out.push({ name: m[1].trim(), def: m[2] ?? "", r, c, raw: t });
  }
  return out;
}
function b64ToBytes(s: string): Uint8Array { const bin = atob(s); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; }
export function imgSize(b: Uint8Array): { w: number; h: number } {
  if (b[0] === 0x89 && b[1] === 0x50) return { w: ((b[16] << 24) | (b[17] << 16) | (b[18] << 8) | b[19]) >>> 0, h: ((b[20] << 24) | (b[21] << 16) | (b[22] << 8) | b[23]) >>> 0 };
  if (b[0] === 0xFF && b[1] === 0xD8) { let i = 2; while (i + 9 < b.length) { if (b[i] !== 0xFF) { i++; continue; } const m = b[i + 1]; if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) return { h: (b[i + 5] << 8) | b[i + 6], w: (b[i + 7] << 8) | b[i + 8] }; if (m === 0xD8 || m === 0x01 || (m >= 0xD0 && m <= 0xD7)) { i += 2; continue; } const len = (b[i + 2] << 8) | b[i + 3]; i += 2 + len; } }
  return { w: 4, h: 3 };
}
export function fmtNextVisit(s: string): string { if (!s) return ""; const d = new Date(s + "T00:00:00"); if (isNaN(d.getTime())) return ""; return (d.getMonth() + 1) + "月" + d.getDate() + "日（" + ["日", "月", "火", "水", "木", "金", "土"][d.getDay()] + "曜日）"; }

export function fillMarked(wb: any, ws: any, f: any, imgs: { day?: string | null; dinner?: string | null; breakfast?: string | null }) {
  const mm = mergeMap(ws); const marks = scanMarkers(ws);
  const colW = (n: number) => { const w = ws.getColumn(n).width; return (w && w > 0) ? w : 8.43; };
  const rngOf = (r: number, c: number): Rng => mm.at(r, c) || { r1: r, c1: c, r2: r, c2: c };
  const widthOf = (g: Rng) => { let w = 0; for (let c = g.c1; c <= g.c2; c++) w += colW(c); return w; };
  const v1 = f.v1 || {}, v2 = f.v2 || {};
  let appear = String(f.appearance || ""); if (f.special && String(f.special).trim()) appear += (appear ? "\n\n" : "") + "【特記】" + String(f.special).trim();
  const V: Record<string, string> = {
    "利用者名": f.name || "", "利用日時": f.period || "", "測定時間1": v1.time || "", "体温1": v1.temp || "", "血圧1": v1.bp || "", "脈拍1": v1.pulse || "",
    "測定時間2": v2.time || "", "体温2": v2.temp || "", "血圧2": v2.bp || "", "脈拍2": v2.pulse || "", "排尿": f.urine || "", "排便": f.stool || "",
    "投薬": f.medication || "", "担当職員": f.staff || "", "他利用者": f.otherUsers || "", "次回利用予定日": fmtNextVisit(f.nextVisitDate || ""),
    "次回へのお願い": f.nextRequest || "", "記録者": f.writer || "", "日付": f.date || "", "ご様子": appear,
  };
  const names = new Set(marks.map((m) => m.name));
  const cells = new Map<string, { r: number; c: number; raw: string; names: { name: string; def: string }[] }>();
  for (const m of marks) { const k = m.r + "_" + m.c; if (!cells.has(k)) cells.set(k, { r: m.r, c: m.c, raw: m.raw, names: [] }); cells.get(k)!.names.push({ name: m.name, def: m.def }); }
  const hourRow: Record<string, boolean> = {}; for (const h of ["16","17","18","19","20","21","22","23","0","1","2","3","4","5","6","7","8","9"]) hourRow[h] = true;
  const bucket: Record<string, string[]> = {};
  String(f.program || "").split("\n").forEach((l: string) => { l = l.trim(); if (!l) return; const m = l.match(/^(\d{1,2})[:：]/); const h = m ? String(parseInt(m[1], 10)) : null; const key = (h !== null && hourRow[h]) ? h : "16"; (bucket[key] = bucket[key] || []).push(l); });
  const rowH: Record<number, number> = {}; const setMax = (r: number, h: number) => { rowH[r] = Math.max(rowH[r] || 0, h); };
  const imgJobs: { g: Rng; data: string; kind: "full" | "right" | "belowLabel"; w?: number }[] = [];
  const menuFits: { rows: number[]; needPt: number }[] = [];
  let matome: { g: Rng; cell: any } | null = null;
  for (const it of cells.values()) {
    const cell = ws.getCell(it.r, it.c); const first = it.names[0].name; const only = it.names.length === 1 && it.raw.trim() === it.raw.match(/\{\{[^{}]*\}\}/)![0];
    const act = first.match(/^活動(\d{1,2})$/);
    if (act && only) {
      const h = String(parseInt(act[1], 10)); const arr = bucket[h] || []; const r = it.r;
      const sorted = arr.map((l, i) => ({ l, i, t: parseTime(l) })).sort((a, b) => (Number(a.t == null) - Number(b.t == null)) || ((a.t ?? 0) - (b.t ?? 0)) || (a.i - b.i)).map((x) => x.l);
      const content = sorted.join(" / "); const titleStr = content ? detectTitles(content).join("・") : "";
      cell.value = titleStr || null;
      const cg = rngOf(r, it.c + 1); const dCell = ws.getCell(cg.r1, cg.c1);
      if (content) { dCell.value = content; const da = dCell.alignment || {}; dCell.alignment = { wrapText: true, vertical: "top", horizontal: da.horizontal || "left" }; }
      const dFont = (dCell.font && dCell.font.size) ? dCell.font.size : 12;
      if (titleStr) { const bf = cell.font || {}; cell.font = { name: bf.name || "游ゴシック", size: dFont, bold: bf.bold, italic: bf.italic, underline: bf.underline, color: bf.color }; cell.alignment = { wrapText: true, vertical: "top", horizontal: "left" }; }
      if (!content) { setMax(r, 20); continue; }
      const nar = cg.c2 < 11; const dLines = wrapLines(content, widthOf(cg), dFont, nar ? DIV_NARROW : DIV_WIDE);
      const cLines = titleStr ? wrapLines(titleStr, colW(it.c), dFont, DIV_NARROW) : 1;
      setMax(r, Math.max(20, Math.round(Math.max(dLines, cLines) * dFont * LH + 4)));
      continue;
    }
    if ((first === "夕食メニュー" || first === "朝食メニュー") && only) {
      const isD = first === "夕食メニュー"; const g = rngOf(it.r, it.c);
      const body = String((isD ? f.dinnerMenu : f.breakfastMenu) || "") + ((isD ? f.dinnerAmount : f.breakfastAmount) ? "（" + (isD ? f.dinnerAmount : f.breakfastAmount) + "）" : "");
      const eat = isD ? eatLine(f.dinnerMainLeft, f.dinnerSideLeft) : eatLine(f.breakfastMainLeft, f.breakfastSideLeft);
      const photo = (isD ? imgs.dinner : imgs.breakfast) || null; const sepMarker = names.has(isD ? "夕食写真" : "朝食写真");
      const inBox = !!photo && !sepMarker; const boxPx = widthOf(g) * 7 + 5 * (g.c2 - g.c1 + 1);
      const photoW = inBox ? Math.round(Math.min(boxPx * 0.42, 210)) : 0; const textUnits = widthOf(g) - (inBox ? (photoW + 10) / 7 : 0);
      const label = isD ? "夕食メニュー" : "朝食メニュー"; const base = cell.font || {}; const nm = base.name || "游ゴシック";
      const bodyW = inBox && body.trim() ? hardWrap(body, textUnits, 13, DIV_WIDE) : body;
      const runs: any[] = [{ text: label + (bodyW || eat ? "\n" : ""), font: { name: nm, size: 12, bold: true, underline: true, color: base.color } }];
      if (bodyW.trim()) runs.push({ text: bodyW + (eat ? "\n" : ""), font: { name: nm, size: 13, bold: false, underline: false, color: base.color } });
      if (eat) runs.push({ text: eat, font: { name: nm, size: 11, bold: false, underline: false, color: base.color } });
      cell.value = { richText: runs }; const a = cell.alignment || {}; cell.alignment = { wrapText: true, vertical: "top", horizontal: a.horizontal || "left" };
      const bodyLines = body.trim() ? wrapLines(bodyW, textUnits, 13, DIV_WIDE) : 0;
      let need = 12 * LH + bodyLines * 13 * LH + (eat ? 11 * LH : 0) + 8; if (inBox) need = Math.max(need, 82);
      const rows: number[] = []; for (let r = g.r1; r <= g.r2; r++) rows.push(r); menuFits.push({ rows, needPt: need });
      if (inBox) imgJobs.push({ g, data: photo!, kind: "right", w: photoW });
      continue;
    }
    if (first === "ご様子" && only) { matome = { g: rngOf(it.r, it.c), cell }; continue; }
    const pic = it.names.find((x) => x.name === "写真" || x.name === "夕食写真" || x.name === "朝食写真");
    if (pic) {
      const rest = it.raw.replace(MK, (m0: string, n: string, d?: string) => { const k = n.trim(); if (k === "写真" || k === "夕食写真" || k === "朝食写真") return ""; const v = V[k]; return (v != null && String(v).trim() !== "") ? String(v) : (d ?? ""); }).trim();
      cell.value = rest || null;
      const data = pic.name === "写真" ? imgs.day : pic.name === "夕食写真" ? imgs.dinner : imgs.breakfast;
      if (data) imgJobs.push({ g: rngOf(it.r, it.c), data, kind: rest ? "belowLabel" : "full" });
      continue;
    }
    const out = it.raw.replace(MK, (m0: string, n: string, d?: string) => { const v = V[n.trim()]; return (v != null && String(v).trim() !== "") ? String(v) : (d ?? ""); });
    cell.value = out;
  }
  for (const mf of menuFits) { const cur = mf.rows.reduce((s, r) => s + (rowH[r] || ws.getRow(r).height || 18), 0); if (mf.needPt > cur) { const add = (mf.needPt - cur) / mf.rows.length; mf.rows.forEach((r) => { rowH[r] = (rowH[r] || ws.getRow(r).height || 18) + add; }); } }
  Object.keys(rowH).forEach((k) => { const r = Number(k); ws.getRow(r).height = Math.min(240, Math.round(rowH[r])); });
  if (matome) {
    const { g, cell } = matome;
    if (appear) {
      cell.value = appear; const aa = cell.alignment || {}; cell.alignment = { wrapText: true, vertical: "top", horizontal: aa.horizontal || "left" };
      const baseFont = (cell.font && cell.font.size) ? cell.font.size : 14; const wB = widthOf(g); const nrows = g.r2 - g.r1 + 1;
      let budget = 0; for (let rr = g.r1; rr <= g.r2; rr++) budget += (ws.getRow(rr).height || 20);
      let aFont = baseFont; for (let fs = baseFont; fs >= 12; fs--) { aFont = fs; if (wrapLines(appear, wB, fs, DIV_MATOME) * fs * LH + 8 <= budget) break; }
      cell.font = Object.assign({}, cell.font || {}, { size: aFont });
      const needH = wrapLines(appear, wB, aFont, DIV_MATOME) * aFont * LH + 8;
      if (needH > budget) { const per = Math.round(needH / nrows); for (let rr = g.r1; rr <= g.r2; rr++) if (per > (ws.getRow(rr).height || 0)) ws.getRow(rr).height = per; }
    } else cell.value = null;
  }
  const cpx = (c: number) => Math.round(colW(c) * 7 + 5); const rpx = (r: number) => (ws.getRow(r).height || 15) * 4 / 3;
  const placed: any[] = [];
  for (const job of imgJobs) {
    const m = String(job.data).match(/^data:image\/(png|jpe?g|gif);base64,(.*)$/); if (!m) continue;
    const ext = m[1] === "jpg" ? "jpeg" : m[1]; let sz = { w: 4, h: 3 }; try { sz = imgSize(b64ToBytes(m[2].slice(0, 400000))); } catch (_e) { /* keep default */ }
    const g = job.g; let W = 0; for (let c = g.c1; c <= g.c2; c++) W += cpx(c); let H = 0; for (let r = g.r1; r <= g.r2; r++) H += rpx(r);
    let x0 = 4, y0 = 4, w = W - 8, h = H - 8;
    if (job.kind === "right") { w = job.w || Math.round(W * 0.4); x0 = W - w - 5; y0 = 5; h = H - 10; }
    if (job.kind === "belowLabel") { y0 = 24; h = H - y0 - 4; }
    if (w < 10 || h < 10 || !sz.w || !sz.h) continue;
    const s = Math.min(w / sz.w, h / sz.h); const dw = sz.w * s, dh = sz.h * s; const x = x0 + (w - dw) / 2, y = y0 + (h - dh) / 2;
    const anchor = (px: number, py: number) => { let c = g.c1, xx = px; while (c < g.c2 && xx > cpx(c)) { xx -= cpx(c); c++; } let r = g.r1, yy = py; while (r < g.r2 && yy > rpx(r)) { yy -= rpx(r); r++; } return { nativeCol: c - 1, nativeColOff: Math.round(Math.max(0, xx) * 9525), nativeRow: r - 1, nativeRowOff: Math.round(Math.max(0, yy) * 9525) }; };
    const id = wb.addImage({ base64: job.data, extension: ext });
    const tl = anchor(x, y), br = anchor(x + dw, y + dh); ws.addImage(id, { tl, br, editAs: "oneCell" }); placed.push({ kind: job.kind, tl, br });
  }
  return { markers: [...names], placed };
}
