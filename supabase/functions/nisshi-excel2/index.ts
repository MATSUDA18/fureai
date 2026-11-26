// 利用日誌Excel（目印方式・写真対応）とひな形の管理。
// 従来の nisshi-excel には一切触れない。ひな形は専用の場所（templates/nisshi_v2/）に置き、
// ここに有効なひな形が無いときは 409 {legacy:true} を返してアプリ側が従来の出力を使う。
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import ExcelJS from "npm:exceljs@4.4.0";
import { createClient } from "npm:@supabase/supabase-js@2";
import { fillMarked, scanMarkers, KNOWN_MARKERS } from "./core.ts";

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const BUCKET = "templates";
const ACTIVE = "nisshi_v2/active.xlsx";
const META = "nisshi_v2/active.json";
const INDEX = "nisshi_v2/index.json";
const VER = "nisshi_v2/versions";
const LEGACY = "nisshi_template.xlsx";
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const json = (o: unknown, status = 200) => new Response(JSON.stringify(o), { status, headers: { ...cors, "Content-Type": "application/json" } });

async function sha256hex(t: string) { const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(t)); return Array.from(new Uint8Array(b)).map((x) => x.toString(16).padStart(2, "0")).join(""); }
function b64ToBytes(s: string) { const bin = atob(s); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; }

serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors });
  const sb = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
  const dl = async (path: string): Promise<Uint8Array | null> => { const r = await sb.storage.from(BUCKET).download(path); if (r.error || !r.data) return null; return new Uint8Array(await r.data.arrayBuffer()); };
  const up = async (path: string, bytes: Uint8Array, type: string) => { const r = await sb.storage.from(BUCKET).upload(path, bytes, { contentType: type, upsert: true }); if (r.error) throw new Error("保存に失敗：" + r.error.message); };
  const readJson = async (path: string, def: any) => { const b = await dl(path); if (!b) return def; try { return JSON.parse(new TextDecoder().decode(b)); } catch (_e) { return def; } };
  const isAdmin = async (pw: string) => {
    if (!pw) return false; const h = await sha256hex(pw);
    const r = await sb.from("app_state").select("data").eq("id", 1).maybeSingle(); const d = (r.data && (r.data as any).data) || {};
    if (d.master && h === d.master) return true;
    return (d.accounts || []).some((a: any) => a && a.isAdmin && a.hash === h);
  };
  const resolveImg = async (p: any): Promise<string | null> => {
    if (!p) return null; if (typeof p === "string" && p.startsWith("data:image")) return p;
    if (p.data && String(p.data).startsWith("data:image")) return String(p.data);
    if (p.photoId) { const r = await sb.from("photos").select("image").eq("id", String(p.photoId)).maybeSingle(); const im = r.data && (r.data as any).image; if (typeof im === "string" && im.startsWith("data:image")) return im; }
    return null;
  };
  try {
    const f = await req.json();
    const action = f.action || "render";

    if (action === "status") {
      const active = await dl(ACTIVE); const meta = await readJson(META, null); const index = await readJson(INDEX, {});
      const ls = await sb.storage.from(BUCKET).list(VER, { limit: 100, sortBy: { column: "created_at", order: "desc" } });
      const versions = (ls.data || []).filter((o: any) => o.name && o.name.endsWith(".xlsx")).map((o: any) => { const path = VER + "/" + o.name; return { path, name: (index[path] && index[path].name) || o.name, at: (index[path] && index[path].at) || o.created_at, by: (index[path] && index[path].by) || "", markers: (index[path] && index[path].markers) || [] }; });
      return json({ active: !!active, meta, versions, known: KNOWN_MARKERS });
    }
    if (action === "download") {
      const path = f.path === "active" ? ACTIVE : f.path === "legacy" ? LEGACY : String(f.path || "");
      if (!(path === ACTIVE || path === LEGACY || path.startsWith(VER + "/"))) return json({ error: "path" }, 400);
      const b = await dl(path); if (!b) return json({ error: "見つかりません" }, 404);
      return new Response(b, { headers: { ...cors, "Content-Type": XLSX_TYPE } });
    }
    if (action === "upload" || action === "activate" || action === "deactivate" || action === "delete") {
      if (!(await isAdmin(String(f.pw || "")))) return json({ error: "管理者のパスワードが違います" }, 403);
      if (action === "upload") {
        const bytes = b64ToBytes(String(f.b64 || "")); if (bytes.length < 100 || bytes.length > 5_000_000) return json({ error: "ファイルの大きさが不正です" }, 400);
        const wb = new ExcelJS.Workbook(); try { await wb.xlsx.load(bytes); } catch (_e) { return json({ error: "Excelファイル（.xlsx）として読み込めませんでした" }, 400); }
        const ws = wb.worksheets[0]; if (!ws) return json({ error: "シートがありません" }, 400);
        const found = [...new Set(scanMarkers(ws).map((m) => m.name))];
        if (!found.length) return json({ error: "目印（{{利用者名}} など）が1つも見つかりません" }, 400);
        const unknown = found.filter((n) => !KNOWN_MARKERS.includes(n));
        const path = VER + "/" + Date.now() + "_" + Math.random().toString(36).slice(2, 8) + ".xlsx";
        await up(path, bytes, XLSX_TYPE);
        const index = await readJson(INDEX, {}); index[path] = { name: String(f.filename || "ひな形.xlsx").slice(0, 120), at: new Date().toISOString(), by: String(f.by || "").slice(0, 40), markers: found };
        await up(INDEX, new TextEncoder().encode(JSON.stringify(index)), "application/json");
        return json({ ok: true, path, markers: found, unknown, missing: KNOWN_MARKERS.filter((n) => !found.includes(n) && !/^活動|^夕食写真|^朝食写真|^日付|^他利用者/.test(n)) });
      }
      if (action === "activate") {
        const path = String(f.path || ""); if (!path.startsWith(VER + "/")) return json({ error: "path" }, 400);
        const b = await dl(path); if (!b) return json({ error: "見つかりません" }, 404);
        await up(ACTIVE, b, XLSX_TYPE);
        const index = await readJson(INDEX, {});
        await up(META, new TextEncoder().encode(JSON.stringify({ path, name: (index[path] && index[path].name) || path, at: new Date().toISOString(), by: String(f.by || "").slice(0, 40) })), "application/json");
        return json({ ok: true });
      }
      if (action === "deactivate") { await sb.storage.from(BUCKET).remove([ACTIVE, META]); return json({ ok: true }); }
      if (action === "delete") {
        const path = String(f.path || ""); if (!path.startsWith(VER + "/")) return json({ error: "path" }, 400);
        const meta = await readJson(META, null); if (meta && meta.path === path) return json({ error: "使用中のひな形は削除できません" }, 400);
        await sb.storage.from(BUCKET).remove([path]); const index = await readJson(INDEX, {}); delete index[path];
        await up(INDEX, new TextEncoder().encode(JSON.stringify(index)), "application/json"); return json({ ok: true });
      }
    }
    // ===== render =====
    let path = ACTIVE; if (f.templatePath) { const p = String(f.templatePath); if (!p.startsWith(VER + "/")) return json({ error: "path" }, 400); path = p; }
    const tb = await dl(path); if (!tb) return json({ legacy: true }, 409);
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(tb); const ws = wb.worksheets[0];
    const imgs = { day: await resolveImg(f.dayPhoto), dinner: await resolveImg(f.dinnerPhoto), breakfast: await resolveImg(f.breakfastPhoto) };
    fillMarked(wb, ws, f, imgs);
    ws.pageSetup = { ...(ws.pageSetup || {}), orientation: "portrait", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true } as any;
    const out = await wb.xlsx.writeBuffer();
    return new Response(out, { headers: { ...cors, "Content-Type": XLSX_TYPE } });
  } catch (e) {
    return json({ error: String(e && (e as any).message || e) }, 500);
  }
});
