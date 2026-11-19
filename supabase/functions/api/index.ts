// ระบบเบิกพัสดุห้องยา — Backend: Supabase Edge Function + Postgres
// Design & Developed by ภก.วทัญญู โตเจริญนิรัติศัย
// โปรโตคอลเดียวกับ Apps Script เดิม: POST text/plain JSON {a:'login'|'pull'|'write'|'acquire'|'changePw', ...}
import postgres from 'npm:postgres@3.4.4';

const sql = postgres(Deno.env.get('SUPABASE_DB_URL')!, { max: 3, prepare: false, idle_timeout: 20 });

const COLS = ['staff', 'items', 'units', 'requests', 'cycles', 'locks'];
const ADMIN_ONLY: Record<string, number> = { staff: 1, items: 1, units: 1, cycles: 1, locks: 1 };
const DEFAULT_PW = '12345678';
const TOKEN_MS = 12 * 3600 * 1000;

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': '*',
  'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
};
const out = (o: unknown) =>
  new Response(JSON.stringify(o), { headers: { ...CORS, 'Content-Type': 'application/json; charset=utf-8' } });

const enc = new TextEncoder();
const hex = (b: ArrayBuffer) => [...new Uint8Array(b)].map((x) => x.toString(16).padStart(2, '0')).join('');
const sha = async (s: string) => hex(await crypto.subtle.digest('SHA-256', enc.encode(s)));
let SECRET = '';
let KEY: CryptoKey | null = null;
async function secret() {
  if (!SECRET) {
    const r = await sql`select v from kv where k='SECRET'`;
    SECRET = r[0].v;
  }
  return SECRET;
}
async function sign(s: string) {
  if (!KEY) KEY = await crypto.subtle.importKey('raw', enc.encode(await secret()), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  return hex(await crypto.subtle.sign('HMAC', KEY, enc.encode(s)));
}
async function fp(h: string | null) { return h ? (await sha('fp|' + (await secret()) + '|' + h)).slice(0, 12) : null; }
async function pub(id: string, d: any) {
  const o: any = {};
  for (const k in d) if (k !== 'passHash' && k !== '__id') o[k] = d[k];
  o.id = id; o.passHash = await fp(d.passHash || null);
  return o;
}
async function mkToken(st: any) {
  const ts = Date.now(), id = String(st.empId);
  return ts + '.' + id + '.' + (await sign(ts + '|' + id + '|' + (st.passHash || 'default')));
}
async function getDoc(q: any, c: string, id: string) {
  const r = await q`select data from docs where coll=${c} and id=${id}`;
  return r.length ? r[0].data : null;
}
async function auth(token: string) {
  const t = String(token || '').split('.');
  if (t.length !== 3) return null;
  if (Date.now() - Number(t[0]) > TOKEN_MS) return null;
  const st = await getDoc(sql, 'staff', t[1]);
  if (!st) return null;
  if ((await sign(t[0] + '|' + t[1] + '|' + (st.passHash || 'default'))) !== t[2]) return null;
  st.empId = String(st.empId || t[1]);
  return st;
}
async function bump(q: any, c: string) {
  await q`insert into colver(coll,ver) values (${c},1) on conflict (coll) do update set ver = colver.ver + 1`;
}
async function putDoc(q: any, c: string, id: string, data: any) {
  const j = JSON.stringify(data);
  if (j.length > 400000) throw new Error('too_large');
  await q`insert into docs(coll,id,data,ver) values (${c},${id},${j}::text::jsonb,${Date.now()})
          on conflict (coll,id) do update set data=excluded.data, ver=excluded.ver`;
}

async function login(req: any) {
  const id = String(req.id || '').replace(/\D/g, ''), h = String(req.h || '');
  const now = Date.now();
  const f = await sql`select n, until_ms from fails where id=${id}`;
  if (f.length && Number(f[0].until_ms) > now && f[0].n >= 8) return { ok: false, code: 'locked' };
  const st = await getDoc(sql, 'staff', id);
  let good = false;
  if (st) good = st.passHash ? st.passHash === h : h === (await sha('ch9sup|' + id + '|' + DEFAULT_PW));
  if (!good) {
    const n = f.length && Number(f[0].until_ms) > now ? f[0].n + 1 : 1;
    await sql`insert into fails(id,n,until_ms) values (${id},${n},${now + 600000})
              on conflict (id) do update set n=excluded.n, until_ms=excluded.until_ms`;
    return { ok: false, code: st ? 'bad_pw' : 'no_user' };
  }
  if (f.length) await sql`delete from fails where id=${id}`;
  st.empId = id;
  return { ok: true, token: await mkToken(st), user: await pub(id, st) };
}

async function pull(req: any) {
  const vers = req.vers || {};
  const vr = await sql`select coll, ver from colver`;
  const cols: any = {};
  for (const { coll, ver } of vr) {
    if (!COLS.includes(coll)) continue;
    const v = Number(ver);
    if (vers[coll] === v) continue;
    const rows = await sql`select id, data from docs where coll=${coll}`;
    const docs = [];
    for (const r of rows) docs.push(coll === 'staff' ? await pub(r.id, r.data) : { id: r.id, ...r.data });
    cols[coll] = { ver: v, docs };
  }
  return { ok: true, cols, now: Date.now() };
}

const LOCK_KEY = 90210;
async function tx<T>(fn: (q: any) => Promise<T>): Promise<T> {
  return await sql.begin(async (q: any) => {
    await q`select pg_advisory_xact_lock(${LOCK_KEY})`;
    return await fn(q);
  }) as T;
}

async function changePw(me: any, req: any) {
  const h = String(req.h || '');
  if (!/^[0-9a-f]{64}$/.test(h)) return { ok: false, code: 'bad_request' };
  return await tx(async (q) => {
    const st = await getDoc(q, 'staff', me.empId);
    st.passHash = h; st.pwChangedAt = Date.now();
    await putDoc(q, 'staff', me.empId, st); await bump(q, 'staff');
    return { ok: true, token: await mkToken(st), user: await pub(me.empId, st) };
  });
}

async function adminsExcept(q: any, id: string) {
  const r = await q`select id from docs where coll='staff' and data->>'isAdmin'='true' and id<>${id}`;
  return r.length;
}

async function write(me: any, req: any) {
  const ops = req.ops || [];
  return await tx(async (q) => {
    const touched: Record<string, number> = {};
    for (const o of ops) {
      const parts = String(o.path || '').split('/'), c = parts[0], id = parts[1];
      if (!COLS.includes(c) || !id || parts.length !== 2) throw new Error('bad_path');
      if (ADMIN_ONLY[c] && !me.isAdmin) throw new Error('forbidden');
      const cur = await getDoc(q, c, id), data = o.data || {};
      if (o.op === 'delete') {
        if (c === 'staff' && cur && cur.isAdmin && !(await adminsExcept(q, id))) throw new Error('last_admin');
        await q`delete from docs where coll=${c} and id=${id}`;
      } else {
        let nd: any;
        if (o.op === 'update') { if (!cur) throw new Error('not_found'); nd = { ...cur, ...data }; }
        else nd = { ...data };
        if (c === 'staff') {
          const keep = cur ? cur.passHash : null;
          if (o.op === 'update' && 'passHash' in data && data.passHash === null) nd.passHash = null;
          else nd.passHash = keep || null;
          if (o.op === 'set' && !cur) nd.passHash = null;
          if (id === me.empId && !nd.isAdmin && me.isAdmin && !(await adminsExcept(q, id))) throw new Error('last_admin');
          nd.empId = id;
        }
        await putDoc(q, c, id, nd);
      }
      touched[c] = 1;
    }
    for (const c of Object.keys(touched)) await bump(q, c);
    return { ok: true };
  });
}

async function acquire(me: any, req: any) {
  if (!me.isAdmin) throw new Error('forbidden');
  const id = String(req.id || ''), holder = String(req.holder || me.empId);
  const ttl = Math.min(Number(req.ttlMs) || 20000, 120000);
  if (!id) throw new Error('bad_path');
  return await tx(async (q) => {
    const cur = await getDoc(q, 'locks', id), now = Date.now();
    if (cur && cur.until > now && cur.holderId !== holder) return { ok: true, acquired: false };
    const d = { ...(req.data || {}), holderId: holder, until: now + ttl };
    await putDoc(q, 'locks', id, d); await bump(q, 'locks');
    return { ok: true, acquired: true };
  });
}

Deno.serve(async (r: Request) => {
  if (r.method === 'OPTIONS') return new Response(null, { headers: CORS });
  if (r.method === 'GET') return out({ ok: true, name: 'ch9-supply', time: Date.now() });
  let req: any;
  try { req = JSON.parse(await r.text()); } catch { return out({ ok: false, code: 'bad_request' }); }
  try {
    if (req.a === 'login') return out(await login(req));
    if (req.a === 'ping') { await sql`select 1`; return out({ ok: true }); }
    const me = await auth(req.token);
    if (!me) return out({ ok: false, code: 'auth' });
    if (req.a === 'pull') return out(await pull(req));
    if (req.a === 'changePw') return out(await changePw(me, req));
    if (req.a === 'write') return out(await write(me, req));
    if (req.a === 'acquire') return out(await acquire(me, req));
    return out({ ok: false, code: 'bad_request' });
  } catch (e) {
    return out({ ok: false, code: String((e as Error)?.message || e) });
  }
});
