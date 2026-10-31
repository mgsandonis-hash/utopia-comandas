/* Utopía Comandas · conexión con Supabase
   Ofrece a la app la misma forma de trabajar que tenía dentro de Claude
   (colecciones y documentos en tiempo real), pero sobre vuestra propia base de datos. */
(() => {
  const CFG = window.UTOPIA_CONFIG || {};
  const $ = s => document.querySelector(s);
  const clone = v => JSON.parse(JSON.stringify(v));

  // ---------- errores con los mismos códigos que espera la app ----------
  function mapErr(e){
    const msg = String((e && (e.message || e.details)) || e || "");
    if (e && (e.code === "42501" || /row-level security|permission denied|JWT/i.test(msg))) return {code:"invalid_argument", message:msg};
    if (e && e.code === "P0002") return {code:"invalid_argument", message:"no existe"};
    return {code:"unavailable", message:msg};
  }

  // ---------- descargas (Excel) ----------
  const downloads = {
    async save({filename, data}){
      const blob = data instanceof Blob ? data : new Blob([data]);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 10000);
      return {status:"saved"};
    }
  };

  // ---------- base de datos ----------
  let sb = null;
  const cache = new Map(), loaded = new Set(), listeners = new Map(), persistT = {};
  const lsKey = c => "utopia_cache_" + c;
  function getCol(c){
    if (!cache.has(c)){
      const m = new Map();
      try { const raw = JSON.parse(localStorage.getItem(lsKey(c)) || "null"); if (raw) Object.entries(raw).forEach(([k,v]) => m.set(k, v)); } catch(e){}
      cache.set(c, m);
    }
    return cache.get(c);
  }
  function persist(c){
    clearTimeout(persistT[c]);
    persistT[c] = setTimeout(() => { try { localStorage.setItem(lsKey(c), JSON.stringify(Object.fromEntries(getCol(c)))); } catch(e){} }, 400);
  }
  function emit(c){ (listeners.get(c) || new Set()).forEach(fn => { try { fn(); } catch(e){ console.error(e); } }); }
  function addL(c, fn){ if (!listeners.has(c)) listeners.set(c, new Set()); listeners.get(c).add(fn); }
  function removeL(c, fn){ const s = listeners.get(c); if (s) s.delete(fn); }
  const loading = {};
  function load(c){
    if (loading[c]) return loading[c];
    loading[c] = (async () => {
      const {data, error} = await sb.from("docs").select("id,data").eq("col", c);
      if (error) throw mapErr(error);
      const m = new Map(); data.forEach(r => m.set(r.id, r.data));
      cache.set(c, m); loaded.add(c); persist(c); emit(c);
    })().finally(() => { delete loading[c]; });
    return loading[c];
  }
  function reloadAll(){ [...listeners.keys()].forEach(c => load(c).catch(() => {})); }

  const snapDoc = (id, data, fromCache) => ({id, exists: data !== undefined, data: () => data, metadata:{fromCache, hasPendingWrites:false}});
  function query(c, ord, lim){
    const build = fromCache => {
      let rows = [...getCol(c).entries()].map(([id, d]) => ({id, d}));
      if (ord) rows.sort((a, b) => {
        const x = a.d?.[ord.f], y = b.d?.[ord.f];
        if (x === y) return 0; if (x === undefined || x === null) return 1; if (y === undefined || y === null) return -1;
        return (x < y ? -1 : 1) * (ord.dir === "desc" ? -1 : 1);
      }); else rows.sort((a, b) => a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
      if (lim) rows = rows.slice(0, lim);
      const docs = rows.map(r => snapDoc(r.id, r.d, fromCache));
      return {docs, size:docs.length, empty:!docs.length, docChanges:() => [], metadata:{fromCache, hasPendingWrites:false}};
    };
    const q = {
      where(){ return q; },
      orderBy(f, dir = "asc"){ return query(c, {f, dir}, lim); },
      limit(n){ return query(c, ord, n); },
      async get(){ await load(c); return build(false); },
      onSnapshot(next){
        const fn = () => next(build(!loaded.has(c)));
        addL(c, fn);
        if (getCol(c).size) setTimeout(fn, 0);
        load(c).catch(() => {});
        return () => removeL(c, fn);
      },
      doc(id){ return docRef(c, id || newId()); },
      async add(data){ const r = docRef(c, newId()); await r.set(data); return r; },
      path: c
    };
    return q;
  }
  function newId(){
    const a = "abcdefghijklmnopqrstuvwxyz0123456789"; let s = "";
    const r = crypto.getRandomValues(new Uint8Array(20)); r.forEach(b => s += a[b % a.length]); return s;
  }
  function docRef(c, id){
    return {
      id, path: c + "/" + id,
      async get(){
        const {data, error} = await sb.from("docs").select("data").eq("col", c).eq("id", id).maybeSingle();
        if (error) throw mapErr(error);
        return snapDoc(id, data ? data.data : undefined, false);
      },
      async set(d){
        const clean = clone(d);
        const {error} = await sb.from("docs").upsert({col:c, id, data:clean, updated_at:new Date().toISOString()});
        if (error) throw mapErr(error);
        getCol(c).set(id, clean); persist(c); emit(c);
      },
      async update(d){
        const {data, error} = await sb.rpc("merge_doc", {p_col:c, p_id:id, p_patch:clone(d)});
        if (error) throw mapErr(error);
        getCol(c).set(id, data); persist(c); emit(c);
      },
      async delete(){
        const {error} = await sb.from("docs").delete().eq("col", c).eq("id", id);
        if (error) throw mapErr(error);
        getCol(c).delete(id); persist(c); emit(c);
      },
      onSnapshot(next){
        const fn = () => next(snapDoc(id, getCol(c).get(id), !loaded.has(c)));
        addL(c, fn);
        setTimeout(fn, 0);
        load(c).catch(() => {});
        return () => removeL(c, fn);
      }
    };
  }
  const db = {
    collection: path => query(path, null, null),
    doc: path => { const p = path.split("/"); return docRef(p.slice(0, -1).join("/"), p[p.length - 1]); }
  };

  let realtime = null;
  function startRealtime(){
    if (realtime) return;
    realtime = sb.channel("utopia-docs")
      .on("postgres_changes", {event:"*", schema:"public", table:"docs"}, p => {
        const row = (p.new && p.new.col) ? p.new : p.old; if (!row || !row.col) return;
        const m = getCol(row.col);
        if (p.eventType === "DELETE") m.delete(row.id); else m.set(row.id, p.new.data);
        persist(row.col); emit(row.col);
      })
      .subscribe(status => { if (status === "SUBSCRIBED") reloadAll(); });
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") reloadAll(); });
    addEventListener("online", reloadAll);
  }

  // ---------- acceso ----------
  let resolveDb; const dbReady = new Promise(r => resolveDb = r);
  function showLogin(msg){
    const box = $("#login"); box.hidden = false;
    $("#loginErr").textContent = msg || ""; $("#loginErr").hidden = !msg;
  }
  async function afterLogin(session){
    // Comprueba que el correo está en el equipo
    const {data, error} = await sb.rpc("es_equipo");
    if (error && mapErr(error).code === "unavailable"){ $("#login").hidden = true; startRealtime(); resolveDb(db); return; }
    if (!data){ await sb.auth.signOut(); showLogin(`El correo ${session.user.email} no está en el equipo de Utopía. Pide que lo añadan.`); return; }
    $("#login").hidden = true; startRealtime(); resolveDb(db);
  }
  async function init(){
    if (!CFG.supabaseUrl || !CFG.supabaseKey || !window.supabase){ showLogin("La app todavía no está conectada a la base de datos."); return; }
    sb = window.supabase.createClient(CFG.supabaseUrl, CFG.supabaseKey, {auth:{persistSession:true, autoRefreshToken:true}});
    window.utopiaSupabase = sb;
    const {data:{session}} = await sb.auth.getSession();
    if (session){
      // Sin cobertura al abrir: entra con lo guardado en el móvil
      if (!navigator.onLine){ $("#login").hidden = true; startRealtime(); resolveDb(db); return; }
      afterLogin(session); return;
    }
    showLogin();
  }
  document.addEventListener("submit", async e => {
    if (e.target.id !== "loginForm") return;
    e.preventDefault();
    const email = $("#lEmail").value.trim(), pass = $("#lPass").value;
    const mode = e.submitter && e.submitter.dataset.mode;
    const btns = document.querySelectorAll("#loginForm button"); btns.forEach(b => b.disabled = true);
    try {
      if (mode === "signup"){
        const {data, error} = await sb.auth.signUp({email, password:pass});
        if (error) throw error;
        if (!data.session){ showLogin("Te hemos mandado un correo para confirmar la cuenta. Ábrelo y vuelve aquí para entrar."); return; }
        await afterLogin(data.session);
      } else {
        const {data, error} = await sb.auth.signInWithPassword({email, password:pass});
        if (error) throw error;
        await afterLogin(data.session);
      }
    } catch(err){
      const m = String(err.message || err);
      showLogin(/Invalid login/i.test(m) ? "Correo o contraseña incorrectos." : /already registered/i.test(m) ? "Ese correo ya tiene cuenta: pulsa Entrar." : /Password should/i.test(m) ? "La contraseña debe tener al menos 6 caracteres." : /fetch/i.test(m) ? "Sin conexión. Prueba cuando tengas cobertura." : m);
    } finally { btns.forEach(b => b.disabled = false); }
  });
  document.addEventListener("click", async e => {
    if (!e.target.closest("#logoutBtn")) return;
    if (!sb) return;
    await sb.auth.signOut();
    Object.keys(localStorage).filter(k => k.startsWith("utopia_cache_")).forEach(k => localStorage.removeItem(k));
    location.reload();
  });

  window.claude = {
    use: async name => {
      if (name === "db") return dbReady;
      if (name === "user") return {can: async () => true, canEdit: async () => true, isOwner: async () => true};
      if (name === "downloads") return downloads;
      return null;
    }
  };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
