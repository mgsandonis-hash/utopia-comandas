/* Utopía Comandas · traer la asistencia de WodBuster
   Se ejecuta en TU navegador, con WodBuster abierto y tu sesión iniciada (no guarda ni envía tu contraseña).
   Lee cuántas personas hubo en cada clase de los últimos días y te lleva a la app para que lo confirmes.
   Solo sale de WodBuster el número de personas por clase: ningún nombre. */
(function(){
  "use strict";
  var me = document.currentScript && document.currentScript.src || "";
  var APP = me ? me.replace(/wb\.js.*$/, "") : "https://mgsandonis-hash.github.io/utopia-comandas/";
  var DIAS = 15;                                  // se mira hasta 15 días atrás, pero solo sábados y domingos (los días que abre la cafetería)
  var old = document.getElementById("utopia-wb"); if (old) old.remove();
  var box = document.createElement("div"); box.id = "utopia-wb";
  box.style.cssText = "position:fixed;left:0;right:0;bottom:0;z-index:2147483647;background:#f8f2e8;color:#3b0a12;font:16px/1.4 system-ui,sans-serif;padding:16px 16px calc(16px + env(safe-area-inset-bottom,0px));border-top:3px solid #5e0f1c;box-shadow:0 -8px 30px rgba(0,0,0,.35);max-height:80vh;overflow:auto";
  document.body.appendChild(box);
  function esc(s){ return String(s).replace(/[&<>"]/g, function(c){ return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]; }); }
  function show(html){ box.innerHTML = '<div style="display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:8px"><b style="font-size:18px">Utopía · asistencia</b><button type="button" id="utopia-wb-x" style="font:inherit;background:none;border:1px solid #5e0f1c;color:#5e0f1c;border-radius:8px;padding:4px 10px">Cerrar</button></div>' + html; document.getElementById("utopia-wb-x").onclick = function(){ box.remove(); }; }
  function fail(msg, extra){ show('<p style="margin:0 0 8px">' + esc(msg) + '</p>' + (extra ? '<p style="margin:0;font-size:12px;opacity:.75;word-break:break-all">Detalle para Claude: ' + esc(extra) + '</p>' : "")); }

  if (!/(^|\.)wodbuster\.com$/.test(location.hostname) && location.hostname !== "localhost"){
    return fail("Esto solo funciona con WodBuster abierto. Entra en utopia.wodbuster.com, inicia sesión y vuelve a pulsar el marcador.");
  }
  show('<p style="margin:0" id="utopia-wb-p">Leyendo las clases…</p>');
  var pad = function(n){ return (n < 10 ? "0" : "") + n; };
  var hhmm = function(s){ var m = /(\d{1,2}):(\d{2})/.exec(String(s || "")); return m ? pad(Number(m[1])) + ":" + m[2] : ""; };
  var sleep = function(ms){ return new Promise(function(r){ setTimeout(r, ms); }); };
  var RE_IDU = /Crossfit\.Init\(\s*(?:[^,]*,\s*){5}'([^']+)'/;

  function findIdu(){
    var m = RE_IDU.exec(document.documentElement.innerHTML);
    if (m) return Promise.resolve(m[1]);
    return fetch("/athlete/reservas.aspx", {credentials:"include"}).then(function(r){ return r.text(); }).then(function(t){ var k = RE_IDU.exec(t); return k ? k[1] : ""; }).catch(function(){ return ""; });
  }
  function loadDay(date, idu){
    var ticks = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) / 1000;
    var url = "/athlete/handlers/LoadClass.ashx?ticks=" + ticks + (idu ? "&idu=" + encodeURIComponent(idu) : "");
    return fetch(url, {credentials:"include", headers:{"X-Requested-With":"XMLHttpRequest", "Accept":"application/json, text/javascript, */*; q=0.01"}})
      .then(function(r){ return r.text().then(function(t){ return {status:r.status, url:r.url, text:t}; }); });
  }
  function parse(res, date, now){
    var j; try { j = JSON.parse(res.text); } catch(e){ return {err:"no-json", det:res.status + " " + res.text.slice(0, 140).replace(/\s+/g, " ")}; }
    var out = {}, isToday = date.toDateString() === now.toDateString(), nowMin = now.getHours()*60 + now.getMinutes();
    (j && Array.isArray(j.Data) ? j.Data : []).forEach(function(b){
      (b && Array.isArray(b.Valores) ? b.Valores : []).forEach(function(v){
        var c = v && v.Valor; if (!c) return;
        var h = hhmm(c.HoraComienzo) || hhmm(b.Hora); if (!h) return;
        var n = Array.isArray(c.AtletasEntrenando) ? c.AtletasEntrenando.length : 0; if (!n) return;
        if (isToday && Number(h.slice(0,2))*60 + Number(h.slice(3)) > nowMin) return;      // hoy: solo las clases que ya han empezado
        var name = String(c.Nombre || "").replace(/\s+/g, " ").trim().slice(0, 30), k = h + "|" + name;
        out[k] = (out[k] || 0) + n;
      });
    });
    var ses = Object.keys(out).sort().map(function(k){ var p = k.split("|"); return [p[0], p.slice(1).join("|"), out[k]]; });
    return {ses:ses, keys:j ? Object.keys(j).slice(0, 8).join(",") : ""};
  }

  (async function(){
    var now = new Date(), idu = await findIdu(), days = [], lastDet = "", noJson = 0;
    var fechas = [];
    for (var i = DIAS - 1; i >= 0; i--){ var f = new Date(now.getFullYear(), now.getMonth(), now.getDate() - i); if (f.getDay() === 6 || f.getDay() === 0) fechas.push(f); }
    for (var q = 0; q < fechas.length; q++){
      var d = fechas[q];
      var p = document.getElementById("utopia-wb-p"); if (p) p.textContent = "Leyendo sábados y domingos… " + (q + 1) + " de " + fechas.length;
      var res; try { res = await loadDay(d, idu); } catch(e){ lastDet = "red: " + e.message; continue; }
      var r = parse(res, d, now);
      if (r.err){ noJson++; lastDet = r.det + (idu ? "" : " (sin idu)"); }
      else { lastDet = lastDet || ("ok, campos: " + r.keys); if (r.ses.length) days.push([d.getFullYear() + "-" + pad(d.getMonth()+1) + "-" + pad(d.getDate()), r.ses]); }
      await sleep(350);
    }
    if (!days.length){
      if (noJson) return fail("WodBuster no ha devuelto las clases. Lo más probable es que no hayas iniciado sesión en esta pestaña: entra con tu usuario y vuelve a pulsar el marcador.", lastDet);
      return fail("No hay clases con gente apuntada en los sábados y domingos de las dos últimas semanas.", lastDet);
    }
    var json = JSON.stringify({v:1, d:days});
    var b64 = btoa(unescape(encodeURIComponent(json))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
    var fmt = function(k){ return new Date(k + "T12:00").toLocaleDateString("es-ES", {weekday:"short", day:"numeric", month:"short"}); };
    show('<ul style="margin:0 0 12px;padding-left:18px">' + days.map(function(x){ return "<li>" + esc(fmt(x[0])) + ": " + x[1].length + (x[1].length === 1 ? " clase · " : " clases · ") + x[1].reduce(function(a, s){ return a + s[2]; }, 0) + " personas</li>"; }).join("") + "</ul>" +
      '<a href="' + APP + "#wb=" + b64 + '" style="display:block;text-align:center;background:#5e0f1c;color:#fff;text-decoration:none;font-weight:700;padding:14px;border-radius:12px">Enviar a la app de Utopía</a>' +
      '<p style="margin:8px 0 0;font-size:13px;opacity:.8">Solo se envía cuántas personas hubo en cada clase, sin nombres. En la app podrás revisarlo antes de guardar.</p>');
  })().catch(function(e){ fail("Algo ha fallado leyendo WodBuster.", String(e && e.message || e)); });
})();
