/* =====================================================================
   GESTIÓN PREVENTIVA EJECUTADA EN TERRENO
   Lee el informe completo (PDF) en el navegador, propone un resumen ejecutivo por bloques
   y las fotos de cada bloque; todo es editable antes de generar las láminas.
   Nada se sube a ningún servidor (pdf.js corre aquí mismo).
   ===================================================================== */
const PDFJS_VER = "4.10.38";
const PDFJS_URL = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VER}/build/pdf.min.mjs`;
const PDFJS_WORKER = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VER}/build/pdf.worker.min.mjs`;

const PREV = {
  incluir: false,
  lema: "GESTIÓN EN SEGURIDAD",
  titulo: "Gestión preventiva ejecutada en terreno",
  subtitulo: "Acciones desarrolladas durante el período",
  bloques: [],      // {titulo, intro, items(texto, uno por línea), claves:[...], fotos:[id,...]}
  fotos: [],        // {id, src(dataURL), w, h, pag}
  lineas: 0
};

/* Reglas de clasificación (palabras clave por bloque). */
const PREV_REGLAS = [
  { titulo: "Coordinación", max: 5, tope: 140,
    re: /coordin|reuni[oó]n|tfds|charla|campa[nñ]a|an[aá]lisis|reflexi[oó]n|comit[eé]|capacitaci[oó]n|planific|acordamos|se defini|foco de cada|entrega de foco|difusi[oó]n/i },
  { titulo: "Procesos y rol", max: 4, tope: 140,
    re: /proceso|tarea cr[ií]tica|repetitiva|\brol\b|jefe de terreno|auditor[ií]a|verificaci[oó]n|confirmaci[oó]n|\bedc\b|\bartp\b|\bpts\b|\bcdp\b|cumplimiento/i },
  { titulo: "Ejecución", max: 6, tope: 150,
    re: /terreno|chequeo|inspecci[oó]n|izaje|energ[ií]a|aislamiento|emergencia|suspel|control(es)?\b|observaci[oó]n|ejecut|se verific|se revis|se realiz|estabilidad|disciplina|bloqueo|andamio|excavaci[oó]n/i },
  { titulo: "Visita gerencial", max: 2, tope: 220,
    re: /gerente|gerencia|gemba|visita|superintendente|subgerente|direcci[oó]n/i }
];

/** Importa pdf.js solo cuando se necesita (el worker se sirve como blob para evitar el bloqueo entre orígenes). */
let _pdfjs = null;
async function cargarPdfjs() {
  if (_pdfjs) return _pdfjs;
  const lib = await import(PDFJS_URL);
  const r = await fetch(PDFJS_WORKER);
  if (!r.ok) throw new Error("No se pudo descargar el lector de PDF (pdf.js).");
  const blob = new Blob([await r.text()], { type: "text/javascript" });
  lib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
  _pdfjs = lib;
  return lib;
}

/** Texto de una página agrupado en líneas (por posición vertical). */
async function textoDePagina(pagina) {
  const tc = await pagina.getTextContent();
  const filas = [];
  tc.items.forEach(it => {
    if (!it.str || !it.str.trim()) return;
    const y = Math.round(it.transform[5] / 3), x = it.transform[4];
    let f = filas.find(q => Math.abs(q.y - y) <= 1);
    if (!f) { f = { y, partes: [] }; filas.push(f); }
    f.partes.push({ x, s: it.str });
  });
  filas.sort((a, b) => b.y - a.y);
  return filas.map(f => f.partes.sort((a, b) => a.x - b.x).map(p => p.s).join(" ").replace(/\s+/g, " ").trim()).filter(Boolean);
}

/** Convierte un objeto de imagen de pdf.js en un JPEG (dataURL) de ancho máx. 1000 px. */
function imagenAJpeg(img) {
  const w = img.width, h = img.height;
  const cv = document.createElement("canvas");
  cv.width = w; cv.height = h;
  const ctx = cv.getContext("2d");
  if (img.bitmap) {
    ctx.drawImage(img.bitmap, 0, 0);
  } else if (img.data) {
    const rgba = ctx.createImageData(w, h);
    const d = img.data, o = rgba.data;
    if (d.length === w * h * 4) o.set(d);
    else if (d.length === w * h * 3) { for (let i = 0, j = 0; i < d.length; i += 3, j += 4) { o[j] = d[i]; o[j + 1] = d[i + 1]; o[j + 2] = d[i + 2]; o[j + 3] = 255; } }
    else if (d.length === w * h) { for (let i = 0, j = 0; i < d.length; i++, j += 4) { o[j] = o[j + 1] = o[j + 2] = d[i]; o[j + 3] = 255; } }
    else return null;
    ctx.putImageData(rgba, 0, 0);
  } else return null;
  let dest = cv;
  if (w > 1000) {
    dest = document.createElement("canvas");
    dest.width = 1000; dest.height = Math.round(h * 1000 / w);
    const c2 = dest.getContext("2d");
    c2.fillStyle = "#fff"; c2.fillRect(0, 0, dest.width, dest.height);
    c2.drawImage(cv, 0, 0, dest.width, dest.height);
  }
  return { src: dest.toDataURL("image/jpeg", 0.82), w: dest.width, h: dest.height };
}

function objetoImagen(pagina, nombre) {
  return new Promise(res => {
    let listo = false;
    const t = setTimeout(() => { if (!listo) res(null); }, 4000);
    try { pagina.objs.get(nombre, o => { listo = true; clearTimeout(t); res(o); }); }
    catch (e) { clearTimeout(t); res(null); }
  });
}

async function imagenesDePagina(lib, pagina, numPag) {
  const ops = await pagina.getOperatorList();
  const OPS = lib.OPS, salida = [];
  for (let i = 0; i < ops.fnArray.length; i++) {
    const f = ops.fnArray[i];
    if (f !== OPS.paintImageXObject && f !== OPS.paintInlineImageXObject) continue;
    let obj = null;
    if (f === OPS.paintImageXObject) obj = await objetoImagen(pagina, ops.argsArray[i][0]);
    else obj = ops.argsArray[i][0];
    if (!obj || !obj.width || obj.width < 250 || obj.height < 170) continue;
    const ratio = obj.width / obj.height;
    if (ratio < 0.4 || ratio > 3.5) continue;
    try { const j = imagenAJpeg(obj); if (j) salida.push({ ...j, pag: numPag }); } catch (e) { /* imagen ilegible */ }
  }
  return salida;
}

/** Resumen ejecutivo por reglas: cada frase va al bloque donde más calza (los títulos del informe ayudan) y se dejan las más informativas. */
function resumirPorReglas(paginas) {
  // Líneas que se repiten en muchas páginas (encabezados/pies) se descartan.
  const cuenta = {};
  paginas.forEach(p => new Set(p.lineas.map(l => l.toLowerCase())).forEach(l => { cuenta[l] = (cuenta[l] || 0) + 1; }));
  const repetida = l => paginas.length >= 4 && cuenta[l.toLowerCase()] >= Math.max(3, paginas.length * 0.4);
  const esTitulo = l => l.length <= 60 && !/[.;!?]$/.test(l) && l.split(" ").length <= 7;

  const frases = [];
  let pista = -1;                         // bloque sugerido por el último título encontrado
  const vaciar = (buf, pag) => {
    const texto = buf.join(" ").replace(/\s+/g, " ");
    texto.split(/(?<=[.;!?])\s+|\s[•●▪■]\s/).forEach(s => {
      s = s.replace(/^[\s•●▪■\-–*\d.)]+/, "").trim();
      if (s.length < 25 || s.length > 400) return;
      if (/^(p[aá]gina|informe|fecha|n[°º]|tabla|figura|gr[aá]fico)\b/i.test(s)) return;
      if ((s.match(/[A-Za-zÁÉÍÓÚáéíóúñÑ]/g) || []).length < s.length * 0.55) return;   // mucha cifra/símbolo
      frases.push({ s, pag, pista });
    });
    buf.length = 0;
  };
  paginas.forEach(p => {
    const buf = [];
    p.lineas.filter(l => !repetida(l)).forEach(l => {
      if (esTitulo(l)) {
        vaciar(buf, p.num);
        const i = PREV_REGLAS.findIndex(r => r.re.test(l));
        if (i >= 0) pista = i;
      } else buf.push(l);
    });
    vaciar(buf, p.num);
  });

  const asignadas = PREV_REGLAS.map(() => []);
  frases.forEach(f => {
    let mejor = -1, mp = 0;
    PREV_REGLAS.forEach((r, i) => {
      const coinc = (f.s.match(new RegExp(r.re.source, "gi")) || []).length;
      let pt = coinc * 2 + (f.pista === i ? 3 : 0) + (i === 3 && /gerente|gemba/i.test(f.s) ? 4 : 0);
      if (coinc === 0 && i !== f.pista) pt = 0;
      if (pt > mp) { mp = pt; mejor = i; }
    });
    if (mejor >= 0 && mp >= 2) {
      let puntos = mp + (/\d/.test(f.s) ? 1 : 0) + (/^[A-ZÁÉÍÓÚÑ]/.test(f.s) ? 0.5 : 0);
      if (f.s.length > PREV_REGLAS[mejor].tope * 1.8) puntos -= 2;
      asignadas[mejor].push({ ...f, puntos });
    }
  });

  return PREV_REGLAS.map((r, i) => {
    const cand = asignadas[i].sort((a, b) => b.puntos - a.puntos).slice(0, r.max).sort((a, b) => a.pag - b.pag);
    return { titulo: r.titulo, intro: "", items: cand.map(c => acortar(c.s, r.tope)).join("\n"),
             pags: cand.map(c => c.pag), fotos: [], re: r.re };
  });
}

function acortar(s, tope) {
  s = s.replace(/\s+/g, " ").trim();
  if (s.length <= tope) return s.replace(/[.;,]$/, "") + ".";
  const corte = s.slice(0, tope);
  const k = Math.max(corte.lastIndexOf(", "), corte.lastIndexOf(" y "), corte.lastIndexOf(" "));
  return corte.slice(0, k > tope * 0.5 ? k : tope).replace(/[,;:\s]+$/, "") + "…";
}

/** Reparte las fotos entre bloques según la página y las palabras del texto de esa página. */
function proponerFotos(bloques, paginas, fotos) {
  const textoPag = {}; paginas.forEach(p => { textoPag[p.num] = p.lineas.join(" "); });
  const orden = [...fotos].sort((a, b) => (b.w * b.h) - (a.w * a.h)).slice(0, 24);
  bloques.forEach(b => { b.fotos = []; });
  orden.forEach(f => {
    let mejor = -1, puntos = 0;
    bloques.forEach((b, i) => {
      const pt = (bloques[i].pags || []).includes(f.pag) ? 3 : 0;
      const m = ((textoPag[f.pag] || "").match(new RegExp(b.re.source, "gi")) || []).length;
      const tot = pt + Math.min(m, 6);
      if (tot > puntos) { puntos = tot; mejor = i; }
    });
    if (mejor >= 0 && bloques[mejor].fotos.length < (mejor === 3 ? 1 : 2)) bloques[mejor].fotos.push(f.id);
  });
}

async function extraerInforme(archivo, onProgreso) {
  const lib = await cargarPdfjs();
  const pdf = await lib.getDocument({ data: new Uint8Array(await archivo.arrayBuffer()) }).promise;
  const paginas = [], fotos = [], vistas = {};
  for (let n = 1; n <= pdf.numPages; n++) {
    onProgreso && onProgreso(n, pdf.numPages);
    const pag = await pdf.getPage(n);
    paginas.push({ num: n, lineas: await textoDePagina(pag) });
    (await imagenesDePagina(lib, pag, n)).forEach(im => {
      const k = im.w + "x" + im.h;
      (vistas[k] = vistas[k] || []).push(im);
      fotos.push(im);
    });
    pag.cleanup();
  }
  // Logos o fondos repetidos: misma medida en 3 o más páginas
  const repetidas = new Set(Object.keys(vistas).filter(k => vistas[k].length >= 3));
  const utiles = fotos.filter(f => !repetidas.has(f.w + "x" + f.h)).slice(0, 80);
  utiles.forEach((f, i) => { f.id = i; });
  return { paginas, fotos: utiles, descartadas: fotos.length - utiles.length, numPaginas: pdf.numPages };
}

/* ------------------------------ INTERFAZ ------------------------------ */
function campoPrev(etiqueta, valor, alCambiar, opc) {
  opc = opc || {};
  const e = el(opc.linea ? "input" : "textarea", opc.linea ? { type: "text", value: valor } : { rows: String(opc.filas || 2), value: valor });
  e.addEventListener("input", () => { alCambiar(e.value); invalidarSalida(); if (!opc.linea) autoAlto(e); });
  if (!opc.linea) setTimeout(() => autoAlto(e), 0);
  return el("div", { class: "campo" }, el("div", { class: "et" }, el("span", { text: etiqueta })), e);
}

function pintarPreventiva() {
  const cont = $("prevCuerpo");
  cont.innerHTML = "";
  if (!PREV.bloques.length) return;
  cont.appendChild(el("label", { style: "display:flex;gap:8px;align-items:center;margin:8px 0;font-weight:600" },
    el("input", { type: "checkbox", checked: PREV.incluir, onchange: ev => { PREV.incluir = ev.target.checked; invalidarSalida(); } }),
    "Incluir estas láminas en el PPT del Assurance"));
  cont.appendChild(el("div", { class: "grid3" },
    el("div", {}, campoPrev("Lema (arriba)", PREV.lema, v => { PREV.lema = v; }, { linea: true })),
    el("div", {}, campoPrev("Título", PREV.titulo, v => { PREV.titulo = v; }, { linea: true })),
    el("div", {}, campoPrev("Subtítulo", PREV.subtitulo, v => { PREV.subtitulo = v; }, { linea: true }))));

  const nombres = PREV.bloques.map((b, i) => `${i + 1}. ${b.titulo || "(sin título)"}`);
  PREV.bloques.forEach((b, i) => {
    const fotosB = PREV.fotos.filter(f => b.fotos.includes(f.id));
    const caja = el("div", { class: "foco", style: "border-left-color:var(--teal);margin-bottom:18px;padding-top:6px" });
    caja.appendChild(el("div", { class: "grid2" },
      campoPrev(`Bloque ${i + 1} · título`, b.titulo, v => { b.titulo = v; }, { linea: true }),
      el("div", { style: "display:flex;align-items:flex-end;gap:6px;padding-bottom:6px;flex-wrap:wrap" },
        el("button", { class: "sec", type: "button", text: "↑", disabled: i === 0, onclick: () => { [PREV.bloques[i - 1], PREV.bloques[i]] = [PREV.bloques[i], PREV.bloques[i - 1]]; invalidarSalida(); pintarPreventiva(); } }),
        el("button", { class: "sec", type: "button", text: "↓", disabled: i === PREV.bloques.length - 1, onclick: () => { [PREV.bloques[i + 1], PREV.bloques[i]] = [PREV.bloques[i], PREV.bloques[i + 1]]; invalidarSalida(); pintarPreventiva(); } }),
        el("button", { class: "sec", type: "button", text: "Quitar bloque", onclick: () => { PREV.bloques.splice(i, 1); invalidarSalida(); pintarPreventiva(); } }))));
    caja.appendChild(campoPrev("Texto introductorio (opcional)", b.intro, v => { b.intro = v; }));
    caja.appendChild(campoPrev("Puntos (uno por línea; «Palabra clave: detalle» deja la palabra en negrita)", b.items, v => { b.items = v; }, { filas: 4 }));
    caja.appendChild(el("div", { class: "ayuda", text: fotosB.length ? `Fotos de este bloque (${fotosB.length}, máx. 4):` : "Sin fotos asignadas: elige abajo en la galería." }));
    if (fotosB.length) caja.appendChild(el("div", { style: "display:flex;gap:8px;flex-wrap:wrap" }, fotosB.map(f => el("img", { src: f.src, style: "height:70px;border:1px solid var(--line);border-radius:4px" }))));
    cont.appendChild(caja);
  });
  cont.appendChild(el("button", { class: "sec", type: "button", text: "+ Agregar bloque", onclick: () => { PREV.bloques.push({ titulo: "Nuevo bloque", intro: "", items: "", pags: [], fotos: [], re: /$^/ }); invalidarSalida(); pintarPreventiva(); } }));

  cont.appendChild(el("h3", { text: `Galería de fotos del informe (${PREV.fotos.length}) · asigna cada una a un bloque` }));
  const gal = el("div", { style: "display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px" });
  PREV.fotos.forEach(f => {
    const actual = PREV.bloques.findIndex(b => b.fotos.includes(f.id));
    const sel = el("select", { style: "width:100%;padding:4px;border:1px solid var(--line);border-radius:6px;margin-top:4px" },
      el("option", { value: "-1", text: "— no usar —" }),
      nombres.map((n, i) => el("option", { value: String(i), text: n })));
    sel.value = String(actual);
    sel.addEventListener("change", () => {
      const nuevo = Number(sel.value);
      PREV.bloques.forEach(b => { b.fotos = b.fotos.filter(x => x !== f.id); });
      if (nuevo >= 0) {
        if (PREV.bloques[nuevo].fotos.length >= 4) { sel.value = "-1"; reg("Un bloque admite máximo 4 fotos.", "warn"); return; }
        PREV.bloques[nuevo].fotos.push(f.id);
      }
      invalidarSalida(); pintarPreventiva();
    });
    gal.appendChild(el("div", {}, el("img", { src: f.src, style: "width:100%;aspect-ratio:4/3;object-fit:cover;border:1px solid var(--line);border-radius:4px" }),
      el("div", { class: "ayuda", text: (f.pag === "propia" ? "propia" : `pág. ${f.pag}`) + (f.leyenda ? " · " + f.leyenda : "") }), sel));
  });
  cont.appendChild(gal);
}

/** Objeto que recibe el motor (lamina_preventiva.agregar_preventiva); null si no se incluye. */
function construirPreventiva() {
  if (!PREV.incluir || !PREV.bloques.length) return null;
  const porId = {}; PREV.fotos.forEach(f => { porId[f.id] = f.src; });
  return {
    lema: PREV.lema, titulo: PREV.titulo, subtitulo: PREV.subtitulo,
    bloques: PREV.bloques.map(b => ({
      titulo: b.titulo, intro: b.intro,
      items: String(b.items || "").split(/\r?\n/).map(x => x.trim()).filter(Boolean),
      fotos: b.fotos.map(id => porId[id]).filter(Boolean)
    })).filter(b => b.titulo || b.intro || b.items.length || b.fotos.length)
  };
}

$("bPrevExtraer").onclick = async () => {
  const f = $("filePrev").files[0];
  const est = $("estPrev");
  if (f && /\.pptx$/i.test(f.name)) { await extraerPptxJej(f, est); return; }
  if (!f) { est.innerHTML = ""; est.appendChild(el("span", { class: "bad", text: "✖ Primero elige el informe (PPTX de JEJ o PDF)." })); return; }
  est.textContent = "Leyendo el informe…";
  $("bPrevExtraer").disabled = true;
  try {
    const r = await extraerInforme(f, (n, t) => { est.textContent = `Leyendo el informe… página ${n} de ${t}`; });
    const bloques = resumirPorReglas(r.paginas);
    proponerFotos(bloques, r.paginas, r.fotos);
    PREV.bloques = bloques; PREV.fotos = r.fotos; PREV.incluir = true;
    const conTexto = bloques.filter(b => b.items).length;
    est.innerHTML = "";
    est.appendChild(el("span", { class: "ok", text: `✔ ${r.numPaginas} páginas leídas · ${conTexto} bloques con texto · ${r.fotos.length} fotos (${r.descartadas} descartadas por repetidas o chicas). Revisa y corrige abajo.` }));
    if (!r.paginas.some(p => p.lineas.length)) est.appendChild(el("div", { class: "bad", text: "⚠ El PDF no trae texto seleccionable (parece escaneado): escribe los puntos a mano." }));
    reg(`Informe leído: ${r.numPaginas} páginas, ${r.fotos.length} fotos.`, "ok");
    invalidarSalida();
    pintarPreventiva();
  } catch (e) {
    console.error(e);
    est.innerHTML = ""; est.appendChild(el("span", { class: "bad", text: "✖ No se pudo leer el PDF: " + (e && e.message ? e.message : e) }));
    reg("✖ Informe PDF: " + (e && e.message ? e.message : e), "bad");
  } finally { $("bPrevExtraer").disabled = false; }
};

$("bPrevVacio").onclick = () => {
  PREV.bloques = PREV_REGLAS.map(r => ({ titulo: r.titulo, intro: "", items: "", pags: [], fotos: [], re: r.re }));
  PREV.fotos = []; PREV.incluir = true; invalidarSalida(); pintarPreventiva();
};

/** Informe «Cumplimiento de Focos» de JEJ (PPTX): se lee con el motor Python del navegador. */
async function extraerPptxJej(archivo, est) {
  est.textContent = "Leyendo el informe de JEJ…";
  $("bPrevExtraer").disabled = true;
  try {
    const u8 = new Uint8Array(await archivo.arrayBuffer());
    const r = await llamar("puente_web.leer_informe_pptx(_B.to_py())", { _B: u8 });
    if (!r.ok) throw new Error(r.error || "No se pudo leer el informe.");
    const fotos = r.fotos.map(f => ({ id: f.id, src: f.src, w: f.w, h: f.h, pag: f.pag, leyenda: f.leyenda || "" }));
    PREV.fotos = fotos;
    PREV.bloques = PREV_REGLAS.map(x => ({ titulo: x.titulo, intro: "", items: "", pags: [], fotos: [], re: x.re }));
    const ej = PREV.bloques.find(b => b.titulo === "Ejecución");
    ej.intro = r.intro; ej.items = r.items.join("\n"); ej.fotos = (r.sugeridas || []).slice(0, 4);
    if (r.lema) PREV.lema = r.lema;
    PREV.incluir = true;
    RESEJ.datos = (r.resumen && r.resumen.filas && r.resumen.filas.length) ? r.resumen : null;
    RESEJ.incluir = !!RESEJ.datos;
    pintarResEj();
    est.innerHTML = "";
    est.appendChild(el("span", { class: "ok", text: `✔ Informe JEJ · turno ${r.turno || "?"}: ${r.focos.length} focos, ${fotos.length} fotos. El bloque Ejecución quedó armado; completa los otros tres (Coordinación, Procesos y rol, Visita gerencial) y agrega tus fotos si las tienes.` }));
    reg(`Informe JEJ leído: ${r.focos.length} focos, ${fotos.length} fotos.`, "ok");
    invalidarSalida(); pintarPreventiva();
  } catch (e) {
    console.error(e);
    est.innerHTML = ""; est.appendChild(el("span", { class: "bad", text: "✖ " + (e && e.message ? e.message : e) }));
    reg("✖ Informe JEJ: " + (e && e.message ? e.message : e), "bad");
  } finally { $("bPrevExtraer").disabled = false; }
}

/** Fotos propias (reunión de coordinación, visita gerencial, etc.). */
$("filePrevFotos").onchange = async ev => {
  for (const f of ev.target.files) {
    try {
      const bm = await createImageBitmap(f);
      const esc = Math.min(1, 1000 / bm.width);
      const cv = document.createElement("canvas"); cv.width = Math.round(bm.width * esc); cv.height = Math.round(bm.height * esc);
      cv.getContext("2d").drawImage(bm, 0, 0, cv.width, cv.height);
      const id = PREV.fotos.reduce((m, x) => Math.max(m, x.id), -1) + 1;
      PREV.fotos.push({ id, src: cv.toDataURL("image/jpeg", 0.85), w: cv.width, h: cv.height, pag: "propia", leyenda: f.name });
    } catch (e) { reg("✖ No se pudo leer la foto " + f.name, "bad"); }
  }
  ev.target.value = "";
  if (!PREV.bloques.length) $("bPrevVacio").click(); else { invalidarSalida(); pintarPreventiva(); }
};


/* ============ RESUMEN EJECUTIVO ASSURANCE (indicadores + tabla por foco) ============ */
const RESEJ = { incluir: false, datos: null };
const COLORES_KPI = [["verde", "Verde"], ["azul", "Azul"], ["ambar", "Ámbar"], ["negro", "Negro"]];
const COLORES_EST = [["verde", "Verde (conforme)"], ["ambar", "Ámbar (hallazgo / mejora)"], ["rojo", "Rojo"], ["gris", "Gris"]];

function selColor(opciones, valor, alCambiar) {
  const s = el("select", { style: "padding:5px;border:1px solid var(--line);border-radius:6px" },
    opciones.map(([v, t]) => el("option", { value: v, text: t })));
  s.value = valor; s.addEventListener("change", () => { alCambiar(s.value); invalidarSalida(); });
  return s;
}

function pintarResEj() {
  const cont = $("resEjCuerpo"); cont.innerHTML = "";
  const d = RESEJ.datos;
  if (!d || !d.filas) { cont.appendChild(el("div", { class: "mut", text: "Este informe no trae la tabla «Resumen de Verificaciones»: no se pudo armar la lámina." })); return; }
  cont.appendChild(el("label", { style: "display:flex;gap:8px;align-items:center;margin:8px 0;font-weight:600" },
    el("input", { type: "checkbox", checked: RESEJ.incluir, onchange: ev => { RESEJ.incluir = ev.target.checked; invalidarSalida(); } }),
    "Incluir esta lámina en el PPT del Assurance"));
  cont.appendChild(el("div", { class: "grid2" },
    campoPrev("Título", d.titulo, v => { d.titulo = v; }, { linea: true }),
    campoPrev("Subtítulo", d.subtitulo, v => { d.subtitulo = v; }, { linea: true })));
  cont.appendChild(el("h3", { text: "Indicadores (tarjetas de arriba)" }));
  d.kpis.forEach((k, i) => cont.appendChild(el("div", { class: "grid3", style: "align-items:end" },
    campoPrev(`Indicador ${i + 1} · título`, k.titulo, v => { k.titulo = v; }, { linea: true }),
    campoPrev("Valor", k.valor, v => { k.valor = v; }, { linea: true }),
    el("div", { class: "campo" }, el("div", { class: "et" }, el("span", { text: "Color" })), selColor(COLORES_KPI, k.color, v => { k.color = v; })))));
  cont.appendChild(el("h3", { text: `Tabla por foco (${d.filas.length} filas)` }));
  d.filas.forEach((f, i) => {
    const caja = el("div", { class: "foco", style: "border-left-color:var(--teal);margin-bottom:14px;padding-top:6px" });
    caja.appendChild(el("div", { class: "grid2" },
      campoPrev(`Fila ${i + 1} · foco`, f.foco, v => { f.foco = v; }, { linea: true }),
      campoPrev("Ámbito", f.ambito, v => { f.ambito = v; }, { linea: true })));
    caja.appendChild(el("div", { class: "grid2" },
      campoPrev("Frente / empresa", f.frente, v => { f.frente = v; }, { filas: 2 }),
      campoPrev("ITO", f.ito, v => { f.ito = v; }, { filas: 2 })));
    caja.appendChild(el("div", { class: "grid2" },
      campoPrev("Herramienta SGE (una por línea)", f.herramienta, v => { f.herramienta = v; }, { filas: 2 }),
      campoPrev("Estado (admite 2 líneas)", f.estado, v => { f.estado = v; }, { filas: 2 })));
    caja.appendChild(campoPrev("Evidencia de cumplimiento / acción en terreno", f.evidencia, v => { f.evidencia = v; }, { filas: 3 }));
    caja.appendChild(el("div", { style: "display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:6px" },
      el("span", { class: "mut", text: "Color del estado:" }), selColor(COLORES_EST, f.estado_color, v => { f.estado_color = v; }),
      el("button", { class: "sec", type: "button", text: "↑", disabled: i === 0, onclick: () => { [d.filas[i - 1], d.filas[i]] = [d.filas[i], d.filas[i - 1]]; invalidarSalida(); pintarResEj(); } }),
      el("button", { class: "sec", type: "button", text: "↓", disabled: i === d.filas.length - 1, onclick: () => { [d.filas[i + 1], d.filas[i]] = [d.filas[i], d.filas[i + 1]]; invalidarSalida(); pintarResEj(); } }),
      el("button", { class: "sec", type: "button", text: "Quitar fila", onclick: () => { d.filas.splice(i, 1); invalidarSalida(); pintarResEj(); } })));
    cont.appendChild(caja);
  });
  cont.appendChild(el("div", { style: "display:flex;gap:8px;flex-wrap:wrap;margin-top:8px" },
    el("button", { class: "sec", type: "button", text: "+ Agregar fila", onclick: () => { d.filas.push({ foco: "Foco", ambito: "", frente: "", ito: "", herramienta: "", evidencia: "", estado: "Conforme", estado_color: "verde" }); invalidarSalida(); pintarResEj(); } }),
    el("button", { id: "bResSuelto", type: "button", text: "Descargar solo esta lámina (PPTX)", onclick: descargarResumenSuelto })));
  cont.appendChild(el("div", { class: "estado", id: "estResSuelto" }));
}

function construirResumenEj() {
  const d = RESEJ.datos;
  if (!RESEJ.incluir || !d || !d.filas || !d.filas.length) return null;
  return JSON.parse(JSON.stringify(d));
}

async function descargarResumenSuelto() {
  const est = $("estResSuelto"); est.textContent = "Generando…";
  const d = RESEJ.datos;
  if (!d || !d.filas || !d.filas.length) { est.textContent = "✖ No hay filas."; return; }
  const liberar = ocupar("gen");
  try {
    if (!AS.plantilla) AS.plantilla = await bajar(PLANTILLA, "bin");
    const r = await llamar("puente_web.resumen_suelto(_D, bytes(_PL.to_py()))", { _D: JSON.stringify(d), _PL: AS.plantilla });
    regPython(r.log);
    if (!r.ok) throw new Error((r.errores || ["No se pudo generar la lámina."]).join(" "));
    const p = await PY.runPythonAsync("import puente_web\npuente_web.obtener_suelto()");
    const u8 = (p && typeof p.toJs === "function") ? p.toJs().slice() : new Uint8Array(p);
    if (p && typeof p.destroy === "function") p.destroy();
    const nombre = conHora("Resumen_ejecutivo_Assurance", ".pptx", "RESUMEN_EJECUTIVO");
    guardarBytes(u8, nombre, TIPO_PPTX);
    est.innerHTML = ""; est.appendChild(el("span", { class: "ok", text: `✔ Descargada: ${nombre}` }));
    reg(`Descarga: ${nombre}`);
  } catch (e) {
    console.error(e);
    est.innerHTML = ""; est.appendChild(el("span", { class: "bad", text: "✖ " + (e && e.message ? e.message : e) }));
  } finally { liberar(); }
}
