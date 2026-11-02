/* Utopía Coffee Club · Excel para el dueño
   Genera un .xlsx con la marca de la app (granate y crema), tablas con filtros,
   totales y gráficos. No depende de librerías de Excel: escribe el formato a mano y lo comprime con JSZip. */
(function(){
  const C = {wine:"5E0F1C", cream:"F1E9DB", surface:"FAF5EC", soft:"ECDCD6", text:"3B0A12", muted:"8A4A52", line:"DCC6BD", teal:"066D77", gold:"A8680F", red:"C2360F", white:"FFFFFF"};
  const SERIES = [C.wine, C.gold, C.teal, C.muted, C.red, "B98F90", "4A2C2F", "D9A441"];
  const xe = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  const colL = n => { let s = ""; n++; while (n){ const m = (n-1) % 26; s = String.fromCharCode(65+m) + s; n = Math.floor((n-1)/26); } return s; };
  const ref = (c, r) => colL(c) + (r+1);
  const aref = (sheet, c1, r1, c2, r2) => `'${sheet.replace(/'/g, "''")}'!$${colL(c1)}$${r1+1}:$${colL(c2)}$${r2+1}`;

  /* ---------- estilos ---------- */
  const FMT = {eur:164, pct:9, int:3, num1:165, eur0:166, eur4:167};
  const NUMFMTS = `<numFmts count="4"><numFmt numFmtId="167" formatCode="#,##0.0000\\ &quot;€&quot;"/><numFmt numFmtId="164" formatCode="#,##0.00\\ &quot;€&quot;"/><numFmt numFmtId="165" formatCode="#,##0.0"/><numFmt numFmtId="166" formatCode="#,##0\\ &quot;€&quot;"/></numFmts>`;
  const FONTS = [
    `<font><sz val="11"/><color rgb="FF${C.text}"/><name val="Calibri"/><family val="2"/></font>`,              // 0 normal
    `<font><b/><sz val="26"/><color rgb="FF${C.cream}"/><name val="Bebas Neue"/></font>`,                         // 1 título
    `<font><sz val="10"/><color rgb="FF${C.cream}"/><name val="Calibri"/><family val="2"/></font>`,             // 2 subtítulo
    `<font><b/><sz val="11"/><color rgb="FF${C.cream}"/><name val="Calibri"/><family val="2"/></font>`,         // 3 cabecera tabla
    `<font><b/><sz val="9"/><color rgb="FF${C.muted}"/><name val="Calibri"/><family val="2"/></font>`,          // 4 etiqueta KPI
    `<font><b/><sz val="20"/><color rgb="FF${C.wine}"/><name val="Calibri"/><family val="2"/></font>`,          // 5 valor KPI
    `<font><b/><sz val="11"/><color rgb="FF${C.wine}"/><name val="Calibri"/><family val="2"/></font>`,          // 6 totales
    `<font><b/><sz val="14"/><color rgb="FF${C.wine}"/><name val="Bebas Neue"/></font>`,                         // 7 título de sección
    `<font><i/><sz val="9"/><color rgb="FF${C.muted}"/><name val="Calibri"/><family val="2"/></font>`,          // 8 nota
    `<font><b/><sz val="11"/><color rgb="FF${C.red}"/><name val="Calibri"/><family val="2"/></font>`,           // 9 negativo
    `<font><b/><sz val="11"/><color rgb="FF${C.teal}"/><name val="Calibri"/><family val="2"/></font>`           // 10 positivo
  ];
  const FILLS = [`<fill><patternFill patternType="none"/></fill>`, `<fill><patternFill patternType="gray125"/></fill>`,
    ...[C.wine, C.cream, C.surface, C.soft, C.white].map(c => `<fill><patternFill patternType="solid"><fgColor rgb="FF${c}"/><bgColor indexed="64"/></patternFill></fill>`)];
  const FILL = {none:0, wine:2, cream:3, surface:4, soft:5, white:6};
  const BORDERS = [`<border><left/><right/><top/><bottom/><diagonal/></border>`,
    `<border><left/><right/><top/><bottom style="thin"><color rgb="FF${C.line}"/></bottom><diagonal/></border>`,
    `<border><left/><right/><top style="medium"><color rgb="FF${C.wine}"/></top><bottom style="medium"><color rgb="FF${C.wine}"/></bottom><diagonal/></border>`,
    `<border><left style="thick"><color rgb="FF${C.wine}"/></left><right/><top/><bottom/><diagonal/></border>`];
  const xfs = [], xfKey = new Map();
  function sx({font = 0, fill = 0, border = 0, fmt = 0, h = null, v = "center", wrap = false, indent = 0} = {}){
    const k = [font, fill, border, fmt, h, v, wrap, indent].join("|");
    if (xfKey.has(k)) return xfKey.get(k);
    const al = `<alignment${h ? ` horizontal="${h}"` : ""} vertical="${v}"${wrap ? ` wrapText="1"` : ""}${indent ? ` indent="${indent}"` : ""}/>`;
    xfs.push(`<xf numFmtId="${fmt}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${al}</xf>`);
    xfKey.set(k, xfs.length - 1); return xfs.length - 1;
  }
  sx(); // 0 = normal
  const stylesXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${NUMFMTS}
<fonts count="${FONTS.length}">${FONTS.join("")}</fonts><fills count="${FILLS.length}">${FILLS.join("")}</fills><borders count="${BORDERS.length}">${BORDERS.join("")}</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleLight1" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;

  /* ---------- textos compartidos ---------- */
  let sst = [], sstMap = new Map(), sstCount = 0;
  function ssIdx(t){ sstCount++; if (!sstMap.has(t)){ sstMap.set(t, sst.length); sst.push(t); } return sstMap.get(t); }
  const sstXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sstCount}" uniqueCount="${sst.length}">${sst.map(t => `<si><t xml:space="preserve">${xe(t)}</t></si>`).join("")}</sst>`;
  const THEME = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\"><a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2><a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2><a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1><a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2><a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3><a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4><a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5><a:accent6><a:srgbClr val=\"F79646\"/></a:accent6><a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Cambria\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"\uff2d\uff33 \uff30\u30b4\u30b7\u30c3\u30af\"/><a:font script=\"Hang\" typeface=\"\ub9d1\uc740 \uace0\ub515\"/><a:font script=\"Hans\" typeface=\"\u5b8b\u4f53\"/><a:font script=\"Hant\" typeface=\"\u65b0\u7d30\u660e\u9ad4\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"\uff2d\uff33 \uff30\u30b4\u30b7\u30c3\u30af\"/><a:font script=\"Hang\" typeface=\"\ub9d1\uc740 \uace0\ub515\"/><a:font script=\"Hans\" typeface=\"\u5b8b\u4f53\"/><a:font script=\"Hant\" typeface=\"\u65b0\u7d30\u660e\u9ad4\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"35000\"><a:schemeClr val=\"phClr\"><a:tint val=\"37000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:tint val=\"15000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"1\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:shade val=\"51000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"80000\"><a:schemeClr val=\"phClr\"><a:shade val=\"93000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"94000\"/><a:satMod val=\"135000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"95000\"/><a:satMod val=\"105000\"/></a:schemeClr></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln><a:ln w=\"38100\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"20000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"38000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst=\"orthographicFront\"><a:rot lat=\"0\" lon=\"0\" rev=\"0\"/></a:camera><a:lightRig rig=\"threePt\" dir=\"t\"><a:rot lat=\"0\" lon=\"0\" rev=\"1200000\"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w=\"63500\" h=\"25400\"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"40000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"40000\"><a:schemeClr val=\"phClr\"><a:tint val=\"45000\"/><a:shade val=\"99000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"20000\"/><a:satMod val=\"255000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"-80000\" r=\"50000\" b=\"180000\"/></a:path></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"80000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"30000\"/><a:satMod val=\"200000\"/></a:schemeClr></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>";
  /* ---------- hoja ---------- */
  function Sheet(name, widths){
    this.name = name.slice(0, 31); this.cells = new Map(); this.heights = {}; this.merges = []; this.widths = widths || []; this.tables = []; this.charts = []; this.maxC = Math.max(0, (widths||[]).length - 1); this.maxR = 0;
  }
  Sheet.prototype.set = function(c, r, v, s, f){ this.cells.set(r*1000 + c, {c, r, v, s: s || 0, f}); this.maxC = Math.max(this.maxC, c); this.maxR = Math.max(this.maxR, r); return this; };
  Sheet.prototype.xml = function(){
    const rows = new Map();
    [...this.cells.values()].forEach(x => { if (!rows.has(x.r)) rows.set(x.r, []); rows.get(x.r).push(x); });
    const body = [...rows.keys()].sort((a,b) => a-b).map(r => {
      const cs = rows.get(r).sort((a,b) => a.c - b.c).map(x => {
        const a = ref(x.c, x.r);
        if (x.f) return `<c r="${a}" s="${x.s}"><f>${xe(x.f)}</f>${typeof x.v === "number" && isFinite(x.v) ? `<v>${x.v}</v>` : ""}</c>`;
        if (x.v === null || x.v === undefined || x.v === "") return `<c r="${a}" s="${x.s}"/>`;
        if (typeof x.v === "number") return isFinite(x.v) ? `<c r="${a}" s="${x.s}"><v>${Math.round(x.v*1e6)/1e6}</v></c>` : `<c r="${a}" s="${x.s}"/>`;
        return `<c r="${a}" s="${x.s}" t="s"><v>${ssIdx(String(x.v))}</v></c>`;
      }).join("");
      const h = this.heights[r];
      return `<row r="${r+1}"${h ? ` ht="${h}" customHeight="1"` : ""}>${cs}</row>`;
    }).join("");
    const cols = this.widths.length ? `<cols>${this.widths.map((w,i) => `<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join("")}</cols>` : "";
    const merges = this.merges.length ? `<mergeCells count="${this.merges.length}">${this.merges.map(m => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` : "";
    const drawing = this.charts.length ? `<drawing r:id="rId${this.tables.length + 1}"/>` : "";
    const tparts = this.tables.length ? `<tableParts count="${this.tables.length}">${this.tables.map((t,i) => `<tablePart r:id="rId${i+1}"/>`).join("")}</tableParts>` : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><tabColor rgb="FF${C.wine}"/><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:${ref(this.maxC, this.maxR)}"/>
<sheetViews><sheetView workbookViewId="0" showGridLines="0" zoomScale="100"/></sheetViews><sheetFormatPr defaultRowHeight="18" customHeight="1"/>
${cols}<sheetData>${body}</sheetData>${merges}
<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>
<headerFooter><oddFooter>&amp;L&amp;8Utopía Coffee Club&amp;R&amp;8Página &amp;P de &amp;N</oddFooter></headerFooter>${drawing}${tparts}</worksheet>`;
  };

  /* Cabecera de marca: banda granate con el título y una línea de subtítulo */
  function banner(sh, title, subtitle, span){
    const last = Math.max(span - 1, 5);
    for (let c = 0; c <= last; c++){ sh.set(c, 0, c ? "" : title, sx({font:1, fill:FILL.wine, h:"left", indent: c ? 0 : 1})); sh.set(c, 1, c ? "" : subtitle, sx({font:2, fill:FILL.wine, h:"left", indent: c ? 0 : 1, v:"top"})); }
    sh.merges.push(`A1:${ref(last, 0)}`, `A2:${ref(last, 1)}`);
    sh.heights[0] = 40; sh.heights[1] = 22; sh.heights[2] = 10;
    return 3;
  }
  /* Tarjetas de cifras clave: etiqueta pequeña y valor grande, con barra granate a la izquierda */
  function kpis(sh, row, list){
    list.forEach((k, i) => {
      const c = i*2;
      sh.set(c, row, String(k.label).toUpperCase(), sx({font:4, fill:FILL.surface, border:3, h:"left", indent:1, v:"bottom"}));
      sh.set(c+1, row, "", sx({fill:FILL.surface}));
      sh.set(c, row+1, k.value, sx({font: k.tone === "bad" ? 9 : 5, fill:FILL.surface, border:3, fmt: FMT[k.fmt] || 0, h:"left", indent:1}));
      sh.set(c+1, row+1, "", sx({fill:FILL.surface}));
      sh.merges.push(`${ref(c,row)}:${ref(c+1,row)}`, `${ref(c,row+1)}:${ref(c+1,row+1)}`);
    });
    sh.heights[row] = 20; sh.heights[row+1] = 32; sh.heights[row+2] = 12;
    return row + 3;
  }
  function section(sh, row, text, span){ sh.set(0, row, text, sx({font:7, v:"bottom"})); sh.heights[row] = 26; return row + 1; }
  function note(sh, row, text, span){
    sh.set(0, row, text, sx({font:8, wrap:true, v:"top"})); if (span > 1) sh.merges.push(`A${row+1}:${ref(span-1, row)}`);
    sh.heights[row] = Math.max(18, Math.ceil(String(text).length / (span*11)) * 14); return row + 1;
  }
  let tableSeq = 0, NO_TABLES = false;
  /* Tabla con cabecera granate, filas alternas, filtros y fila de totales opcional.
     cols: [{h:"Ventas", fmt:"eur", total:"sum"|"avg"|"label"|fn, tone:true}] */
  function table(sh, row, col, name, cols, rows, opts = {}){
    const r0 = row, c0 = col, c1 = col + cols.length - 1;
    cols.forEach((k, j) => sh.set(c0+j, r0, k.h, sx({font:3, fill:FILL.wine, h: k.fmt ? "right" : "left", wrap:true, indent: k.fmt ? 0 : 1})));
    sh.heights[r0] = opts.headH || 30;
    rows.forEach((rw, i) => {
      const fill = i % 2 ? FILL.surface : FILL.white;
      cols.forEach((k, j) => {
        const v = rw[j];
        const neg = k.tone && typeof v === "number" && v < 0;
        sh.set(c0+j, r0+1+i, v, sx({font: neg ? 9 : 0, fill, border:1, fmt: FMT[k.fmt] || 0, h: k.fmt ? "right" : "left", indent: k.fmt ? 0 : 1, wrap: !k.fmt, v: "top"}));
      });
      const lines = Math.max(1, ...cols.map((k, j) => k.fmt ? 1 : Math.ceil(String(rw[j] ?? "").length / Math.max(4, ((sh.widths[c0+j] || 12) - 2) * 1.05))));
      if (lines > 1) sh.heights[r0+1+i] = 15 * lines + 4;
    });
    const rEnd = r0 + Math.max(rows.length, 1);
    if (!rows.length) cols.forEach((k, j) => sh.set(c0+j, r0+1, j ? "" : "Sin datos todavía", sx({font:8, fill:FILL.white, border:1, indent:1})));
    let totalRow = null;
    if (opts.totals && rows.length){
      totalRow = rEnd + 1;
      cols.forEach((k, j) => {
        const st = sx({font: 6, fill:FILL.soft, border:2, fmt: FMT[k.fmt] || 0, h: k.fmt ? "right" : "left", indent: j ? 0 : 1});
        const rng = `${ref(c0+j, r0+1)}:${ref(c0+j, rEnd)}`;
        const nums = rows.map(r => r[j]).filter(v => typeof v === "number" && isFinite(v));
        if (j === 0) sh.set(c0, totalRow, opts.totalLabel || "Total", st);
        else if (k.total === "sum") sh.set(c0+j, totalRow, nums.reduce((a,b) => a+b, 0), st, `SUBTOTAL(109,${rng})`);
        else if (k.total === "avg") sh.set(c0+j, totalRow, nums.length ? nums.reduce((a,b) => a+b, 0)/nums.length : 0, st, `IFERROR(SUBTOTAL(101,${rng}),0)`);
        else if (typeof k.total === "function"){ const t = k.total(totalRow, (jj) => ref(c0+jj, totalRow)); sh.set(c0+j, totalRow, t.v, st, t.f); }
        else sh.set(c0+j, totalRow, "", st);
      });
      sh.heights[totalRow] = 22;
    }
    if (rows.length && !NO_TABLES){
      tableSeq++;
      sh.tables.push({id: tableSeq, name: (name.replace(/[^A-Za-z0-9_]/g, "") || "Tabla") + tableSeq, ref: `${ref(c0, r0)}:${ref(c1, rEnd)}`, cols: cols.map(k => k.h)});
    }
    return {r0, rEnd, c0, c1, next: (totalRow !== null ? totalRow : rEnd) + 2, n: rows.length};
  }
  /* Gráfico anclado en celdas. type: col | bar | line | doughnut. series: [{name, col}] con col = columna de la tabla */
  function chart(sh, t, spec){
    if (!t.n) return;
    sh.charts.push({...spec, sheet: sh.name, cat: {c: t.c0 + spec.cat, r1: t.r0+1, r2: t.rEnd, head: t.r0}, series: spec.series.map(s => ({...s, c: t.c0 + s.col, r1: t.r0+1, r2: t.rEnd, head: t.r0}))});
  }

  /* ---------- gráficos ---------- */
  function chartXml(ch, sheetObj){
    const cell = (c, r) => sheetObj.cells.get(r*1000 + c);
    const catVals = []; for (let r = ch.cat.r1; r <= ch.cat.r2; r++){ const x = cell(ch.cat.c, r); catVals.push(x ? String(x.v ?? "") : ""); }
    const strCache = vals => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v,i) => `<c:pt idx="${i}"><c:v>${xe(v)}</c:v></c:pt>`).join("")}</c:strCache>`;
    const numCache = (vals, code) => `<c:numCache><c:formatCode>${xe(code)}</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v,i) => typeof v === "number" && isFinite(v) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : "").join("")}</c:numCache>`;
    const txPr = (sz, color, bold) => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}" b="${bold ? 1 : 0}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="es-ES"/></a:p></c:txPr>`;
    const isPie = ch.type === "doughnut";
    const ser = ch.series.map((s, i) => {
      const vals = []; for (let r = s.r1; r <= s.r2; r++){ const x = cell(s.c, r); vals.push(x ? x.v : null); }
      const name = cell(s.c, s.head)?.v || s.name || "";
      const color = s.color || SERIES[i % SERIES.length];
      const sp = ch.type === "line"
        ? `<c:spPr><a:ln w="31750" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:marker>`
        : isPie ? `<c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>${vals.map((v,k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${SERIES[k % SERIES.length]}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join("")}`
        : `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>`;
      const lbl = ch.labels ? `<c:dLbls><c:numFmt formatCode="${xe(isPie ? "0%" : (ch.labelFmt || "#,##0"))}" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txPr(800, isPie ? "FFFFFF" : C.text, true)}${isPie ? "" : `<c:dLblPos val="${ch.type === "line" ? "t" : "outEnd"}"/>`}<c:showLegendKey val="0"/><c:showVal val="${isPie ? 0 : 1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${isPie ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>` : "";
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>${xe(aref(ch.sheet, s.c, s.head, s.c, s.head))}</c:f>${strCache([String(name)])}</c:strRef></c:tx>${sp}${lbl}
<c:cat><c:strRef><c:f>${xe(aref(ch.sheet, ch.cat.c, ch.cat.r1, ch.cat.c, ch.cat.r2))}</c:f>${strCache(catVals)}</c:strRef></c:cat>
<c:val><c:numRef><c:f>${xe(aref(ch.sheet, s.c, s.r1, s.c, s.r2))}</c:f>${numCache(vals, ch.fmt || "General")}</c:numRef></c:val>${ch.type === "line" ? `<c:smooth val="0"/>` : ""}</c:ser>`;
    }).join("");
    const gridLn = `<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="${C.line}"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:majorGridlines>`;
    const axLn = `<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="${C.line}"/></a:solidFill></a:ln></c:spPr>`;
    const horiz = ch.type === "bar";
    const catAx = `<c:catAx><c:axId val="1001"/><c:scaling><c:orientation val="${horiz ? "maxMin" : "minMax"}"/></c:scaling><c:delete val="0"/><c:axPos val="${horiz ? "l" : "b"}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axLn}${txPr(900, C.text, false)}<c:crossAx val="1002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
    const valAx = `<c:valAx><c:axId val="1002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${ch.hideValAx ? 1 : 0}"/><c:axPos val="${horiz ? "t" : "l"}"/>${gridLn}<c:numFmt formatCode="${xe(ch.fmt || "General")}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${txPr(800, C.muted, false)}<c:crossAx val="1001"/><c:crosses val="${horiz ? "max" : "autoZero"}"/><c:crossBetween val="between"/></c:valAx>`;
    let plot;
    if (ch.type === "line") plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}<c:marker val="1"/><c:axId val="1001"/><c:axId val="1002"/></c:lineChart>${catAx}${valAx}`;
    else if (isPie) plot = `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="58"/></c:doughnutChart>`;
    else plot = `<c:barChart><c:barDir val="${horiz ? "bar" : "col"}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser}<c:gapWidth val="${ch.gap ?? 60}"/><c:overlap val="${ch.series.length > 1 ? -10 : 0}"/><c:axId val="1001"/><c:axId val="1002"/></c:barChart>${catAx}${valAx}`;
    const legend = (ch.series.length > 1 || isPie) ? `<c:legend><c:legendPos val="${isPie ? "r" : "t"}"/><c:overlay val="0"/>${txPr(900, C.text, false)}</c:legend>` : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<c:lang val="es-ES"/><c:roundedCorners val="0"/><c:chart>
<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="${C.wine}"/></a:solidFill><a:latin typeface="Bebas Neue"/></a:defRPr></a:pPr><a:r><a:rPr lang="es-ES" sz="1400" b="1"><a:solidFill><a:srgbClr val="${C.wine}"/></a:solidFill><a:latin typeface="Bebas Neue"/></a:rPr><a:t>${xe(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${plot}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>
<c:spPr><a:solidFill><a:srgbClr val="${C.surface}"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="${C.line}"/></a:solidFill></a:ln></c:spPr>
<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="es-ES"/></a:p></c:txPr></c:chartSpace>`;
  }
  const drawingXml = charts => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
${charts.map((ch, i) => `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${ch.at.c}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${ch.at.r}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${ch.at.c + ch.at.w}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${ch.at.r + ch.at.h}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i+2}" name="Gráfico ${i+1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId${i+1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`).join("")}
</xdr:wsDr>`;
  const tableXml = t => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${t.id}" name="${t.name}" displayName="${t.name}" ref="${t.ref}" totalsRowShown="0"><autoFilter ref="${t.ref}"/>
<tableColumns count="${t.cols.length}">${t.cols.map((h,i) => `<tableColumn id="${i+1}" name="${xe(uniq(t.cols, i))}"/>`).join("")}</tableColumns><tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/></table>`;
  function uniq(cols, i){ const h = String(cols[i]).replace(/\n/g, " "); const before = cols.slice(0, i).filter(x => String(x).replace(/\n/g, " ") === h).length; return before ? `${h} ${before+1}` : h; }

  /* ---------- libro ---------- */
  async function pack(JSZip, sheets){
    const files = []; const z = {file: (n, d) => files.push([n, d])};
    let chartN = 0, drawN = 0;
    const ct = [`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`, `<Default Extension="xml" ContentType="application/xml"/>`,
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`,
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`,
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`,
      `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`];
    sheets.forEach((sh, i) => {
      const n = i + 1, rels = [];
      ct.push(`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
      sh.tables.forEach((t, k) => {
        z.file(`xl/tables/table${t.id}.xml`, tableXml(t));
        ct.push(`<Override PartName="/xl/tables/table${t.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
        rels.push(`<Relationship Id="rId${k+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${t.id}.xml"/>`);
      });
      if (sh.charts.length){
        drawN++;
        const drels = sh.charts.map((ch, k) => {
          chartN++;
          z.file(`xl/charts/chart${chartN}.xml`, chartXml(ch, sh));
          ct.push(`<Override PartName="/xl/charts/chart${chartN}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
          return `<Relationship Id="rId${k+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartN}.xml"/>`;
        });
        z.file(`xl/drawings/drawing${drawN}.xml`, drawingXml(sh.charts));
        z.file(`xl/drawings/_rels/drawing${drawN}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drels.join("")}</Relationships>`);
        ct.push(`<Override PartName="/xl/drawings/drawing${drawN}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
        rels.push(`<Relationship Id="rId${sh.tables.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawN}.xml"/>`);
      }
      z.file(`xl/worksheets/sheet${n}.xml`, sh.xml());
      if (rels.length) z.file(`xl/worksheets/_rels/sheet${n}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`);
    });
    z.file("xl/styles.xml", stylesXml());
    z.file("xl/theme/theme1.xml", THEME);
    z.file("xl/sharedStrings.xml", sstXml());
    ct.push(`<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`, `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`);
    z.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="16000" activeTab="0"/></bookViews>
<sheets>${sheets.map((s,i) => `<sheet name="${xe(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join("")}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
    z.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s,i) => `<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId${sheets.length+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId${sheets.length+3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
    z.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
    const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
    z.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Utopía Coffee Club · Informe</dc:title><dc:creator>Utopía Comandas</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    z.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Utopía Comandas</Application></Properties>`);
    z.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${ct.join("")}</Types>`);
    const zip = new JSZip();
    const order = n => n === "[Content_Types].xml" ? 0 : n === "_rels/.rels" ? 1 : n.startsWith("docProps/") ? 2 : n === "xl/workbook.xml" ? 3 : 4;
    files.sort((x, y) => order(x[0]) - order(y[0])).forEach(([n, d]) => zip.file(n, d, {createFolders:false}));
    return zip.generateAsync({type: typeof Blob !== "undefined" && typeof window !== "undefined" ? "blob" : "uint8array", mimeType:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", compression:"DEFLATE"});
  }

  /* ---------- el informe ----------
     d = {generado, periodo, kpis:{ventas,pedidos,ticket,gastos,beneficio,margen,dias},
          meses:[[mes,dias,pedidos,ventas,ticket,gastos,fijos,beneficio,margen]],
          cierres:[[fecha,abre,cierra,pedidos,ventas,bizum,ticket,propinas,entradas,salidas,debia,contado,diferencia,espera]],
          horas:[[hora,pedidos,ventas]], semana:[[dia,dias,pedidos,ventas,media]],
          productos:[[nombre,grupo,uds,ventas,pct,precio,coste,margen]],
          gastos:[[fecha,concepto,categoria,importe,origen]], categorias:[[cat,importe]], fijos:[[concepto,categoria,importe]] } */
  function build(JSZip, d){
    xfs.length = 0; xfKey.clear(); sx(); tableSeq = 0; sst = []; sstMap = new Map(); sstCount = 0; NO_TABLES = !!d.interno;
    const sheets = [];
    const sub = `${d.periodo} · Generado el ${d.generado}`;

    // 1. Resumen
    {
      const sh = new Sheet("Resumen", [22, 11, 11, 13, 12, 14, 13, 13, 11]);
      let r = banner(sh, "UTOPÍA COFFEE CLUB · INFORME DE NÚMEROS", sub, 9);
      const k = d.kpis;
      r = kpis(sh, r, [{label:"Ventas", value:k.ventas, fmt:"eur"}, {label:"Pedidos", value:k.pedidos, fmt:"int"}, {label:"Ticket medio", value:k.ticket, fmt:"eur"}, {label:"Beneficio", value:k.beneficio, fmt:"eur", tone: k.beneficio < 0 ? "bad" : ""}]);
      r = kpis(sh, r, [{label:"Días abiertos", value:k.dias, fmt:"int"}, {label:"Venta media al día", value:k.dias ? k.ventas/k.dias : 0, fmt:"eur"}, {label:"Costes (con fijos)", value:k.gastos, fmt:"eur"}, {label:"Margen", value:k.margen, fmt:"pct", tone: k.margen < 0 ? "bad" : ""}]);
      r = section(sh, r, "Mes a mes", 9);
      const t = table(sh, r, 0, "Meses", [
        {h:"Mes"}, {h:"Días abiertos", fmt:"int", total:"sum"}, {h:"Pedidos", fmt:"int", total:"sum"}, {h:"Ventas", fmt:"eur", total:"sum"}, {h:"Ticket medio", fmt:"eur", total:(tr, R) => ({f:`IFERROR(${R(3)}/${R(2)},0)`, v: k.pedidos ? k.ventas/k.pedidos : 0})},
        {h:"Coste y gastos", fmt:"eur", total:"sum"}, {h:"Gastos fijos", fmt:"eur", total:"sum"}, {h:"Beneficio", fmt:"eur", total:"sum", tone:true}, {h:"Margen", fmt:"pct", total:(tr, R) => ({f:`IFERROR(${R(7)}/${R(3)},0)`, v: k.margen})}
      ], d.meses, {totals:true});
      r = t.next;
      chart(sh, t, {type:"col", title:"Ventas, gastos y beneficio por mes", cat:0, series:[{col:3, color:C.wine}, {col:5, color:C.gold}, {col:7, color:C.teal}], fmt:'#,##0 "€"', at:{c:0, r, w:9, h:16}});
      r += 17;
      r = note(sh, r, "Gastos fijos: el importe mensual actual de los fijos (alquiler, autónomos, luz…) aplicado a cada mes. Coste y gastos = coste de lo vendido según recetas y escandallos (ingredientes, vaso, pegatina, hielo, agua, también lo invitado) + otros gastos que no son compras de suministros. Beneficio = ventas − coste y gastos − fijos. Las cifras salen de los cierres de caja de la app; el Bizum cuenta como venta.", 9);
      sheets.push(sh);
    }
    // 2. Caja por día
    {
      const sh = new Sheet("Caja por día", [12, 8, 8, 9, 11, 11, 11, 10, 10, 10, 12, 11, 11, 11]);
      let r = banner(sh, "CAJA POR DÍA", sub, 14);
      r = section(sh, r, "Cierres de caja", 14);
      const t = table(sh, r, 0, "Cierres", [
        {h:"Fecha"}, {h:"Abre"}, {h:"Cierra"}, {h:"Pedidos", fmt:"int", total:"sum"}, {h:"Ventas", fmt:"eur", total:"sum"}, {h:"De ellas Bizum", fmt:"eur", total:"sum"}, {h:"Ticket medio", fmt:"eur", total:"avg"},
        {h:"Propinas", fmt:"eur", total:"sum"}, {h:"Entradas", fmt:"eur", total:"sum"}, {h:"Salidas", fmt:"eur", total:"sum"}, {h:"Debía haber", fmt:"eur"}, {h:"Contado", fmt:"eur"}, {h:"Diferencia", fmt:"eur", total:"sum", tone:true}, {h:"Espera media (min)", fmt:"num1", total:"avg"}
      ], d.cierres, {totals:true, headH:34});
      r = t.next;
      chart(sh, t, {type: d.cierres.length > 1 ? "line" : "col", title:"Ventas de cada día", cat:0, series:[{col:4, color:C.wine}, {col:5, color:C.gold}], fmt:'#,##0 "€"', labels: d.cierres.length <= 14, labelFmt:'#,##0 "€"', at:{c:0, r, w:10, h:16}});
      r += 17;
      note(sh, r, "Diferencia = contado − lo que debía haber en el cajón. El Bizum no entra en el cajón: ya está descontado de «Debía haber».", 14);
      sheets.push(sh);
    }
    // 3. Cuándo se vende
    {
      const sh = new Sheet("Cuándo se vende", [14, 10, 12, 4, 14, 12, 10, 12, 14]);
      let r = banner(sh, "CUÁNDO SE VENDE", sub, 9);
      r = section(sh, r, "Por hora del día", 9);
      const top = r;
      const t1 = table(sh, r, 0, "Horas", [{h:"Hora"}, {h:"Pedidos", fmt:"int", total:"sum"}, {h:"Ventas", fmt:"eur", total:"sum"}], d.horas, {totals:true});
      sh.set(4, top - 1, "Por día de la semana", sx({font:7, v:"bottom"}));
      const t2 = table(sh, top, 4, "Semana", [{h:"Día"}, {h:"Días abiertos", fmt:"int", total:"sum"}, {h:"Pedidos", fmt:"int", total:"sum"}, {h:"Ventas", fmt:"eur", total:"sum"}, {h:"Media por día", fmt:"eur"}], d.semana, {totals:true});
      r = Math.max(t1.next, t2.next);
      chart(sh, t1, {type:"col", title:"Ventas por hora", cat:0, series:[{col:2, color:C.wine}], fmt:'#,##0 "€"', labels:true, labelFmt:'#,##0 "€"', gap:40, at:{c:0, r, w:4, h:15}});
      chart(sh, t2, {type:"col", title:"Venta media por día de la semana", cat:0, series:[{col:4, color:C.gold}], fmt:'#,##0 "€"', labels:true, labelFmt:'#,##0 "€"', gap:40, at:{c:4, r, w:5, h:15}});
      sheets.push(sh);
    }
    // 4. Productos
    {
      const sh = new Sheet("Productos", [30, 13, 10, 12, 11, 11, 12, 13]);
      let r = banner(sh, "PRODUCTOS", sub, 8);
      r = section(sh, r, "Lo que más se vende", 8);
      const t = table(sh, r, 0, "Productos", [{h:"Producto"}, {h:"Grupo"}, {h:"Unidades", fmt:"int", total:"sum"}, {h:"Ventas", fmt:"eur", total:"sum"}, {h:"% de las ventas", fmt:"pct", total:"sum"}, {h:"Precio actual", fmt:"eur"}, {h:"Coste ingredientes", fmt:"eur"}, {h:"Margen por unidad", fmt:"eur", tone:true}], d.productos, {totals:true, headH:34});
      r = t.next;
      const topN = Math.min(10, d.productos.length);
      chart(sh, {...t, rEnd: t.r0 + topN, n: topN}, {type:"bar", title:"Top productos por ventas", cat:0, series:[{col:3, color:C.wine}], fmt:'#,##0 "€"', labels:true, labelFmt:'#,##0.00 "€"', gap:35, at:{c:0, r, w:4, h: Math.max(12, topN*1.6 + 4)}});
      chart(sh, {...t, rEnd: t.r0 + topN, n: topN}, {type:"doughnut", title:"Reparto de las ventas", cat:0, series:[{col:3}], labels:true, at:{c:4, r, w:4, h: Math.max(12, topN*1.6 + 4)}});
      r += Math.max(12, Math.ceil(topN*1.6 + 4)) + 1;
      note(sh, r, "Coste ingredientes = receta × coste por gramo/ml/unidad de los escandallos. Margen por unidad = precio − coste de ingredientes (sin contar fijos).", 8);
      sheets.push(sh);
    }
    // 5. Gastos
    {
      const sh = new Sheet("Gastos", [12, 34, 24, 12, 20, 3, 24, 12]);
      let r = banner(sh, "GASTOS", sub, 8);
      r = section(sh, r, "Todos los gastos apuntados", 8);
      const top = r;
      const t = table(sh, r, 0, "Gastos", [{h:"Fecha"}, {h:"Concepto"}, {h:"Categoría"}, {h:"Importe", fmt:"eur", total:"sum"}, {h:"Origen"}], d.gastos, {totals:true});
      sh.set(6, top - 1, "Por categoría", sx({font:7, v:"bottom"}));
      const tc = table(sh, top, 6, "Categorias", [{h:"Categoría"}, {h:"Importe", fmt:"eur", total:"sum"}], d.categorias, {totals:true});
      chart(sh, tc, {type:"doughnut", title:"En qué se va el dinero", cat:0, series:[{col:1}], labels:true, at:{c:6, r: tc.next, w:2, h:16}});
      r = Math.max(t.next, tc.next + 17);
      r = section(sh, r, "Gastos fijos al mes", 8);
      table(sh, r, 1, "Fijos", [{h:"Concepto"}, {h:"Categoría"}, {h:"Al mes", fmt:"eur", total:"sum"}], d.fijos, {totals:true});
      sheets.push(sh);
    }
    if (d.interno) addInterno(sheets, d.interno, sub);
    return pack(JSZip, sheets);
  }
  /* Hoja de varias tablas: secciones [{h:"Título", cols, rows, totals, note, chart}] */
  function listSheet(name, title, sub, widths, secs){
    const sh = new Sheet(name, widths);
    let r = banner(sh, title, sub, widths.length);
    secs.forEach(sc => {
      if (sc.intro){ r = note(sh, r, sc.intro, widths.length); r++; }
      if (sc.h) r = section(sh, r, sc.h, widths.length);
      const t = table(sh, r, 0, sc.tname || name, sc.cols, sc.rows, {totals: !!sc.totals, headH: 30});
      r = t.next;
      if (sc.chart && t.n){ chart(sh, t, {...sc.chart, at:{c:0, r, w: Math.min(widths.length, 8), h:15}}); r += 16; }
      if (sc.note){ r = note(sh, r, sc.note, widths.length); r++; }
    });
    return sh;
  }
  function addInterno(sheets, I, sub){
    const first = [
      listSheet("Sugerencias", "SUGERENCIAS DE LA APP", sub, [20, 44, 70, 12], [{intro:"Todo lo que la app os recomienda ahora mismo, reunido de Caja, Gastos, Datos y Suministros. «Aviso» = conviene actuar; «Idea» = para tenerlo en cuenta.", cols:[{h:"Apartado"}, {h:"Sugerencia"}, {h:"Detalle"}, {h:"Tipo"}], rows:I.sugerencias}]),
      listSheet("A tener en cuenta", "A TENER EN CUENTA PARA EL FUTURO", sub, [24, 60, 44], [{cols:[{h:"Tema"}, {h:"Qué vigilar"}, {h:"Cómo está ahora"}], rows:I.futuro}])
    ];
    sheets.unshift(...first);
    sheets.push(listSheet("Escandallos", "ESCANDALLOS", sub, [28, 14, 12, 9, 16, 16, 11], [{intro:"Coste por gramo, ml o unidad de cada ingrediente = lo pagado ÷ lo que trae. Es lo que usan las recetas para calcular el coste de cada bebida.", cols:[{h:"Ingrediente"}, {h:"Lo pagado", fmt:"eur"}, {h:"Contiene", fmt:"int"}, {h:"Unidad"}, {h:"Coste por unidad", fmt:"eur4"}, {h:"Coste anterior", fmt:"eur4"}, {h:"Cambio", fmt:"pct", tone:true}], rows:I.escandallos}]));
    sheets.push(listSheet("Coste por bebida", "COSTE POR BEBIDA", sub, [28, 12, 12, 10, 12, 16, 16], [{cols:[{h:"Producto"}, {h:"Os cuesta", fmt:"eur"}, {h:"Precio cliente", fmt:"eur"}, {h:"% coste", fmt:"pct"}, {h:"Margen", fmt:"eur", tone:true}, {h:"Si sube: mismo margen", fmt:"eur"}, {h:"Si sube: mismo %", fmt:"eur"}], rows:I.costes,
      chart:{type:"bar", title:"Lo que os cuesta cada bebida", cat:0, series:[{col:1, color:"5E0F1C"}, {col:2, color:"A8680F"}], fmt:'#,##0.00 "€"', gap:40},
      note:"Margen = precio − coste de ingredientes (sin fijos). Si sube un suministro: «mismo margen» = precio + lo que ha subido el coste; «mismo %» = precio × coste nuevo ÷ coste anterior."},
      {h:"Recetas", tname:"Recetas", cols:[{h:"Producto"}, {h:"Ingrediente"}, {h:"Cantidad", fmt:"num1"}, {h:"Unidad"}, {h:"Coste", fmt:"eur"}], rows:I.recetas}]));
    sheets.push(listSheet("Inventario", "INVENTARIO Y COMPRA", sub, [28, 9, 14, 14, 13, 12, 22, 13], [{cols:[{h:"Suministro"}, {h:"Unidad"}, {h:"Quedan", fmt:"num1"}, {h:"Gasto al día", fmt:"num1"}, {h:"Días que quedan", fmt:"num1"}, {h:"Mínimo", fmt:"num1"}, {h:"A comprar"}, {h:"Coste aprox.", fmt:"eur", total:"sum"}], rows:I.inventario, totals:true,
      note:"«Quedan» sale del último conteo menos lo vendido desde entonces. Si está vacío es que no se ha contado: contadlo en Ajustes → Suministros → Inventario. «A comprar» es para los días de la lista de la compra."},
      {h:"Últimas compras", tname:"Compras", cols:[{h:"Fecha"}, {h:"Suministro"}, {h:"Paquetes", fmt:"int"}, {h:"Cantidad", fmt:"num1"}, {h:"Coste", fmt:"eur", total:"sum"}], rows:I.compras, totals:true}]));
    if (I.box && I.box.length) sheets.push(listSheet("Gente del box", "GENTE DEL BOX", sub, [14, 10, 22, 12], [{cols:[{h:"Día"}, {h:"Hora"}, {h:"Clase"}, {h:"Personas", fmt:"int", total:"sum"}], rows:I.box, totals:true}]));
  }
  const api = {build};
  if (typeof module !== "undefined" && module.exports) module.exports = api; else window.UtopiaInforme = api;
})();
