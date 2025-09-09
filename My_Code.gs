/**** CONFIG ****/
const CFG = {
  YEAR: 2025,
  SC_ID: "",     // ID da planilha SC
  ES_ID: "",     // ID da planilha ES
  ES_TAB_CANDIDATES: ["Devoluções","Devolucoes","DEVOLUÇÕES","DEVOLUCOES"],
  HEADERS: ["Estoque","Data","Ticket","Pedido/NF","Produto","Analista","Motivo","Valor","Status"],
  // Regras de status a manter
  STATUS_KEEPERS: [
    s => s === "",                                 // em branco
    s => /cliente\s*informado/i.test(s),
    s => /card.*(criad|gerad)/i.test(s),
    s => /chargeback/i.test(s),
    s => /(devolu.*negad|garanti.*negad)/i.test(s),
    s => /finaliz/i.test(s),                       // <= adicionado
  ],
};

/**** MENU ****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Controle Devoluções")
    .addItem("Atualizar mês (aba atual)", "updateActiveMonth")
    .addItem("Atualizar todas (2025)", "updateAllMonths2025")
    .addSeparator()
    .addItem("Testar fontes (listar abas)", "debugListTabs")
    .addToUi();
}

/**** AÇÕES DO MENU ****/
function updateActiveMonth() {
  const dest = SpreadsheetApp.getActiveSheet();
  const monthIdx = monthIndexFromName(dest.getName());
  if (!monthIdx) {
    SpreadsheetApp.getUi().alert('Abra uma aba chamada exatamente um mês: "Janeiro", "Fevereiro", ..., "Dezembro".');
    return;
  }
  updateMonth(dest, CFG.YEAR, monthIdx);
}

function updateAllMonths2025() {
  const ss = SpreadsheetApp.getActive();
  const months = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
  months.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) updateMonth(sh, CFG.YEAR, i+1);
  });
}

/**** DEPURAÇÃO (opcional) ****/
function debugListTabs() {
  const sc = SpreadsheetApp.openById(CFG.SC_ID);
  const es = SpreadsheetApp.openById(CFG.ES_ID);
  const scNames = sc.getSheets().map(s=>s.getName()).join(", ");
  const esNames = es.getSheets().map(s=>s.getName()).join(", ");
  SpreadsheetApp.getUi().alert("SC abas: " + scNames + "\nES abas: " + esNames);
}

/**** PRINCIPAL: ATUALIZA UMA ABA ****/
function updateMonth(destSheet, year, monthIdx) {
  const {start, end} = monthStartEnd(year, monthIdx);

  // 1) SC
  const sc = SpreadsheetApp.openById(CFG.SC_ID);
  const scTab = findMonthSheet(sc, monthIdx);
  const scRowsRaw = scTab ? readSC(scTab) : [];
  const scRows = scRowsRaw.filter(r => keepStatus(r[8]));

  // 2) ES
  const es = SpreadsheetApp.openById(CFG.ES_ID);
  const esTab = findSheetByNames(es, CFG.ES_TAB_CANDIDATES);
  if (!esTab) {
    SpreadsheetApp.getUi().alert('Aba "Devoluções" (ou variação) não encontrada na planilha ES.');
    return;
  }
  const esRowsRaw = readES(esTab, start, end);
  const esRows = esRowsRaw.filter(r => keepStatus(r[8]));

  // 3) Escreve
  const rows = scRows.concat(esRows);
  writeTable(destSheet, rows);

  // 4) Diagnóstico rápido
  const scStatuses = uniqueCounts(scRowsRaw.map(r => normStr(r[8])));
  const esStatuses = uniqueCounts(esRowsRaw.map(r => normStr(r[8])));
  SpreadsheetApp.getUi().alert(
    `Diagnóstico ${indexToPtBrMonth(monthIdx)}/${year}\n\n` +
    `SC: lidas ${scRowsRaw.length}, após filtro ${scRows.length}\n` +
    `ES: lidas ${esRowsRaw.length}, após filtro ${esRows.length}\n\n` +
    `Alguns STATUS SC:\n${previewCounts(scStatuses)}\n\n` +
    `Alguns STATUS ES:\n${previewCounts(esStatuses)}`
  );
}
function normStr(s){ return (s||"").toString().trim(); }
function uniqueCounts(arr){
  const m = new Map();
  arr.forEach(v => m.set(v, (m.get(v)||0)+1));
  // ordena por contagem desc e pega top 10
  return Array.from(m.entries()).sort((a,b)=>b[1]-a[1]).slice(0,10);
}
function previewCounts(list){
  if (!list.length) return "(vazio)";
  return list.map(([k,v]) => `${k||"(em branco)"} — ${v}`).join("\n");
}


/**** LEITURAS E MAPEAMENTOS ****/
// SC colunas: A Data, B Ticket, C Pedido, D ID, E Produto, F SN, G OPENBOX, H Analista, I Motivo, J Valor, K Status
function readSC(scTab) {
  const lastRow = scTab.getLastRow();
  if (lastRow < 2) return [];
  const data = scTab.getRange(2,1,lastRow-1,11).getValues(); // A2:K
  return data
    .filter(r => hasAny(r[0], r[1], r[2])) // ignora linhas totalmente vazias
    .map(r => ([
      "SC",
      asDate(r[0]),
      safe(r[1]),
      safe(r[2]),
      safe(r[4]),
      safe(r[7]),
      safe(r[8]),
      asNumber(r[9]),
      safe(r[10])
    ]));
}

// ES colunas: A Data, B Ticket, C NF SAÍDA, D EAN, E Produto, F Situação, G SN, H NF-D, I Analista, J Motivo, K Loja, L Valor, M Status
function readES(esTab, start, end) {
  const lastRow = esTab.getLastRow();
  if (lastRow < 2) return [];
  const data = esTab.getRange(2,1,lastRow-1,13).getValues(); // A2:M
  return data
    .filter(r => inRange(parseAnyDate(r[0]), start, end)) // Data em A
    .map(r => ([
      "ES",
      asDate(parseAnyDate(r[0])),
      safe(r[1]),
      safe(r[7]),   // NF-D -> Pedido/ND
      safe(r[4]),
      safe(r[8]),
      safe(r[9]),
      asNumber(r[11]),
      safe(r[12])
    ]));
}

/**** ESCRITA ****/
function writeTable(sheet, rows) {
  const headers = CFG.HEADERS;
  // garante cabeçalho
  sheet.getRange(1,1,1,headers.length).setValues([headers]);

  // detecta até onde havia dados antes (apenas em A:I)
  const lastRow = sheet.getLastRow();          // última linha usada na aba
  const lastDataRow = Math.max(2, lastRow);    // dados começam na linha 2

  // limpa apenas o bloco existente de dados em A2:I (conteúdo, não formatação)
  if (lastRow >= 2) {
    sheet.getRange(2,1,lastRow-1, headers.length).clearContent();
  }

  // escreve os dados novos (ou uma linha vazia, se não houver)
  const out = rows.length ? rows : [["","","","","","","","",""]];
  sheet.getRange(2,1,out.length, headers.length).setValues(out);

  // formata somente tipo de dados (não mexe em borda/cor)
  sheet.getRange("B2:B" + (out.length+1)).setNumberFormat("dd/mm/yyyy");
  sheet.getRange("H2:H" + (out.length+1)).setNumberFormat("#,##0.00");
}


/**** STATUS ****/
function keepStatus(statusCell) {
  const s = (statusCell || "").toString().trim();
  return CFG.STATUS_KEEPERS.some(fn => fn(s));
}

/**** HELPERS ****/
function monthIndexFromName(name) {
  const n = normalize(name);
  const map = {janeiro:1,fevereiro:2,marco:3,março:3,abril:4,maio:5,junho:6,julho:7,agosto:8,setembro:9,outubro:10,novembro:11,dezembro:12};
  return map[n] || 0;
}
function monthStartEnd(y, m){
  const start = new Date(y, m-1, 1);
  const end = new Date(y, m, 1);
  return {start, end};
}
function normalize(s){
  return (s||"").toString().toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu,"").trim();
}
function findMonthSheet(ss, monthIdx){
  const target = normalize(indexToPtBrMonth(monthIdx));
  return ss.getSheets().find(sh => normalize(sh.getName()) === target) || null;
}
function indexToPtBrMonth(i){
  return ["","Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"][i] || "";
}
function findSheetByNames(ss, names){
  const all = ss.getSheets();
  for (const n of names){
    const found = all.find(sh => normalize(sh.getName()) === normalize(n));
    if (found) return found;
  }
  return null;
}
function hasAny(){ return Array.from(arguments).some(v => v !== "" && v !== null && v !== undefined); }
function safe(v){ return v == null ? "" : v; }
function asNumber(v){
  if (typeof v === "number") return v;
  const t = (v||"").toString().replace(/\./g,"").replace(",",".").trim();
  const n = parseFloat(t);
  return isNaN(n) ? 0 : n;
}
function asDate(v){
  if (v instanceof Date) return v;
  if (!v) return "";
  const d = new Date(v);
  return isNaN(d) ? "" : d;
}
function parseAnyDate(v){
  if (v instanceof Date) return v;
  const t = (v||"").toString().trim();
  if (!t) return null;
  // tenta dd/mm/aaaa
  const m = t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m){
    const d = new Date(parseInt(m[3],10), parseInt(m[2],10)-1, parseInt(m[1],10));
    return isNaN(d) ? null : d;
  }
  const d = new Date(t);
  return isNaN(d) ? null : d;
}
function inRange(d, start, end){
  return d instanceof Date && d >= start && d < end;
}

/**** ===== RELATÓRIO: CLIENTE INFORMADO ===== ****/

// Cabeçalhos desta visão (10 colunas)
// Cabeçalhos desta visão (AGORA 4 colunas)
const HEADERS_CI = ["Estoque","Data","Ticket","Pedido/NF"];


// Atalho no menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Controle Devoluções")
    .addItem("Atualizar mês (aba atual)", "updateActiveMonth")
    .addItem("Atualizar todas (2025)", "updateAllMonths2025")
    .addSeparator()
    .addItem("Relatório: Cliente Informado", "updateClienteInformado")        // << NOVO
    .addItem("Testar fontes (listar abas)", "debugListTabs")
    .addToUi();
}

// Gera/atualiza a aba "Cliente Informado" com SC+ES de todos os meses do ano
function updateClienteInformado() {
  const ss = SpreadsheetApp.getActive();
  const dest = ss.getSheetByName("Cliente Informado") || ss.insertSheet("Cliente Informado");

  // SC (todas as abas mensais 1..12), só cliente informado
  const sc = SpreadsheetApp.openById(CFG.SC_ID);
  const rowsSC = [];
  for (let i = 1; i <= 12; i++) {
    const scTab = findMonthSheet(sc, i);
    if (!scTab) continue;
    const data = readSC(scTab)                 // retorna 9 colunas: Estoque..Status
      .filter(r => isClienteInformado(r[8]))   // r[8] = Status
      .map(r => [r[0], r[1], r[2], r[3]]);     // << pega só: Estoque, Data, Ticket, Pedido/ND
    rowsSC.push(...data);
  }

  // ES (aba “Devoluções”), quebrado por mês do ano CFG.YEAR
  const es = SpreadsheetApp.openById(CFG.ES_ID);
  const esTab = findSheetByNames(es, CFG.ES_TAB_CANDIDATES);
  if (!esTab) {
    SpreadsheetApp.getUi().alert('Aba "Devoluções" (ou variação) não encontrada na planilha ES.');
    return;
  }
  const rowsES = [];
  for (let i = 1; i <= 12; i++) {
    const {start, end} = monthStartEnd(CFG.YEAR, i);
    const chunk = readES(esTab, start, end)
      .filter(r => isClienteInformado(r[8]))
      .map(r => [r[0], r[1], r[2], r[3]]);     // << mesmas 4 colunas
    rowsES.push(...chunk);
  }

  // Junta, ordena por Data (coluna 2), escreve A1:D
  const all = rowsSC.concat(rowsES).sort((a,b) => {
    const da = a[1] instanceof Date ? a[1].getTime() : -Infinity;
    const db = b[1] instanceof Date ? b[1].getTime() : -Infinity;
    return da - db;
  });

  writeTableCI(dest, all);
  SpreadsheetApp.getUi().alert(`Relatório "Cliente Informado" (4 colunas) atualizado.\nLinhas: ${all.length}`);
}



// Apenas "cliente informado"
function isClienteInformado(statusCell) {
  const s = (statusCell || "").toString().trim();
  return /cliente\s*informado/i.test(s);
}

// Escreve na aba "Cliente Informado" preservando formatação
function writeTableCI(sheet, rows) {
  // Cabeçalho A1:D1
  sheet.getRange(1,1,1,HEADERS_CI.length).setValues([HEADERS_CI]);

  // Limpa só conteúdo anterior A2:D (sem mexer em formatação/bordas fora)
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2,1,lastRow-1, HEADERS_CI.length).clearContent();
  }

  // Garante ao menos 1 linha
  const out = rows.length ? rows : [["","","",""]];
  sheet.getRange(2,1,out.length, HEADERS_CI.length).setValues(out);

  // Formato de Data (coluna B)
  const endRow = 1 + out.length;
  sheet.getRange(`B2:B${endRow}`).setNumberFormat("dd/mm/yyyy");
}


